"""点名：每个文件是谁的、能不能排版、不能的话为什么。只读。"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Callable, Sequence

from .converter import ConversionResult, convert_doc_files
from .decisions import Decisions
from .identity import Identity, Source, resolve_identity, teacher_identity
from .reader import DocContent, DocxReadError, read_docx
from .roster import Roster
from .scanner import FileKind, SourceFile, scan_folder, subfolders
from .structure import Structure, analyze

DIFF_CONTEXT = 8
MAX_DIFFS_SHOWN = 3


class Outcome(Enum):
    DONE = "已完成"
    # 点名时的中间状态：排版之后会变成"已完成"或"需要您确认"
    READY = "可以排版"
    REVIEW = "需要您确认"
    UNPROCESSED = "未处理"
    IGNORED = "已忽略"


@dataclass
class Record:
    source: SourceFile
    outcome: Outcome
    identity: Identity = field(default_factory=Identity)
    notes: list[str] = field(default_factory=list)
    # 排版时除了统一格式以外还动了哪些地方
    changes: list[str] = field(default_factory=list)
    # 质检发现的问题；空列表是通过，None 是没有做质检
    quality: list[str] | None = None
    # 排好的文件实际写出去时叫什么；没写出去就是 None
    output_name: str | None = None
    content: DocContent | None = None
    structure: Structure | None = None
    # 排版时要用的 .docx 内容；旧版 .doc 是转换后的结果
    docx_bytes: bytes | None = None

    @property
    def name(self) -> str | None:
        return self.identity.name

    @property
    def student_id(self) -> str | None:
        return self.identity.student_id

    @property
    def name_source(self) -> Source | None:
        return self.identity.name_source

    @property
    def id_source(self) -> Source | None:
        return self.identity.id_source

    @property
    def new_filename(self) -> str | None:
        # 不会生成文件的（PDF、空白、临时文件），就没有新文件名
        will_be_written = self.outcome in (Outcome.DONE, Outcome.READY, Outcome.REVIEW)
        if will_be_written and self.name and self.student_id:
            return f"{self.student_id}-{self.name}.docx"
        return None

    @property
    def identity_certain(self) -> bool:
        """姓名学号都有，而且几处写的互相对得上、和名单也对得上。"""
        return bool(self.name and self.student_id) and not self.identity.problems

    @property
    def has_content(self) -> bool:
        return self.content is not None and not self.content.is_blank


@dataclass
class Inventory:
    records: list[Record]
    skipped_folders: list[str] = field(default_factory=list)
    missing_students: list[tuple[str, str]] = field(default_factory=list)
    # 决定表里写了、但作业文件夹里没有的文件名（多半是打错了）
    unmatched_decisions: list[str] = field(default_factory=list)
    # 学号 -> 可能是这位学生交的、但身份还没确认的文件
    possible_files: dict[str, list[str]] = field(default_factory=dict)

    def count(self, outcome: Outcome) -> int:
        return sum(1 for r in self.records if r.outcome is outcome)


def take_inventory(
    input_dir: Path,
    roster: Roster | None,
    convert: Callable[[Sequence[Path]], ConversionResult] = convert_doc_files,
    decisions: Decisions | None = None,
) -> Inventory:
    decisions = decisions or Decisions()
    files = scan_folder(input_dir)
    wanted = [f for f in files if f.path.name not in decisions.discarded]
    conversion = convert([f.path for f in wanted if f.kind is FileKind.DOC])

    records = [_examine_safely(source, roster, conversion, decisions) for source in files]
    _flag_duplicates(records)

    missing, possible = [], {}
    if roster is not None:
        # 只有身份确定的才算"交了"。张三把学号写成了李四的：不能因此算李四交了、张三没交
        active = [r for r in records if r.outcome is not Outcome.IGNORED]
        seen = {r.student_id for r in active if r.identity_certain}
        missing = [(i, n) for i, n in roster.entries if i not in seen]
        for record in active:
            if record.identity_certain:
                continue
            candidates = {record.student_id, roster.id_of(record.name) if record.name else None}
            for student_id in sorted(c for c in candidates if c and c not in seen):
                possible.setdefault(student_id, []).append(record.source.path.name)
    unmatched = sorted(decisions.filenames() - {f.path.name for f in files})
    return Inventory(records, subfolders(input_dir), missing, unmatched, possible)


def _examine_safely(
    source: SourceFile, roster: Roster | None, conversion: ConversionResult, decisions: Decisions
) -> Record:
    """一份文件出任何意外，都只影响它自己：记成未处理、写明原因，其余的照常进行。"""
    try:
        return _examine(source, roster, conversion, decisions)
    except Exception as error:
        note = f"读这份文件时出错，没有处理（{type(error).__name__}：{error}）"
        return Record(source, Outcome.UNPROCESSED, notes=[note])


def _examine(
    source: SourceFile, roster: Roster | None, conversion: ConversionResult, decisions: Decisions
) -> Record:
    filename = source.path.name
    if source.kind is FileKind.TEMP:
        return Record(source, Outcome.IGNORED, notes=["Word 打开文件时留下的临时文件，不是作业"])
    if source.error:
        return Record(source, Outcome.UNPROCESSED, notes=[f"这个文件{source.error}"])
    if filename in decisions.discarded:
        remark = decisions.discarded[filename]
        note = "老师决定不要这份" + (f"：{remark}" if remark else "")
        return Record(source, Outcome.IGNORED, notes=[note])

    from_filename = resolve_identity(filename, [], [], roster)
    # 没法读正文的文件，文件名里认出的人只用来告诉老师该找谁，对不对得上不在这里追究
    from_filename.problems.clear()

    if source.kind is FileKind.PDF:
        return Record(source, Outcome.UNPROCESSED, from_filename, ["交的是 PDF，请学生重交 Word 版"])
    if source.kind is FileKind.OTHER:
        return Record(
            source, Outcome.UNPROCESSED, from_filename, ["不是 Word 文件，或者文件已损坏"]
        )

    if source.kind is FileKind.DOC:
        docx_bytes = conversion.converted.get(source.path)
        if docx_bytes is None:
            reason = conversion.failed.get(source.path, "旧版 .doc 没有被转换")
            return Record(source, Outcome.UNPROCESSED, from_filename, [reason])
    else:
        docx_bytes = source.read_bytes()

    try:
        content = read_docx(docx_bytes)
    except DocxReadError as error:
        return Record(source, Outcome.UNPROCESSED, from_filename, [str(error)])

    if filename in decisions.assigned:
        identity = teacher_identity(*decisions.assigned[filename], roster)
    else:
        identity = resolve_identity(filename, content.paragraphs, content.header_lines, roster)
    record = Record(source, Outcome.READY, identity, content=content, docx_bytes=docx_bytes)
    if content.is_blank:
        record.outcome = Outcome.UNPROCESSED
        record.notes.append("空白文档，一个字都没有")
        identity.problems.clear()
        return record

    record.structure = analyze(docx_bytes)
    missing = [label for label, value in (("姓名", identity.name), ("学号", identity.student_id)) if not value]
    if missing:
        record.notes.append(f"找不到{'和'.join(missing)}")
    record.notes.extend(identity.problems)
    if record.structure.title is None:
        record.notes.append("没有找到题目：开头第一行不是题目")
    if record.notes:
        record.outcome = Outcome.REVIEW
    return record


def _flag_duplicates(records: list[Record]) -> None:
    by_id: dict[str, list[Record]] = {}
    for record in records:
        # 读不了的那份（PDF、转换失败的 .doc）也算一份：它说不定才是学生想交的
        if record.student_id and record.outcome is not Outcome.IGNORED:
            by_id.setdefault(record.student_id, []).append(record)
    for group in by_id.values():
        if len(group) < 2:
            continue
        for record in group:
            if not record.has_content:
                continue
            record.outcome = Outcome.REVIEW
            for other in group:
                if other is not record:
                    record.notes.append(_describe_duplicate(record, other, len(group)))


def _describe_duplicate(record: Record, other: Record, group_size: int) -> str:
    other_name = other.source.path.name
    head = f"同一个学号交了 {group_size} 份，另一份是「{other_name}」。"
    if not other.has_content:
        return head + "那一份程序读不了，请确认以哪份为准。"
    mine, theirs = _text_paragraphs(record), _text_paragraphs(other)
    if mine == theirs:
        return head + "两份的文字完全相同。"
    if len(mine) != len(theirs):
        return head + f"两份的文字不同：这份 {len(mine)} 段，另一份 {len(theirs)} 段。"
    diffs = [
        f"这份写「{a}」，另一份写「{b}」"
        for a, b in (_differing_part(m, t) for m, t in zip(mine, theirs) if m != t)
    ]
    shown = "；".join(diffs[:MAX_DIFFS_SHOWN])
    more = f"；另有 {len(diffs) - MAX_DIFFS_SHOWN} 处" if len(diffs) > MAX_DIFFS_SHOWN else ""
    return head + f"两份的文字有 {len(diffs)} 段不同：{shown}{more}。"


def _text_paragraphs(record: Record) -> list[str]:
    return [p.strip() for p in record.content.paragraphs if p.strip()]


def _differing_part(a: str, b: str) -> tuple[str, str]:
    """两段文字里不一样的那一小块，各带前后几个字，方便老师一眼看出差别。"""
    prefix = 0
    while prefix < min(len(a), len(b)) and a[prefix] == b[prefix]:
        prefix += 1
    suffix = 0
    while suffix < min(len(a), len(b)) - prefix and a[-1 - suffix] == b[-1 - suffix]:
        suffix += 1
    start = max(0, prefix - DIFF_CONTEXT)
    # 不要把一个数字或英文单词从中间截断（"18 天"截成"8 天"会误导人）
    while 0 < start < len(a) and _is_word_char(a[start - 1]) and _is_word_char(a[start]):
        start -= 1

    def excerpt(text: str) -> str:
        end = min(len(text), len(text) - suffix + DIFF_CONTEXT)
        while 0 < end < len(text) and _is_word_char(text[end - 1]) and _is_word_char(text[end]):
            end += 1
        return ("…" if start > 0 else "") + text[start:end] + ("…" if end < len(text) else "")

    return excerpt(a), excerpt(b)


def _is_word_char(char: str) -> bool:
    return char.isascii() and char.isalnum()
