"""流水线：点名 → 排版 → 比对内容 → 质检 → 写文件 → 出报告。"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Sequence

from .checker import check_document, check_filename
from .converter import ConversionResult, convert_doc_files
from .decisions import Decisions
from .formatter import format_document
from .integrity import verify_content
from .inventory import Inventory, Outcome, Record, take_inventory
from .report import write_report
from .roster import Roster
from .rules import Rules
from .scanner import fingerprint
from .structure import Structure

DONE_FOLDER = "已完成"
REVIEW_FOLDER = "需人工确认"
REPORT_NAME = "处理报告.xlsx"
# Windows 不许出现在文件名里的字符
INVALID_FILENAME_CHARACTERS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
RESERVED_NAMES = {"con", "prn", "aux", "nul", *(f"com{i}" for i in range(1, 10)), *(f"lpt{i}" for i in range(1, 10))}
# 文件名太长在有的电脑上写不出来；留出文件夹路径的余量
MAX_STEM_LENGTH = 100


class PipelineError(Exception):
    """没法开始运行。信息是写给老师看的。"""


@dataclass
class RunResult:
    inventory: Inventory
    run_dir: Path
    report_path: Path
    originals_untouched: bool
    file_count: int
    # 中途出了意外、没有全部做完时的原因；正常是 None
    incomplete: str | None = None


def _new_run_dir(output_root: Path) -> Path:
    """每次运行一个新的子文件夹，从不覆盖以前的结果。"""
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    run_dir, attempt = output_root / stamp, 1
    while run_dir.exists():
        attempt += 1
        run_dir = output_root / f"{stamp}_{attempt}"
    run_dir.mkdir(parents=True)
    return run_dir


def run(
    input_dir: Path,
    output_root: Path,
    rules: Rules,
    roster: Roster | None = None,
    decisions: Decisions | None = None,
    convert: Callable[[Sequence[Path]], ConversionResult] = convert_doc_files,
) -> RunResult:
    input_dir, output_root = Path(input_dir).resolve(), Path(output_root).resolve()
    if output_root == input_dir or input_dir in output_root.parents:
        raise PipelineError("结果文件夹不能放在作业文件夹里面，否则会和原件混在一起。")

    before = fingerprint(input_dir)
    inventory = take_inventory(input_dir, roster, convert=convert, decisions=decisions)
    run_dir = _new_run_dir(output_root)

    # 排版和写文件万一出了没料到的错，报告也一定要写出来，并且写明"这次没有做完"：
    # 不能留下半截结果又没有报告，让老师以为那就是全部
    incomplete = None
    try:
        formatted: list[tuple[Record, bytes]] = []
        for record in inventory.records:
            if record.outcome in (Outcome.READY, Outcome.REVIEW) and record.has_content:
                data = _format_one(record, rules)
                if data is not None:
                    formatted.append((record, data))
        _write_outputs(formatted, run_dir)
    except Exception as error:
        incomplete = f"{type(error).__name__}：{error}"

    untouched = fingerprint(input_dir) == before
    report_path = write_report(
        run_dir / REPORT_NAME,
        inventory,
        input_dir=input_dir,
        roster=roster,
        originals_untouched=untouched,
        file_count=len(before),
        decisions=decisions,
        incomplete=incomplete,
    )
    return RunResult(inventory, run_dir, report_path, untouched, len(before), incomplete)


def _format_one(record: Record, rules: Rules) -> bytes | None:
    """排一份作业并比对内容。返回排好的文件内容；不该写出文件时返回 None。"""
    # 拿不准是谁的作业时，格式照排，姓名行不动
    name, student_id = (record.name, record.student_id) if record.identity_certain else (None, None)
    info_line = rules.info_line(name, student_id) if record.identity_certain else None
    try:
        result = format_document(record.docx_bytes, name, student_id, rules)
        if not result.info_rewritten:
            info_line = None  # 姓名行没有改写：比对时它也必须原样不变
        problems = verify_content(record.docx_bytes, result.docx_bytes, info_line, name, student_id)
    except Exception as error:  # 一份文件出问题不能连累其他文件
        record.outcome = Outcome.UNPROCESSED
        record.notes.append(f"排版时出错，这份没有处理（{type(error).__name__}：{error}）")
        return None

    if problems:
        # 底线：内容对不上的不算完成，哪个文件夹都不放
        record.outcome = Outcome.REVIEW
        record.notes.extend(problems)
        record.notes.append("因为内容对不上，没有生成排好的文件，请直接看原件")
        return None
    record.structure = Structure(result.title, result.headings)
    record.changes = result.changes
    if result.warnings:
        # 程序拿不准、没有动的地方，要老师看一眼；这种文件不下质检结论
        record.outcome = Outcome.REVIEW
        record.notes.extend(result.warnings)
        return result.docx_bytes

    try:
        quality = check_document(result.docx_bytes, rules, info_line, expect_title=result.title is not None)
        if info_line is not None:
            quality += check_filename(safe_filename(record.new_filename), info_line, rules)
    except Exception as error:
        record.outcome = Outcome.REVIEW
        record.notes.append(f"质检时出错，这份没有质检结论，请打开看一下（{type(error).__name__}：{error}）")
        return result.docx_bytes
    record.quality = quality
    if quality:
        record.outcome = Outcome.REVIEW
        record.notes.append("质检没通过：" + "；".join(quality))
    return result.docx_bytes


def safe_filename(name: str) -> str:
    """把名字变成在 Windows 上一定写得出来的文件名：去掉非法字符，太长就截短。"""
    # 不能先交给 Path 去拆：名字里要是有斜杠，会被当成文件夹分隔符，只剩最后一截
    stem = name[: -len(".docx")] if name.lower().endswith(".docx") else name
    stem = INVALID_FILENAME_CHARACTERS.sub("_", stem).strip(" .")
    if len(stem) > MAX_STEM_LENGTH:
        stem = stem[: MAX_STEM_LENGTH - 1] + "…"
    if not stem or stem.lower() in RESERVED_NAMES:
        stem = "_" + stem
    return stem + ".docx"


def _write_new(folder: Path, name: str, data: bytes) -> str:
    """写一个新文件，绝不覆盖已有的：重名就在后面编号。返回实际用的文件名。"""
    stem, attempt = Path(name).stem, 1
    while True:
        candidate = name if attempt == 1 else f"{stem}_{attempt}.docx"
        try:
            with open(folder / candidate, "xb") as handle:
                handle.write(data)
            return candidate
        except FileExistsError:
            attempt += 1


def _save(record: Record, folder: Path, name: str, data: bytes) -> bool:
    """写一份排好的作业。写不出来（磁盘满、路径太长……）只影响这一份。"""
    try:
        record.output_name = _write_new(folder, name, data)
        return True
    except OSError as error:
        record.outcome = Outcome.UNPROCESSED
        record.notes.append(f"排好了但写文件时出错，没有生成文件（{type(error).__name__}：{error}）")
        return False


def _write_outputs(formatted: list[tuple[Record, bytes]], run_dir: Path) -> None:
    done_dir = run_dir / DONE_FOLDER
    done_dir.mkdir()
    review: list[tuple[Record, bytes]] = []
    for record, data in formatted:
        if record.outcome is not Outcome.READY:
            review.append((record, data))
            continue
        name = safe_filename(record.new_filename)
        if (done_dir / name).exists():
            record.outcome = Outcome.REVIEW
            record.notes.append(f"「{name}」已经有了，另一份作业用了同样的姓名学号")
            review.append((record, data))
        elif _save(record, done_dir, name, data):
            record.outcome = Outcome.DONE

    if not review:
        return
    review_dir = run_dir / REVIEW_FOLDER
    review_dir.mkdir()
    names = _review_names([record for record, _ in review])
    for (record, data), name in zip(review, names):
        _save(record, review_dir, name, data)


def _review_names(records: list[Record]) -> list[str]:
    """「需人工确认」里的文件名：知道是谁的就用新名字，否则保留原名；重名的带上原文件名以示区别。"""
    originals = [record.source.path.stem for record in records]
    names = [
        record.new_filename if record.identity_certain else f"{stem}.docx"
        for record, stem in zip(records, originals)
    ]
    counts = Counter(name.lower() for name in names)
    return [
        safe_filename(f"{Path(name).stem}（{stem}）.docx" if counts[name.lower()] > 1 else name)
        for name, stem in zip(names, originals)
    ]
