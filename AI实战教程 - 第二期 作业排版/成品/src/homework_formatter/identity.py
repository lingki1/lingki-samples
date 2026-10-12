"""从正文、页眉、文件名、名单里找出一份作业是谁的。找不到就留空，不猜。"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path

from .roster import Roster

ID_DIGITS = 8
INFO_LINE_MAX_LENGTH = 40
# 姓名学号行只在开头几段里找，免得把正文里的数字当成学号
SEARCH_DEPTH = 5

ID_PATTERN = re.compile(rf"(?<!\d)\d{{{ID_DIGITS}}}(?!\d)")
CLASS_PATTERN = re.compile(
    r"[高初][一二三1-3]\s*(?:年级)?\s*[（(]?\s*(?:\d{1,2}|[一二三四五六七八九十]{1,3})\s*[）)]?\s*班"
)
# 标签必须在一段话的开头，或者前面是空格、逗号之类：
# "调查的班级：高二3班和高二5班"里的"班级："不是标签
_TOKEN_START = r"(?:^|(?<=[\s，,、;；|/]))"
LABEL_PATTERN = re.compile(_TOKEN_START + r"(班级|姓名|学号)\s*[:：]?")
LABEL_WITH_COLON = re.compile(_TOKEN_START + r"(班级|姓名|学号)\s*[:：]")
ID_LABEL = re.compile(_TOKEN_START + r"学号\s*[:：]?\s*(\d{8})(?!\d)")
NAME_AFTER_LABEL = re.compile(_TOKEN_START + r"姓名\s*[:：]?\s*([一-鿿]{2,4})(?![一-鿿])")
CHINESE_RUN = re.compile(r"[一-鿿]+")
SEPARATORS = re.compile(r"[\s，,、;；:：|/\\\-—_·（）()【】\[\]]+")
NAME_LENGTHS = range(2, 5)
# 只写「高二（3）班 张三」、没有学号也没有标签时，姓名最多认 3 个字：
# 4 个字的更可能是以班级开头的题目（「高二（3）班 睡眠调查」）
BARE_NAME = re.compile(r"[一-鿿]{2,3}")

# 文件名里常见的、不属于姓名的词；长的排前面，先删掉
FILENAME_NOISE = ("研究性学习", "小报告", "最终版", "报告", "作业", "终版", "新建", "文档")
FILENAME_NAME_SUFFIXES = ("的", "改")

CLEAN = "clean"  # 只有班级、姓名、学号
MESSY = "messy"  # 像姓名行，但夹着别的内容


class Source(Enum):
    BODY = "正文"
    HEADER = "页眉"
    FILENAME = "文件名"
    ROSTER = "名单"
    TEACHER = "老师指定"


@dataclass
class Identity:
    name: str | None = None
    student_id: str | None = None
    name_source: Source | None = None
    id_source: Source | None = None
    problems: list[str] = field(default_factory=list)


def _without_known_parts(text: str) -> str:
    """去掉学号、班级、「班级/姓名/学号」标签之后剩下的部分。"""
    return LABEL_PATTERN.sub(" ", CLASS_PATTERN.sub(" ", ID_PATTERN.sub(" ", text)))


def _tokens(text: str) -> list[str]:
    return [t for t in SEPARATORS.split(text) if t]


def info_line_kind(text: str) -> str | None:
    """这一行是不是姓名行。

    CLEAN：只有班级、姓名、学号，可以放心改写成统一写法。
    MESSY：有学号或「姓名：」这样的标签，但还夹着别的内容（指导老师、日期……），不能整行改写。
    None：不是姓名行。
    """
    text = text.strip()
    if not text or len(text) > INFO_LINE_MAX_LENGTH or "\n" in text:
        return None
    has_id = ID_PATTERN.search(text) is not None
    has_label = LABEL_WITH_COLON.search(text) is not None
    rest = _without_known_parts(text)
    tokens = _tokens(rest)
    # 去掉认识的标签后还有冒号，说明有别的"某某：某某"
    other_label = re.search(r"[:：]", rest) is not None
    only_a_name = len(tokens) == 1 and re.fullmatch(r"[一-鿿]{2,4}", tokens[0]) is not None
    tidy = (not tokens or only_a_name) and not other_label

    if has_id or has_label:
        return CLEAN if tidy else MESSY
    # 只写了「高二（3）班 张三」：班级和姓名之间要有空格，而且后面只剩一个短名字
    class_match = CLASS_PATTERN.search(text)
    if class_match and len(tokens) == 1 and BARE_NAME.fullmatch(tokens[0]):
        if text[class_match.end() : class_match.end() + 1].isspace():
            return CLEAN
    return None


def is_info_line(text: str) -> bool:
    return info_line_kind(text) == CLEAN


def parse_info_line(text: str, roster: Roster | None = None) -> tuple[str | None, str | None]:
    """从一行「班级 姓名 学号」里取出（姓名, 学号），各种写法都认。"""
    id_match = ID_PATTERN.search(text)
    student_id = id_match.group() if id_match else None
    rest = _without_known_parts(text)
    if roster is not None:
        # 名单上的名字要单独成词才算：「张伟明」里的「张伟」不算
        exact = [t for t in dict.fromkeys(CHINESE_RUN.findall(rest)) if roster.has_name(t)]
        if len(exact) == 1:
            return exact[0], student_id
    labelled = NAME_AFTER_LABEL.search(text)
    if labelled:
        return labelled.group(1), student_id
    candidates = [t for t in CHINESE_RUN.findall(rest) if len(t) in NAME_LENGTHS]
    return (candidates[0] if len(candidates) == 1 else None), student_id


def identity_leftover(text: str, name: str | None, student_id: str | None) -> str:
    """一行里去掉班级、标签、这位学生的姓名和学号之后，还剩什么。空字符串表示什么都不剩。"""
    rest = text
    if student_id:
        rest = rest.replace(student_id, " ")
    rest = LABEL_PATTERN.sub(" ", CLASS_PATTERN.sub(" ", rest))
    if name:
        rest = rest.replace(name, " ")
    return " ".join(_tokens(rest))


def _filename_tokens(stem: str) -> list[str]:
    rest = CLASS_PATTERN.sub(" ", stem)
    for word in FILENAME_NOISE:
        rest = rest.replace(word, " ")
    tokens = []
    for token in CHINESE_RUN.findall(rest):
        while token.endswith(FILENAME_NAME_SUFFIXES):
            token = token[:-1]
        if token:
            tokens.append(token)
    return tokens


def name_from_filename(filename: str, roster: Roster | None = None) -> str | None:
    stem = Path(filename).stem
    cleaned = _filename_tokens(stem)
    if roster is not None:
        # 名单上的名字要单独成词才算：《我的童年趣事调查》里的「童年」不算
        tokens = dict.fromkeys(CHINESE_RUN.findall(stem) + cleaned)
        names = [t for t in tokens if roster.has_name(t)]
        return names[0] if len(names) == 1 else None
    names = [t for t in cleaned if len(t) in NAME_LENGTHS]
    return names[0] if len(names) == 1 else None


@dataclass
class _Found:
    name: str | None = None
    student_id: str | None = None
    names: list[str] = field(default_factory=list)
    ids: list[str] = field(default_factory=list)


def _from_lines(lines: list[str], roster: Roster | None) -> _Found:
    found = _Found()
    for line in lines:
        kind = info_line_kind(line)
        if kind is None:
            continue
        name, student_id = parse_info_line(line, roster)
        if kind == MESSY:
            # 夹着别的内容的行：数字只有带「学号」标签、或者名单上确有其号，才当学号
            labelled = ID_LABEL.search(line)
            ids = [i for i in ID_PATTERN.findall(line) if roster is not None and roster.name_of(i)]
            student_id = labelled.group(1) if labelled else (ids[0] if ids else None)
            line_ids = [student_id] if student_id else []
            if not (NAME_AFTER_LABEL.search(line) or (roster is not None and name and roster.has_name(name))):
                name = None
        else:
            line_ids = ID_PATTERN.findall(line)
        found.names += [name] if name and name not in found.names else []
        found.ids += [i for i in line_ids if i not in found.ids]
    found.name = found.names[0] if found.names else None
    found.student_id = found.ids[0] if found.ids else None
    return found


def _pick(found: list[tuple[Source, str | None]], label: str, problems: list[str]):
    """按 正文 > 页眉 > 文件名 的顺序取第一个；几处写得不一样就记一笔。"""
    present = [(source, value) for source, value in found if value]
    if not present:
        return None, None
    source, value = present[0]
    for other_source, other_value in present[1:]:
        if other_value != value:
            problems.append(
                f"{label}不一致：{source.value}里是 {value}，{other_source.value}里是 {other_value}"
            )
    return value, source


def resolve_identity(
    filename: str,
    body_lines: list[str],
    header_lines: list[str],
    roster: Roster | None = None,
) -> Identity:
    opening = [line for line in body_lines if line.strip()][:SEARCH_DEPTH]
    body = _from_lines(opening, roster)
    header = _from_lines(header_lines, roster)
    file_id_match = ID_PATTERN.search(Path(filename).stem)
    file_id = file_id_match.group() if file_id_match else None
    file_name = name_from_filename(filename, roster)

    identity = Identity()
    for label, values in (("学号", body.ids), ("姓名", body.names)):
        if len(values) > 1:
            identity.problems.append(
                f"开头写了不止一个{label}（{'、'.join(values)}），可能是几个人合交的，请确认这份算谁的"
            )
    identity.name, identity.name_source = _pick(
        [(Source.BODY, body.name), (Source.HEADER, header.name), (Source.FILENAME, file_name)],
        "姓名",
        identity.problems,
    )
    identity.student_id, identity.id_source = _pick(
        [(Source.BODY, body.student_id), (Source.HEADER, header.student_id), (Source.FILENAME, file_id)],
        "学号",
        identity.problems,
    )
    if roster is not None:
        _check_against_roster(identity, roster)
    return identity


def teacher_identity(name: str | None, student_id: str | None, roster: Roster | None) -> Identity:
    """老师在决定表里指定的姓名学号。照样和名单核对一遍，免得填错。"""
    identity = Identity(
        name=name,
        student_id=student_id,
        name_source=Source.TEACHER if name else None,
        id_source=Source.TEACHER if student_id else None,
    )
    if roster is not None:
        _check_against_roster(identity, roster)
    return identity


def _check_against_roster(identity: Identity, roster: Roster) -> None:
    name, student_id = identity.name, identity.student_id
    if name and student_id:
        expected = roster.name_of(student_id)
        if expected is None:
            identity.problems.append(f"学号 {student_id} 不在名单里")
        elif expected != name:
            identity.problems.append(
                f"姓名和学号对不上：名单里 {student_id} 是 {expected}，这份作业写的是 {name}"
            )
    elif name:
        looked_up = roster.id_of(name)
        if looked_up:
            identity.student_id, identity.id_source = looked_up, Source.ROSTER
        elif not roster.has_name(name):
            identity.problems.append(f"名单里没有「{name}」")
        else:
            identity.problems.append(f"名单里有不止一个「{name}」，无法确定学号")
    elif student_id:
        looked_up = roster.name_of(student_id)
        if looked_up:
            identity.name, identity.name_source = looked_up, Source.ROSTER
        else:
            identity.problems.append(f"学号 {student_id} 不在名单里")
