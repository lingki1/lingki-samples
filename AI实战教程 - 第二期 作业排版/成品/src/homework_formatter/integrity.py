"""底线检查：排版前后一个字都不能少，图片、表格、公式数量不能变。

只比内容，不看格式。做法是把两份文档都拆成一行一行的文字，逐行对照。

这里故意不用"认结构"（structure.classify）：排版用它来决定改哪里，如果比对也靠它，
它认错的时候两边会一起错。比对只认三种允许的差别，每一种都自己重新核实：
1. 空白：行首行尾的空格、空行、连着的多个空格。
2. 姓名行：开头几行里"除了班级、这位学生的姓名和学号以外什么都没有"的行可以消失，
   换成一行统一写法的姓名行。夹着别的字的行不在此列。
3. 小标题编号：一行开头的编号可以换成「一、」这种写法（或去掉），编号后面的字必须一样。
"""

from __future__ import annotations

import io
import re
from collections import Counter
from dataclasses import dataclass

import docx
from docx.oxml.ns import qn

from .identity import CLEAN, SEARCH_DEPTH, identity_leftover, info_line_kind
from .numbering import STYLE_CN, split_number

_M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
CONTEXT = 12
# 不是文字、但同样不能丢的东西：数个数
COUNTED = {
    "表格": (qn("w:tbl"),),
    "图片": (qn("w:drawing"), qn("w:pict"), qn("w:object")),
    "公式": (_M + "oMath",),
    "插入的符号": (qn("w:sym"),),
    "脚注或尾注": (qn("w:footnoteReference"), qn("w:endnoteReference")),
}


@dataclass
class _Line:
    text: str
    # 是不是正文最外层的段落里的（表格、文本框里的不算）；只有这种才可能是姓名行
    top_level: bool


def _own_text(paragraph) -> list[str]:
    """这一段自己的文字，按软回车分成几行。嵌在它里面的别的段落（文本框）不算它的。"""
    lines = [""]

    def walk(element):
        for child in element:
            tag = child.tag
            if tag == qn("w:p"):
                continue
            if tag in (qn("w:t"), _M + "t"):
                lines[-1] += child.text or ""
            elif tag == qn("w:tab"):
                lines[-1] += " "
            elif tag in (qn("w:br"), qn("w:cr")):
                lines.append("")
            else:
                walk(child)

    walk(paragraph)
    return lines


def _normalize(text: str) -> str:
    return " ".join(text.split())


def _read(docx_bytes: bytes) -> tuple[list[_Line], Counter]:
    body = docx.Document(io.BytesIO(docx_bytes)).element.body
    lines = []
    for paragraph in body.iter(qn("w:p")):
        top_level = paragraph.getparent() is body
        for text in _own_text(paragraph):
            text = _normalize(text)
            if text:
                lines.append(_Line(text, top_level))
    counts = Counter({label: sum(1 for _ in body.iter(*tags)) for label, tags in COUNTED.items()})
    return lines, counts


def _same_line(before: str, after: str) -> bool:
    if before == after:
        return True
    # 允许的差别 3：开头的编号换成了「一、」的写法，或者去掉了；后面的字一个不差
    before_style, before_core = split_number(before)
    after_style, after_core = split_number(after)
    renumbered = (before_style, after_style) != (None, None) and after_style in (STYLE_CN, None)
    return renumbered and _normalize(before_core) == _normalize(after_core)


def _excerpts(before: list[str], after: list[str], index: int) -> tuple[str, str]:
    """第一处不一样的地方，前后各带几个字，方便老师一眼看出差别。"""
    a = before[index] if index < len(before) else None
    b = after[index] if index < len(after) else None
    start = 0
    if a is not None and b is not None:
        while start < min(len(a), len(b)) and a[start] == b[start]:
            start += 1
        start = max(0, start - CONTEXT)

    def cut(text):
        if text is None:
            return "（到此结束）"
        piece = text[start : start + 2 * CONTEXT]
        return ("…" if start else "") + piece + ("…" if start + 2 * CONTEXT < len(text) else "")

    return cut(a), cut(b)


def verify_content(
    original: bytes,
    formatted: bytes,
    info_line: str | None,
    name: str | None = None,
    student_id: str | None = None,
) -> list[str]:
    """返回发现的问题；空列表表示内容完全一致。

    info_line 传 None 表示姓名行没有改写（拿不准是谁的，或姓名行里夹着别的内容）：
    这时姓名行也必须原样不变。传了 info_line 就要同时传这位学生的 name 和 student_id。
    """
    before, before_counts = _read(original)
    after, after_counts = _read(formatted)
    problems = []

    if info_line is not None:
        # 允许的差别 2：原文开头的纯姓名行消失，排版后多出一行统一写法的姓名行
        opening = [line for line in before if line.top_level][:SEARCH_DEPTH]
        dropped = {
            id(line)
            for line in opening
            if info_line_kind(line.text) == CLEAN and not identity_leftover(line.text, name, student_id)
        }
        before = [line for line in before if id(line) not in dropped]
        expected = _normalize(info_line)
        matches = [line for line in after if line.top_level and line.text == expected]
        if len(matches) != 1:
            found = f"出现了 {len(matches)} 次" if matches else "没有找到"
            problems.append(f"姓名行不对：应该有且只有一行「{info_line}」，排版后{found}")
        after = [line for line in after if not matches or line is not matches[0]]

    before_texts, after_texts = [l.text for l in before], [l.text for l in after]
    for index in range(max(len(before_texts), len(after_texts))):
        both = index < len(before_texts) and index < len(after_texts)
        if not both or not _same_line(before_texts[index], after_texts[index]):
            was, now = _excerpts(before_texts, after_texts, index)
            problems.append(f"排版前后文字不一致：原文是「{was}」，排版后是「{now}」")
            break

    for label in COUNTED:
        if before_counts[label] != after_counts[label]:
            problems.append(
                f"{label}数量变了：原来 {before_counts[label]} 个，排版后 {after_counts[label]} 个"
            )
    return problems
