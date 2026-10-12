"""小标题的编号：认出学生用的是哪种写法，统一成「一、」「二、」「三、」。

只管文字上怎么算，不碰文档。拿不准的情况不给方案，由调用的地方转交老师。
"""

from __future__ import annotations

import re
from dataclasses import dataclass

CN_DIGITS = "一二三四五六七八九十百"
ANY_DIGIT = "0-9０-９"

# 写法的名字就用它的样子，报告里直接给老师看
STYLE_CN = "一、"
STYLE_DIGIT = "1."
STYLE_PAREN = "（一）"
STYLE_PART = "第一部分"

_PATTERNS = [
    (STYLE_PART, re.compile(rf"^第[{CN_DIGITS}]+(部分|章|节)\s*[、．.：:]?\s*")),
    (STYLE_PAREN, re.compile(rf"^[（(]\s*[{CN_DIGITS}{ANY_DIGIT}]+\s*[）)]\s*[、．.]?\s*")),
    (STYLE_CN, re.compile(rf"^[{CN_DIGITS}]+\s*[、．.]\s*")),
    # 后面紧跟数字的不算编号："3.5 倍"是小数
    (STYLE_DIGIT, re.compile(rf"^[{ANY_DIGIT}]+\s*[、．.](?![{ANY_DIGIT}])\s*")),
]


@dataclass
class NumberingPlan:
    # 每个小标题该改成什么；None 表示拿不准，不改
    headings: list[str] | None
    # 拿不准的原因（写给老师看）
    problem: str | None = None
    # 改了什么（写给老师看）；没有改动就是 None
    note: str | None = None


def split_number(text: str) -> tuple[str | None, str]:
    """把小标题拆成（编号写法, 去掉编号后的文字）。没有编号时写法是 None。"""
    text = text.strip()
    for style, pattern in _PATTERNS:
        match = pattern.match(text)
        if match:
            return style, text[match.end() :]
    return None, text


_CN_VALUES = {c: i for i, c in enumerate("零一二三四五六七八九")}
_NUMERAL = re.compile(rf"[{CN_DIGITS}]+|[{ANY_DIGIT}]+")


def number_value(text: str) -> int | None:
    """小标题开头那个编号是几。认不出来返回 None。"""
    match = _NUMERAL.search(text.strip()[:8])
    if not match:
        return None
    token = match.group()
    if token[0] not in CN_DIGITS:
        return int(token)
    if "百" in token:
        return None
    if "十" not in token:
        return _CN_VALUES.get(token) if len(token) == 1 else None
    tens, ones = token.split("十", 1)
    tens_value = _CN_VALUES.get(tens) if tens else 1
    ones_value = _CN_VALUES.get(ones) if ones else 0
    if tens_value is None or ones_value is None:
        return None
    return tens_value * 10 + ones_value


def chinese_numeral(number: int) -> str:
    """1 到 99 的中文数字。"""
    tens, ones = divmod(number, 10)
    digits = "零一二三四五六七八九"
    if tens == 0:
        return digits[ones]
    return ("" if tens == 1 else digits[tens]) + "十" + (digits[ones] if ones else "")


def _is_unnumbered_name(core: str, unnumbered: tuple[str, ...]) -> bool:
    return core.strip().rstrip("：:").strip() in unnumbered


def plan_numbering(headings: list[str], unnumbered: tuple[str, ...]) -> NumberingPlan:
    parts = [split_number(h) for h in headings]
    styles = {style for style, core in parts if not _is_unnumbered_name(core, unnumbered)}

    if STYLE_PART in styles:
        return NumberingPlan(None, "小标题用的是「第一部分」「第一章」这类写法，没有改成「一、」")
    if len(styles) > 1:
        if None in styles and len(styles) == 2:
            return NumberingPlan(None, "有的小标题编了号、有的没编，不确定该怎么统一，编号没有改")
        used = "、".join(f"「{s}」" for s in sorted(s for s in styles if s))
        return NumberingPlan(
            None, f"小标题混用了两种以上的编号写法（{used}），可能是大标题套小标题，编号没有改"
        )

    numbered = [
        (h.strip(), core)
        for h, (style, core) in zip(headings, parts)
        if style is not None and not _is_unnumbered_name(core, unnumbered)
    ]
    values = [number_value(heading) for heading, _ in numbered]
    if values and (None in values or values[0] != 1 or any(b <= a for a, b in zip(values, values[1:]))):
        # 1、2、1、2、3 这样往回数的，多半是大标题下面套着同样写法的小条目
        return NumberingPlan(
            None, "小标题原来的编号不是从 1 开始一路往上排的，可能是大标题下面套着小条目，编号没有改"
        )
    if values != list(range(1, len(values) + 1)):
        # 编号有跳号时，「一、二年级的情况」「三、四年级的情况」这种就分不清哪个数字是编号了。
        # 编号正好是 1、2、3 连着的，「三、四点建议」里的「三、」可以放心当编号。
        for heading, core in numbered:
            if _NUMERAL.match(core) or (heading[:1] in "（(" and core.startswith("班")):
                return NumberingPlan(
                    None, f"小标题「{heading}」开头不确定哪个数字是编号，编号没有改"
                )

    result, count, stripped = [], 0, []
    for style, core in parts:
        if _is_unnumbered_name(core, unnumbered):
            result.append(core)
            if style is not None:
                stripped.append(core)
        else:
            count += 1
            result.append(f"{chinese_numeral(count)}、{core}")

    notes = []
    numbered_before = [h.strip() for h, (_, c) in zip(headings, parts) if not _is_unnumbered_name(c, unnumbered)]
    numbered_after = [r for r, (_, c) in zip(result, parts) if not _is_unnumbered_name(c, unnumbered)]
    if numbered_before != numbered_after:
        (style,) = styles
        if style is None:
            was = "原来没有编号"
        elif style == STYLE_CN:
            was = "原来的编号中间有跳号，或标点不同"
        else:
            was = f"原来是「{style}」这种写法"
        notes.append(f"小标题的编号统一成了「一、二、三」（{was}）")
    if stripped:
        notes.append(f"去掉了「{'」「'.join(stripped)}」前面的编号（这类小标题统一不编号）")
    return NumberingPlan(result, note="；".join(notes) or None)
