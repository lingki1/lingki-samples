"""格式规则：从「规则.toml」读出学校的要求。"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import tomli

TWIPS_PER_CM = 1440 / 2.54  # Word 内部的长度单位：1 厘米 ≈ 567
EMU_PER_TWIP = 635

SIZE_NAMES = {
    "二号": 22, "小二": 18, "三号": 16, "小三": 15,
    "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
}  # fmt: skip
ALIGNMENTS = {"居中": "center", "靠左": "left", "靠右": "right", "两端对齐": "both"}


class RulesError(Exception):
    """规则表读不懂。信息是写给老师看的。"""


@dataclass(frozen=True)
class TextStyle:
    east_asian_font: str
    western_font: str
    size_pt: float
    alignment: str  # Word 内部写法：center / left / right / both
    first_line_chars: int = 0


@dataclass(frozen=True)
class Rules:
    page_width_cm: float
    page_height_cm: float
    margin_top_bottom_cm: float
    margin_left_right_cm: float
    title: TextStyle
    info: TextStyle
    heading: TextStyle
    body: TextStyle
    table_text: TextStyle
    page_number: TextStyle
    line_spacing: float
    class_name: str
    info_separator: str
    unify_heading_numbers: bool
    unnumbered_headings: tuple[str, ...]

    def info_line(self, name: str, student_id: str) -> str:
        return self.info_separator.join([self.class_name, name, student_id])

    @property
    def page_width_twips(self) -> int:
        return round(self.page_width_cm * TWIPS_PER_CM)

    @property
    def page_height_twips(self) -> int:
        return round(self.page_height_cm * TWIPS_PER_CM)

    @property
    def margin_top_bottom_twips(self) -> int:
        return round(self.margin_top_bottom_cm * TWIPS_PER_CM)

    @property
    def margin_left_right_twips(self) -> int:
        return round(self.margin_left_right_cm * TWIPS_PER_CM)

    @property
    def text_width_twips(self) -> int:
        """版心宽度：纸宽减去左右页边距。"""
        return self.page_width_twips - 2 * self.margin_left_right_twips

    @property
    def line_spacing_value(self) -> int:
        return round(240 * self.line_spacing)  # Word 里单倍行距记作 240


def _get(table: dict, section: str, key: str):
    try:
        return table[section][key]
    except KeyError:
        raise RulesError(f"规则表的 [{section}] 里缺少「{key}」") from None


def _is_number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _number(table: dict, section: str, key: str) -> float:
    """必须是数字。老师手改时容易把数字写进引号里，那样后面算尺寸会出错，所以一读就查。"""
    value = _get(table, section, key)
    if not _is_number(value) or value <= 0:
        raise RulesError(f"规则表 [{section}] 的「{key}」应该是一个大于 0 的数字（不带引号），现在是 {value!r}")
    return value


def _whole_number(table: dict, section: str, key: str) -> int:
    value = _get(table, section, key)
    if not _is_number(value) or value < 0 or value != int(value):
        raise RulesError(f"规则表 [{section}] 的「{key}」应该是整数（不带引号），现在是 {value!r}")
    return int(value)


def _text(table: dict, section: str, key: str) -> str:
    value = _get(table, section, key)
    if not isinstance(value, str) or not value:
        raise RulesError(f"规则表 [{section}] 的「{key}」应该是带引号的文字，现在是 {value!r}")
    return value


def _size(table: dict, section: str) -> float:
    value = _get(table, section, "字号")
    if _is_number(value) and value > 0:
        return value
    if value not in SIZE_NAMES:
        raise RulesError(
            f"规则表 [{section}] 的字号「{value}」不认识，可以写：{'、'.join(SIZE_NAMES)}，或磅值"
        )
    return SIZE_NAMES[value]


def _alignment(table: dict, section: str, key: str = "对齐") -> str:
    value = _get(table, section, key)
    if value not in ALIGNMENTS:
        raise RulesError(
            f"规则表 [{section}] 的{key}「{value}」不认识，可以写：{'、'.join(ALIGNMENTS)}"
        )
    return ALIGNMENTS[value]


def _single_font_style(table: dict, section: str) -> TextStyle:
    font = _text(table, section, "字体")
    return TextStyle(font, font, _size(table, section), _alignment(table, section))


def load_rules(path: Path) -> Rules:
    try:
        # utf-8-sig：记事本存出来的文件开头可能带一个看不见的标记，要认得
        table = tomli.loads(Path(path).read_text(encoding="utf-8-sig"))
    except UnicodeDecodeError as error:
        raise RulesError(
            f"规则表「{Path(path).name}」的编码不对。请用记事本打开它，「另存为」时把编码选成 UTF-8"
        ) from error
    except tomli.TOMLDecodeError as error:
        raise RulesError(f"规则表「{Path(path).name}」的写法有误：{error}") from error

    body = TextStyle(
        east_asian_font=_text(table, "正文", "中文字体"),
        western_font=_text(table, "正文", "西文字体"),
        size_pt=_size(table, "正文"),
        alignment=_alignment(table, "正文"),
        first_line_chars=_whole_number(table, "正文", "首行缩进_字"),
    )
    return Rules(
        page_width_cm=_number(table, "页面", "纸张宽_厘米"),
        page_height_cm=_number(table, "页面", "纸张高_厘米"),
        margin_top_bottom_cm=_number(table, "页面", "上下边距_厘米"),
        margin_left_right_cm=_number(table, "页面", "左右边距_厘米"),
        title=_single_font_style(table, "题目"),
        info=_single_font_style(table, "姓名行"),
        heading=_single_font_style(table, "小标题"),
        body=body,
        table_text=TextStyle(
            body.east_asian_font,
            body.western_font,
            body.size_pt,
            _alignment(table, "表格", "文字对齐"),
        ),
        page_number=_single_font_style(table, "页码"),
        line_spacing=_number(table, "正文", "行距_倍"),
        class_name=_text(table, "姓名行", "班级"),
        info_separator=_get(table, "姓名行", "间隔"),
        unify_heading_numbers=bool(_get(table, "小标题", "统一编号")),
        unnumbered_headings=tuple(_get(table, "小标题", "不编号的小标题")),
    )
