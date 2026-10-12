"""认结构：判断每一段是题目、姓名行、小标题还是正文。"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from enum import Enum

import docx
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .identity import SEARCH_DEPTH, is_info_line
from .numbering import split_number

TITLE_MAX_LENGTH = 50
NUMBERED_HEADING_MAX_LENGTH = 25
PLAIN_HEADING_MAX_LENGTH = 20

NUMBERED_PREFIX = re.compile(
    r"^\s*("
    r"第[一二三四五六七八九十]+(部分|章|节)"
    r"|[一二三四五六七八九十]+[、．.]"
    r"|[（(][一二三四五六七八九十0-9０-９]+[）)]"
    r"|[0-9０-９]+[、．.](?![0-9０-９])"
    r")"
)
REFERENCE_ENTRY = re.compile(r"^\s*[\[［【]\s*\d+\s*[\]］】]")
# 没有编号的短行里，这几种不是小标题：图表的题注、日期、算式
CAPTION = re.compile(r"^\s*[图表]\s*[0-9０-９一二三四五六七八九十]")
DATE = re.compile(r"^\s*\d{4}\s*[年./-]\s*\d{1,2}")
SENTENCE_END = "。！？；.!?;"
ANY_PUNCTUATION = set("，。；：？！、,.;:?!")
HEADING_STYLE = re.compile(r"^(heading|标题)\s*\d", re.IGNORECASE)

OBJECT_TAGS = (qn("w:drawing"), qn("w:pict"), qn("w:object"))
# 段落里只有这些东西时，这一段才算"空的"。别的任何东西（公式、符号、脚注引用、域……）都算内容。
_BLANK_PARAGRAPH_TAGS = {
    qn("w:r"),
    qn("w:t"),
    qn("w:tab"),
    qn("w:br"),
    qn("w:bookmarkStart"),
    qn("w:bookmarkEnd"),
    qn("w:proofErr"),
    qn("w:lastRenderedPageBreak"),
    qn("w:permStart"),
    qn("w:permEnd"),
}
_PROPERTY_TAGS = {qn("w:pPr"), qn("w:rPr")}
# 直接挂在正文下面、不显示任何东西的标记（跨段书签的结尾、拼写检查标记）
_INVISIBLE_BODY_TAGS = {
    qn("w:bookmarkStart"),
    qn("w:bookmarkEnd"),
    qn("w:proofErr"),
    qn("w:permStart"),
    qn("w:permEnd"),
}


class Role(Enum):
    TITLE = "题目"
    INFO = "姓名行"
    HEADING = "小标题"
    BODY = "正文"
    IMAGE = "图片"
    EMPTY = "空段落"
    TABLE = "表格"
    OTHER = "其他"


@dataclass
class Block:
    role: Role
    element: object  # 文档里对应的那个段落或表格
    text: str = ""


@dataclass
class Structure:
    title: str | None
    headings: list[str]


def is_soft_break(element) -> bool:
    """软回车（Shift+回车）；分页符、分栏符不算。"""
    return element.tag == qn("w:br") and element.get(qn("w:type")) in (None, "textWrapping")


def has_page_break(paragraph_element) -> bool:
    return any(not is_soft_break(br) for br in paragraph_element.iter(qn("w:br")))


def has_soft_break(paragraph_element) -> bool:
    return any(is_soft_break(br) for br in paragraph_element.iter(qn("w:br")))


def _content(element):
    """段落里除了格式设置以外的所有东西。"""
    for child in element:
        if child.tag in _PROPERTY_TAGS:
            continue
        yield child
        yield from _content(child)


def is_blank_paragraph(paragraph_element) -> bool:
    """真正什么内容都没有的段落：只有空白，没有公式、符号、图片、脚注引用等任何东西。"""
    for node in _content(paragraph_element):
        if node.tag not in _BLANK_PARAGRAPH_TAGS:
            return False
        if node.tag == qn("w:t") and (node.text or "").strip():
            return False
    return True


def paragraph_role(paragraph_element) -> tuple[Role, str]:
    """只看这一段本身：有字（先当正文）、图片、空段落，还是程序不认识的内容。"""
    text = Paragraph(paragraph_element, None).text
    if text.strip():
        return Role.BODY, text
    if is_blank_paragraph(paragraph_element):
        return Role.EMPTY, ""
    if next(paragraph_element.iter(*OBJECT_TAGS), None) is not None:
        return Role.IMAGE, ""
    # 有东西但不是普通文字：独占一行的公式、插入的符号、整段在修订或域里……原样留着
    return Role.OTHER, ""


def is_heading_text(text: str) -> bool:
    """只看文字本身，像不像一个小标题。"""
    text = text.strip()
    if not text or REFERENCE_ENTRY.match(text):
        return False
    if NUMBERED_PREFIX.match(text):
        return len(text) <= NUMBERED_HEADING_MAX_LENGTH and text[-1] not in SENTENCE_END
    if CAPTION.match(text) or DATE.match(text) or "=" in text or "＝" in text:
        return False
    # 没有编号：要很短，而且中间没有任何标点（末尾的冒号不算）
    core = text.rstrip("：:")
    return len(text) <= PLAIN_HEADING_MAX_LENGTH and not (set(core) & ANY_PUNCTUATION)


def looks_like_title(text: str) -> bool:
    text = text.strip()
    return (
        len(text) <= TITLE_MAX_LENGTH
        and not NUMBERED_PREFIX.match(text)
        and text[-1] not in SENTENCE_END
    )


def _heading_style_ids(document) -> set[str]:
    return {
        style.style_id
        for style in document.styles
        if HEADING_STYLE.match(style.name or "") or HEADING_STYLE.match(style.style_id or "")
    }


def _style_id(paragraph_element) -> str | None:
    ppr = paragraph_element.pPr
    return ppr.pStyle.val if ppr is not None and ppr.pStyle is not None else None


def classify(document) -> list[Block]:
    heading_styles = _heading_style_ids(document)
    blocks: list[Block] = []
    for element in document.element.body.iterchildren():
        if element.tag == qn("w:tbl"):
            blocks.append(Block(Role.TABLE, element))
        elif element.tag == qn("w:p"):
            role, text = paragraph_role(element)
            blocks.append(Block(role, element, text))
        elif element.tag != qn("w:sectPr") and element.tag not in _INVISIBLE_BODY_TAGS:
            blocks.append(Block(Role.OTHER, element))

    texts = [b for b in blocks if b.role is Role.BODY]
    for block in texts[:SEARCH_DEPTH]:
        if is_info_line(block.text) and not has_soft_break(block.element):
            block.role = Role.INFO
    if texts and texts[0].role is Role.BODY and looks_like_title(texts[0].text):
        texts[0].role = Role.TITLE

    guessed: list[Block] = []  # 只凭文字猜出来的小标题（没有用标题样式）
    for block in texts:
        if block.role is not Role.BODY or has_soft_break(block.element):
            continue
        if _style_id(block.element) in heading_styles:
            block.role = Role.HEADING
        elif is_heading_text(block.text):
            block.role = Role.HEADING
            guessed.append(block)
    _demote_headings_with_nothing_under_them(blocks, guessed)
    return blocks


def _demote_headings_with_nothing_under_them(blocks: list[Block], guessed: list[Block]) -> None:
    """只凭文字猜出来的小标题，有两种情况其实不是小标题，改回正文：

    1. 同一种写法的短行连着出现（中间没有正文）：「1. 手机」「2. 电脑」「3. 平板」是一组小条目，
       连着几行的人名、诗句也是。大标题紧跟着不同写法的小标题（「一、背景」「（一）起因」）不算。
    2. 后面什么内容都没有了：文末单独一行的「谢谢老师」、落款、日期。
    """
    visible = [b for b in blocks if b.role is not Role.EMPTY]
    is_guessed = {id(b) for b in guessed}

    run: list[Block] = []
    for block in visible + [None]:
        same_run = (
            block is not None
            and id(block) in is_guessed
            and (not run or split_number(block.text)[0] == split_number(run[0].text)[0])
        )
        if same_run:
            run.append(block)
            continue
        if len(run) > 1:
            for item in run:
                item.role = Role.BODY
        run = [block] if block is not None and id(block) in is_guessed else []

    for block in reversed(visible):  # 从文末往前：最后一行不是小标题，那它前一行也可能不是
        if block.role is Role.HEADING and id(block) in is_guessed:
            block.role = Role.BODY
        else:
            break


def summarize(blocks: list[Block]) -> Structure:
    titles = [b.text.strip() for b in blocks if b.role is Role.TITLE]
    return Structure(
        title=titles[0] if titles else None,
        headings=[b.text.strip() for b in blocks if b.role is Role.HEADING],
    )


def analyze(docx_bytes: bytes) -> Structure:
    return summarize(classify(docx.Document(io.BytesIO(docx_bytes))))
