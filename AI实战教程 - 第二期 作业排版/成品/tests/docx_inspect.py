"""测试用：把排好的文档重新打开，读出每一段实际写着的格式。"""

import io

import docx
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from homework_formatter.rules import Rules, TextStyle


def open_bytes(data: bytes):
    return docx.Document(io.BytesIO(data))


def save_bytes(document) -> bytes:
    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def _attr(element, name):
    return element.get(qn(name)) if element is not None else None


def _on(rpr, tag) -> bool:
    element = rpr.find(qn(tag)) if rpr is not None else None
    return element is not None and _attr(element, "w:val") not in ("0", "false", "off", "none")


def paragraph_format(p) -> dict:
    ppr = p.find(qn("w:pPr"))
    child = (lambda tag: ppr.find(qn(tag))) if ppr is not None else (lambda tag: None)
    spacing, ind = child("w:spacing"), child("w:ind")
    return {
        "style": _attr(child("w:pStyle"), "w:val"),
        "jc": _attr(child("w:jc"), "w:val"),
        "line": _attr(spacing, "w:line"),
        "line_rule": _attr(spacing, "w:lineRule"),
        "before": _attr(spacing, "w:before"),
        "after": _attr(spacing, "w:after"),
        "first_line": _attr(ind, "w:firstLine"),
        "first_line_chars": _attr(ind, "w:firstLineChars"),
        "left": _attr(ind, "w:left"),
        "has_border": child("w:pBdr") is not None,
    }


def run_format(r) -> dict:
    rpr = r.find(qn("w:rPr"))
    child = (lambda tag: rpr.find(qn(tag))) if rpr is not None else (lambda tag: None)
    fonts, size = child("w:rFonts"), child("w:sz")
    return {
        "east_asian": _attr(fonts, "w:eastAsia"),
        "ascii": _attr(fonts, "w:ascii"),
        "h_ansi": _attr(fonts, "w:hAnsi"),
        "size_pt": int(_attr(size, "w:val")) / 2 if size is not None else None,
        "color": _attr(child("w:color"), "w:val"),
        "bold": _on(rpr, "w:b"),
        "italic": _on(rpr, "w:i"),
        "underline": _on(rpr, "w:u"),
        "highlight": _attr(child("w:highlight"), "w:val"),
        "vert_align": _attr(child("w:vertAlign"), "w:val"),
        "style": _attr(child("w:rStyle"), "w:val"),
    }


def text_runs(p) -> list:
    return [r for r in p.iter(qn("w:r")) if "".join(t.text or "" for t in r.iter(qn("w:t")))]


def paragraph_text(p) -> str:
    return Paragraph(p, None).text


def body_paragraphs(document) -> list:
    """正文里的顶层段落（不含表格里的）。"""
    return [e for e in document.element.body.iterchildren() if e.tag == qn("w:p")]


def find(document, text):
    for p in body_paragraphs(document):
        if paragraph_text(p) == text:
            return p
    raise AssertionError(f"文档里没有这一段：{text!r}")


def style_problems(p, style: TextStyle, rules: Rules, line=None) -> list[str]:
    """这一段和规则不一致的地方；空列表表示完全符合。"""
    problems = []
    fmt = paragraph_format(p)
    expected = {
        "style": None,
        "jc": style.alignment,
        "line": str(line or rules.line_spacing_value),
        "line_rule": "auto",
        "before": "0",
        "after": "0",
        "first_line": str(int(style.first_line_chars * style.size_pt * 20)),
        "first_line_chars": str(style.first_line_chars * 100),
        "left": "0",
        "has_border": False,
    }
    for key, value in expected.items():
        if fmt[key] != value:
            problems.append(f"段落 {key}: 应为 {value!r}，实际 {fmt[key]!r}")
    for r in text_runs(p):
        actual = run_format(r)
        expected_run = {
            "east_asian": style.east_asian_font,
            "ascii": style.western_font,
            "h_ansi": style.western_font,
            "size_pt": style.size_pt,
            "color": "000000",
            "bold": False,
            "italic": False,
            "underline": False,
            "highlight": None,
            "style": None,
        }
        for key, value in expected_run.items():
            if actual[key] != value:
                problems.append(f"文字 {key}: 应为 {value!r}，实际 {actual[key]!r}")
    return problems


def page_setup(document) -> list[dict]:
    result = []
    for sect in document.element.body.iter(qn("w:sectPr")):
        size, margin = sect.find(qn("w:pgSz")), sect.find(qn("w:pgMar"))
        result.append(
            {
                "width": _attr(size, "w:w"),
                "height": _attr(size, "w:h"),
                "orient": _attr(size, "w:orient"),
                "top": _attr(margin, "w:top"),
                "bottom": _attr(margin, "w:bottom"),
                "left": _attr(margin, "w:left"),
                "right": _attr(margin, "w:right"),
            }
        )
    return result


def footer_paragraphs(document) -> list:
    return list(document.sections[0].footer._element.iter(qn("w:p")))


def has_page_number_field(p) -> bool:
    instructions = "".join(t.text or "" for t in p.iter(qn("w:instrText")))
    kinds = [f.get(qn("w:fldCharType")) for f in p.iter(qn("w:fldChar"))]
    return instructions.strip() == "PAGE" and kinds == ["begin", "separate", "end"]


def header_texts(document) -> list[str]:
    texts = []
    for section in document.sections:
        for header in (section.header, section.first_page_header, section.even_page_header):
            if not header.is_linked_to_previous:
                texts += [p.text for p in header.paragraphs if p.text.strip()]
    return texts
