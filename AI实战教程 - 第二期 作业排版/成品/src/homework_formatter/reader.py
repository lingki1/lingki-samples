"""读出一份 .docx 的内容。只在内存里读，不写回任何文件。"""

from __future__ import annotations

import io
from dataclasses import dataclass

import docx
from docx.oxml.ns import qn

IMAGE_TAGS = (qn("w:drawing"), qn("w:pict"), qn("w:object"))
_MATH_TEXT = "{http://schemas.openxmlformats.org/officeDocument/2006/math}t"


class DocxReadError(Exception):
    """文件打不开。信息是写给老师看的。"""


@dataclass
class DocContent:
    paragraphs: list[str]
    header_lines: list[str]
    table_count: int
    image_count: int
    # 整份文档里有没有任何文字（包括表格、文本框、内容控件、公式里的）
    has_any_text: bool = True

    @property
    def is_blank(self) -> bool:
        return not self.has_any_text and self.table_count == 0 and self.image_count == 0


def read_docx(data: bytes) -> DocContent:
    # 坏文件能在任何一步出错（比如页眉指向一个不存在的部件），所以整个读取过程都保护起来
    try:
        document = docx.Document(io.BytesIO(data))
        header_lines = []
        for section in document.sections:
            # 没有自己页眉的节，python-docx 一访问就会新建一个；先判断再读
            if section.header.is_linked_to_previous:
                continue
            for paragraph in section.header._element.iter(qn("w:p")):
                text = "".join(t.text or "" for t in paragraph.iter(qn("w:t"))).strip()
                if text:
                    header_lines.append(text)

        body = document.element.body
        return DocContent(
            paragraphs=[p.text for p in document.paragraphs],
            header_lines=header_lines,
            table_count=sum(1 for _ in body.iter(qn("w:tbl"))),
            image_count=sum(1 for _ in body.iter(*IMAGE_TAGS)),
            has_any_text=any((t.text or "").strip() for t in body.iter(qn("w:t"), _MATH_TEXT)),
        )
    except Exception as error:  # python-docx 对坏文件抛的异常种类很多
        raise DocxReadError(f"文件打不开，可能已损坏或加了密码（{type(error).__name__}）") from error
