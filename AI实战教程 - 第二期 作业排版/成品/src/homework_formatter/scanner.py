"""扫描作业文件夹：列出文件、分清类型、计算指纹。只读。"""

from __future__ import annotations

import hashlib
import io
import zipfile
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path

OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
PDF_MAGIC = b"%PDF"
TEMP_PREFIXES = ("~$", ".~lock.")
UNREADABLE = "读不出来"


class FileKind(Enum):
    DOCX = "Word 文档"
    DOC = "旧版 Word（.doc）"
    PDF = "PDF"
    TEMP = "临时文件"
    OTHER = "其他"


@dataclass(frozen=True)
class SourceFile:
    path: Path
    kind: FileKind
    size: int
    sha256: str
    # 文件内容只读这一次，后面都用这份：免得同一个文件读好几遍，中间被别的程序改了
    data: bytes = field(default=b"", repr=False, compare=False)
    # 读不出来时的原因（文件正被别的程序占着、网盘上的占位文件……）
    error: str | None = None

    def read_bytes(self) -> bytes:
        return self.data


def is_temp_name(name: str) -> bool:
    return name.startswith(TEMP_PREFIXES)


def _sniff_kind(name: str, data: bytes) -> FileKind:
    # 看内容而不是只看扩展名：改过扩展名的文件、损坏的文件不能当成 Word 处理
    if is_temp_name(name):
        return FileKind.TEMP
    suffix = Path(name).suffix.lower()
    if data.startswith(PDF_MAGIC):
        return FileKind.PDF
    if suffix == ".doc" and data.startswith(OLE_MAGIC):
        return FileKind.DOC
    if suffix == ".docx" and _is_docx(data):
        return FileKind.DOCX
    return FileKind.OTHER


def _is_docx(data: bytes) -> bool:
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            return "word/document.xml" in archive.namelist()
    except Exception:  # 坏压缩包能抛出的异常五花八门，一律当作"不是 Word 文件"
        return False


def scan_folder(folder: Path) -> list[SourceFile]:
    """列出文件夹里的文件（不进入子文件夹），按文件名排序。读不出来的也列出来，不中断。"""
    entries = []
    for path in sorted(Path(folder).iterdir(), key=lambda p: p.name):
        if not path.is_file():
            continue
        try:
            data = path.read_bytes()
        except OSError as error:
            kind = FileKind.TEMP if is_temp_name(path.name) else FileKind.OTHER
            reason = f"{UNREADABLE}，可能正被别的程序打开（{type(error).__name__}）"
            entries.append(SourceFile(path, kind, 0, "", error=reason))
            continue
        entries.append(
            SourceFile(
                path=path,
                kind=_sniff_kind(path.name, data),
                size=len(data),
                sha256=hashlib.sha256(data).hexdigest(),
                data=data,
            )
        )
    return entries


def subfolders(folder: Path) -> list[str]:
    return sorted(p.name for p in Path(folder).iterdir() if p.is_dir())


def fingerprint(folder: Path) -> dict[str, str]:
    """文件夹里每个原件（含子文件夹）的指纹，用来证明运行前后原件没有变化。

    Word 的临时文件不算原件：老师中途关掉一份正开着的作业，临时文件会自己消失。
    """
    folder = Path(folder)
    result = {}
    for path in sorted(folder.rglob("*")):
        if not path.is_file() or is_temp_name(path.name):
            continue
        try:
            result[path.relative_to(folder).as_posix()] = hashlib.sha256(path.read_bytes()).hexdigest()
        except OSError:
            result[path.relative_to(folder).as_posix()] = UNREADABLE
    return result
