"""开发时用：把排版前后的作业并排画成图片，给老师看效果。

用法（在项目文件夹里）：
    .venv/Scripts/python tools/make_comparison.py 排版结果/某次运行 原文件名1 原文件名2 ...

原件先复制到临时文件夹再交给 LibreOffice，不在「学生作业」里打开任何东西。
图片存到 该次运行/前后对比/。
"""

from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import fitz  # PyMuPDF
import openpyxl

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from homework_formatter.converter import find_soffice  # noqa: E402

INPUT_DIR = PROJECT_ROOT / "学生作业"
GAP = 30
LABEL_HEIGHT = 34
DPI = 110
LABEL_FONT_FILE = Path(r"C:\Windows\Fonts\simhei.ttf")


def to_pdf(sources: dict[str, Path], workdir: Path) -> dict[str, Path]:
    """sources: 标签 -> 文件。返回 标签 -> PDF。"""
    in_dir, out_dir = workdir / "in", workdir / "out"
    in_dir.mkdir()
    out_dir.mkdir()
    copies = {}
    for index, (label, path) in enumerate(sources.items()):
        copy = in_dir / f"{index}{path.suffix}"
        shutil.copyfile(path, copy)
        copies[label] = copy
    subprocess.run(
        [
            str(find_soffice()),
            f"-env:UserInstallation={(workdir / 'profile').as_uri()}",
            "--headless", "--convert-to", "pdf", "--outdir", str(out_dir),
            *[str(c) for c in copies.values()],
        ],  # fmt: skip
        capture_output=True,
        timeout=300,
        check=True,
    )
    return {label: out_dir / f"{copy.stem}.pdf" for label, copy in copies.items()}


def side_by_side(before_pdf: Path, after_pdf: Path, before_label: str, after_label: str, target: Path):
    before, after = fitz.open(before_pdf), fitz.open(after_pdf)
    left, right = before[0].rect, after[0].rect
    width = left.width + right.width + 3 * GAP
    height = max(left.height, right.height) + LABEL_HEIGHT + 2 * GAP
    sheet = fitz.open()
    page = sheet.new_page(width=width, height=height)
    page.draw_rect(page.rect, color=None, fill=(0.85, 0.85, 0.85))
    x = GAP
    for document, rect, label in ((before, left, before_label), (after, right, after_label)):
        top = GAP + LABEL_HEIGHT
        frame = fitz.Rect(x, top, x + rect.width, top + rect.height)
        page.draw_rect(frame, color=(0.5, 0.5, 0.5), fill=(1, 1, 1), width=0.5)
        page.show_pdf_page(frame, document, 0)
        if LABEL_FONT_FILE.is_file():
            font = {"fontname": "label", "fontfile": str(LABEL_FONT_FILE)}
        else:
            font = {"fontname": "china-s"}  # PyMuPDF 自带的简体中文字体，数字会偏宽
        page.insert_text((x, GAP + 16), label, fontsize=13, **font)
        x += rect.width + GAP
    page.get_pixmap(dpi=DPI).save(target)


def main() -> None:
    run_dir = Path(sys.argv[1]).resolve()
    wanted = sys.argv[2:]
    rows = list(openpyxl.load_workbook(run_dir / "处理报告.xlsx")["明细"].iter_rows(values_only=True))
    header = rows[0]
    new_names = {r[header.index("原文件名")]: r[header.index("新文件名")] for r in rows[1:]}

    target_dir = run_dir / "前后对比"
    target_dir.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="homework_compare_") as tmp:
        sources = {}
        for name in wanted:
            sources[f"before:{name}"] = INPUT_DIR / name
            sources[f"after:{name}"] = run_dir / "已完成" / new_names[name]
        pdfs = to_pdf(sources, Path(tmp))
        for name in wanted:
            target = target_dir / f"{Path(new_names[name]).stem}.png"
            side_by_side(
                pdfs[f"before:{name}"],
                pdfs[f"after:{name}"],
                f"排版前：{name}",
                f"排版后：{new_names[name]}",
                target,
            )
            print(target)


if __name__ == "__main__":
    main()
