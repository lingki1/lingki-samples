"""老师对个别文件的决定：这份是谁的、这份不要了。程序照办，不自己判断。"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

from .roster import _cell_text

FILE_HEADER = "原文件名"
NAME_HEADER = "姓名"
ID_HEADER = "学号"
DISCARD_HEADER = "不要这份"
NOTE_HEADER = "备注"
REQUIRED_HEADERS = (FILE_HEADER, NAME_HEADER, ID_HEADER, DISCARD_HEADER)
# 「不要这份」一列只认明确的写法：填"否"的本意是"这份要"，不能也当成不要
YES = {"是", "不要", "不要这份", "√", "✓", "✔", "1", "y", "yes"}
NO = {"", "否", "要", "0", "n", "no"}


class DecisionsError(Exception):
    """决定表读不懂。信息是写给老师看的。"""


@dataclass
class Decisions:
    # 原文件名 -> (姓名, 学号)，没填的那一项是 None
    assigned: dict[str, tuple[str | None, str | None]] = field(default_factory=dict)
    # 原文件名 -> 备注
    discarded: dict[str, str] = field(default_factory=dict)

    def filenames(self) -> set[str]:
        return set(self.assigned) | set(self.discarded)


def load_decisions(path: Path) -> Decisions:
    workbook = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    try:
        rows = [[_cell_text(v) for v in row] for row in workbook.worksheets[0].iter_rows(values_only=True)]
    finally:
        workbook.close()

    header = rows[0] if rows else []
    missing = [h for h in REQUIRED_HEADERS if h not in header]
    if missing:
        raise DecisionsError(f"「{Path(path).name}」的第一行缺少这几列：{'、'.join(missing)}")
    column = {h: header.index(h) for h in header if h}

    def cell(row: list[str], name: str) -> str:
        index = column.get(name)
        return row[index] if index is not None and index < len(row) else ""

    decisions = Decisions()
    for number, row in enumerate(rows[1:], start=2):
        filename = cell(row, FILE_HEADER)
        if not filename:
            continue
        name, student_id = cell(row, NAME_HEADER), cell(row, ID_HEADER)
        where = f"「{Path(path).name}」第 {number} 行（{filename}）"
        answer = cell(row, DISCARD_HEADER).lower()
        if answer not in YES | NO:
            raise DecisionsError(
                f"{where}的「{DISCARD_HEADER}」填的是「{cell(row, DISCARD_HEADER)}」，看不懂；不要这份请填「是」，否则留空"
            )
        discard = answer in YES
        if filename in decisions.filenames():
            raise DecisionsError(f"{where}：这个文件在表里出现了两次，请只留一行")
        if discard and (name or student_id):
            raise DecisionsError(f"{where}既填了姓名学号，又填了「{DISCARD_HEADER}」，只能选一样")
        if discard:
            decisions.discarded[filename] = cell(row, NOTE_HEADER)
        elif name or student_id:
            decisions.assigned[filename] = (name or None, student_id or None)
        else:
            raise DecisionsError(f"{where}没有填姓名学号，也没有填「{DISCARD_HEADER}」")
    return decisions
