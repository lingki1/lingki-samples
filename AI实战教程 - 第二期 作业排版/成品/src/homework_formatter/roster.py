"""班级名单：学号和姓名的对照表。"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

import openpyxl

ID_HEADER = "学号"
NAME_HEADER = "姓名"
HEADER_SEARCH_ROWS = 10


class RosterError(Exception):
    """名单文件读不懂。信息是写给老师看的。"""


class Roster:
    def __init__(self, entries: Iterable[tuple[str, str]]):
        self.entries: list[tuple[str, str]] = [(str(i), str(n)) for i, n in entries]
        self._name_by_id: dict[str, str] = {}
        self._ids_by_name: dict[str, list[str]] = {}
        for student_id, name in self.entries:
            if student_id in self._name_by_id:
                raise RosterError(f"名单里学号 {student_id} 出现了不止一次")
            self._name_by_id[student_id] = name
            self._ids_by_name.setdefault(name, []).append(student_id)

    def __len__(self) -> int:
        return len(self.entries)

    def name_of(self, student_id: str) -> str | None:
        return self._name_by_id.get(student_id)

    def id_of(self, name: str) -> str | None:
        """按姓名查学号；重名时无法确定，返回 None。"""
        ids = self._ids_by_name.get(name, [])
        return ids[0] if len(ids) == 1 else None

    def has_name(self, name: str) -> bool:
        return name in self._ids_by_name

    def names_in(self, text: str) -> list[str]:
        """text 里出现了名单上的哪些姓名。一个姓名被另一个更长的姓名包含时只算长的。"""
        found = [name for name in self._ids_by_name if name in text]
        longest = [n for n in found if not any(n != m and n in m for m in found)]
        return sorted(longest, key=text.index)


def _cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def load_roster(path: Path) -> Roster:
    workbook = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    try:
        for sheet in workbook.worksheets:
            rows = [[_cell_text(v) for v in row] for row in sheet.iter_rows(values_only=True)]
            for index, row in enumerate(rows[:HEADER_SEARCH_ROWS]):
                if ID_HEADER in row and NAME_HEADER in row:
                    id_col, name_col = row.index(ID_HEADER), row.index(NAME_HEADER)
                    entries = [
                        # 两个字的名字常被敲成「张　三」，中间的空格去掉
                        (r[id_col], "".join(r[name_col].split()))
                        for r in rows[index + 1 :]
                        if len(r) > max(id_col, name_col) and r[id_col] and r[name_col]
                    ]
                    return Roster(entries)
    finally:
        workbook.close()
    raise RosterError(
        f"名单「{Path(path).name}」里没有找到表头为「{ID_HEADER}」和「{NAME_HEADER}」的两列"
    )
