import pytest

from conftest import make_roster
from homework_formatter.roster import Roster, RosterError, load_roster


def test_loads_ids_and_names(tmp_path):
    path = make_roster(tmp_path / "名单.xlsx", [("20240301", "张三"), ("20240302", "李四")])

    roster = load_roster(path)

    assert roster.name_of("20240301") == "张三"
    assert roster.id_of("李四") == "20240302"
    assert len(roster) == 2


def test_ids_typed_as_numbers_in_excel_still_work(tmp_path):
    path = make_roster(tmp_path / "名单.xlsx", [(20240301, "张三"), (20240302.0, " 李四 ")])

    roster = load_roster(path)

    assert roster.id_of("张三") == "20240301"
    assert roster.id_of("李四") == "20240302"


def test_column_order_and_blank_rows_do_not_matter(tmp_path):
    path = make_roster(
        tmp_path / "名单.xlsx",
        [("张三", "20240301"), (None, None), ("李四", "20240302")],
        header=("姓名", "学号"),
    )

    assert load_roster(path).id_of("李四") == "20240302"


def test_missing_columns_are_reported_clearly(tmp_path):
    path = make_roster(tmp_path / "名单.xlsx", [("a", "b")], header=("编号", "名字"))

    with pytest.raises(RosterError, match="学号"):
        load_roster(path)


def test_two_students_with_the_same_name_cannot_be_looked_up_by_name():
    roster = Roster([("20240301", "张三"), ("20240302", "张三")])

    assert roster.id_of("张三") is None
    assert roster.name_of("20240302") == "张三"


def test_finds_roster_names_inside_a_piece_of_text():
    roster = Roster([("20240301", "张三"), ("20240302", "李四")])

    assert roster.names_in("高二3班 张三 小报告") == ["张三"]
    assert roster.names_in("新建 Microsoft Word 文档") == []
