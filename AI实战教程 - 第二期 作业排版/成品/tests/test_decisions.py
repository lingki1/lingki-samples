import openpyxl
import pytest

from conftest import make_docx
from homework_formatter.decisions import Decisions, DecisionsError, load_decisions
from homework_formatter.identity import Source
from homework_formatter.inventory import Outcome, take_inventory
from homework_formatter.roster import Roster

ROSTER = Roster([("20240301", "张三"), ("20240302", "李四"), ("20240303", "王五")])
BODY = ["一、问题", "这是正文的第一段，写了一些内容。"]
HEADER = ("原文件名", "姓名", "学号", "不要这份", "备注")


def make_decisions(path, rows, header=HEADER):
    workbook = openpyxl.Workbook()
    workbook.active.append(list(header))
    for row in rows:
        workbook.active.append(list(row))
    workbook.save(str(path))
    return path


def by_name(inventory):
    return {r.source.path.name: r for r in inventory.records}


def test_loads_assignments_and_discards(tmp_path):
    path = make_decisions(
        tmp_path / "老师的决定.xlsx",
        [
            ("新建 Microsoft Word 文档.docx", "王五", 20240303, None, "问过本人"),
            ("张三.docx", None, None, "是", "以带 (1) 的为准"),
            (None, None, None, None, None),
        ],
    )

    decisions = load_decisions(path)

    assert decisions.assigned == {"新建 Microsoft Word 文档.docx": ("王五", "20240303")}
    assert decisions.discarded == {"张三.docx": "以带 (1) 的为准"}


def test_a_row_that_decides_nothing_is_an_error(tmp_path):
    path = make_decisions(tmp_path / "d.xlsx", [("a.docx", None, None, None, "忘了填")])

    with pytest.raises(DecisionsError, match="第 2 行"):
        load_decisions(path)


def test_a_row_that_decides_both_ways_is_an_error(tmp_path):
    path = make_decisions(tmp_path / "d.xlsx", [("a.docx", "张三", "20240301", "是", None)])

    with pytest.raises(DecisionsError, match="第 2 行"):
        load_decisions(path)


def test_missing_columns_are_reported_clearly(tmp_path):
    path = make_decisions(tmp_path / "d.xlsx", [("a.docx", "x")], header=("文件", "名字"))

    with pytest.raises(DecisionsError, match="原文件名"):
        load_decisions(path)


def test_teacher_can_say_whose_an_anonymous_file_is(tmp_path):
    make_docx(tmp_path / "新建 Microsoft Word 文档.docx", ["题目", *BODY])
    decisions = Decisions(assigned={"新建 Microsoft Word 文档.docx": ("王五", "20240303")})

    inventory = take_inventory(tmp_path, ROSTER, decisions=decisions)
    (record,) = inventory.records

    assert record.outcome is Outcome.READY
    assert (record.name, record.student_id) == ("王五", "20240303")
    assert (record.name_source, record.id_source) == (Source.TEACHER, Source.TEACHER)
    assert record.new_filename == "20240303-王五.docx"
    assert inventory.missing_students == [("20240301", "张三"), ("20240302", "李四")]


def test_teacher_gives_only_the_name_and_roster_supplies_the_id(tmp_path):
    make_docx(tmp_path / "x.docx", ["题目", *BODY])
    decisions = Decisions(assigned={"x.docx": ("王五", None)})

    (record,) = take_inventory(tmp_path, ROSTER, decisions=decisions).records

    assert (record.student_id, record.id_source) == ("20240303", Source.ROSTER)
    assert record.outcome is Outcome.READY


def test_teacher_assignment_that_contradicts_the_roster_still_needs_review(tmp_path):
    make_docx(tmp_path / "x.docx", ["题目", *BODY])
    decisions = Decisions(assigned={"x.docx": ("王五", "20240301")})

    (record,) = take_inventory(tmp_path, ROSTER, decisions=decisions).records

    assert record.outcome is Outcome.REVIEW
    assert any("对不上" in n for n in record.notes)


def test_discarding_one_of_two_submissions_makes_the_other_ready(tmp_path):
    info = "高二（3）班  张三  20240301"
    make_docx(tmp_path / "张三.docx", ["题目", info, "其中有 9 天有人忘了。"])
    make_docx(tmp_path / "张三(1).docx", ["题目", info, "其中有 10 天有人忘了。"])
    decisions = Decisions(discarded={"张三.docx": "以带 (1) 的为准"})

    records = by_name(take_inventory(tmp_path, ROSTER, decisions=decisions))

    assert records["张三.docx"].outcome is Outcome.IGNORED
    assert any("老师决定不要这份" in n and "以带 (1) 的为准" in n for n in records["张三.docx"].notes)
    assert records["张三.docx"].new_filename is None
    assert records["张三(1).docx"].outcome is Outcome.READY
    assert records["张三(1).docx"].notes == []


def test_discarding_a_students_only_file_lists_them_as_missing(tmp_path):
    make_docx(tmp_path / "张三.docx", ["题目", "张三 20240301", *BODY])
    decisions = Decisions(discarded={"张三.docx": ""})

    inventory = take_inventory(tmp_path, Roster([("20240301", "张三")]), decisions=decisions)

    assert inventory.missing_students == [("20240301", "张三")]


def test_decisions_about_files_that_do_not_exist_are_reported(tmp_path):
    make_docx(tmp_path / "a.docx", ["题目", "张三 20240301", *BODY])
    decisions = Decisions(
        assigned={"打错的名字.docx": ("李四", "20240302")}, discarded={"也不存在.docx": ""}
    )

    inventory = take_inventory(tmp_path, ROSTER, decisions=decisions)

    assert inventory.unmatched_decisions == ["也不存在.docx", "打错的名字.docx"]
