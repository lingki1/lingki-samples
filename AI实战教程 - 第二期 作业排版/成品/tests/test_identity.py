import pytest

from homework_formatter.identity import (
    Source,
    is_info_line,
    name_from_filename,
    parse_info_line,
    resolve_identity,
)
from homework_formatter.roster import Roster

ROSTER = Roster(
    [
        ("20240301", "林知夏"),
        ("20240304", "许嘉树"),
        ("20240308", "白鹿鸣"),
        ("20240313", "叶知秋"),
        ("20240322", "乔木"),
    ]
)


# 学生实际用过的五种写法
@pytest.mark.parametrize(
    "line, name, student_id",
    [
        ("高二（3）班  林知夏  20240301", "林知夏", "20240301"),
        ("许嘉树 20240304", "许嘉树", "20240304"),
        ("姓名：白鹿鸣    学号：20240308", "白鹿鸣", "20240308"),
        ("班级：高二3班 姓名:乔木 学号:20240322", "乔木", "20240322"),
        ("班级：高二(3)班　姓名：乔木", "乔木", None),
        ("高二3班张三20240301", "张三", "20240301"),
        ("张三（20240301）", "张三", "20240301"),
        ("高二三班 张三 20240301", "张三", "20240301"),
        ("高二年级（3）班 姓名：张三 学号：20240301", "张三", "20240301"),
        ("学号20240301 姓名张三", "张三", "20240301"),
    ],
)
def test_parses_every_way_students_wrote_the_info_line(line, name, student_id):
    assert is_info_line(line)
    assert parse_info_line(line) == (name, student_id)
    assert parse_info_line(line, ROSTER) == (name, student_id)


@pytest.mark.parametrize(
    "line",
    [
        "教室的灯，白天其实可以少开一半",
        "一、我发现的问题",
        "我数了 24 天，其中有 11 天是晴天，这些天里靠窗的六盏灯等于白白亮了一整天。",
        "",
    ],
)
def test_titles_headings_and_body_text_are_not_info_lines(line):
    assert not is_info_line(line)


@pytest.mark.parametrize(
    "filename, expected",
    [
        ("叶知秋的报告.docx", "叶知秋"),
        ("唐小满改.doc", "唐小满"),
        ("阮青（最终版）.docx", "阮青"),
        ("高二3班 夏以宁 小报告.docx", "夏以宁"),
        ("研究性学习 孟想.docx", "孟想"),
        ("作业_梁思齐_20240324(1).docx", "梁思齐"),
        ("20240308-白鹿鸣.docx", "白鹿鸣"),
        ("陆景明-研究性学习报告.docx", "陆景明"),
        ("新建 Microsoft Word 文档 (2).docx", None),
        ("文档1.docx", None),
        ("report_qinlang.docx", None),
    ],
)
def test_guesses_name_from_filename_without_a_roster(filename, expected):
    assert name_from_filename(filename) == expected


def test_roster_name_in_filename_wins_over_guessing():
    assert name_from_filename("叶知秋的报告.docx", ROSTER) == "叶知秋"
    assert name_from_filename("不在名单的人.docx", ROSTER) is None


def test_body_is_preferred_and_sources_are_recorded():
    identity = resolve_identity(
        filename="report.docx",
        body_lines=["题目", "许嘉树 20240304", "正文"],
        header_lines=[],
    )

    assert (identity.name, identity.student_id) == ("许嘉树", "20240304")
    assert (identity.name_source, identity.id_source) == (Source.BODY, Source.BODY)
    assert identity.problems == []


def test_falls_back_to_header_then_filename():
    from_header = resolve_identity(
        filename="x.docx",
        body_lines=["题目", "正文正文正文"],
        header_lines=["高二（3）班  林知夏  20240301"],
    )
    from_filename = resolve_identity(
        filename="20240308-白鹿鸣.docx", body_lines=["题目"], header_lines=[]
    )

    assert (from_header.name, from_header.name_source) == ("林知夏", Source.HEADER)
    assert (from_header.student_id, from_header.id_source) == ("20240301", Source.HEADER)
    assert (from_filename.name, from_filename.name_source) == ("白鹿鸣", Source.FILENAME)
    assert (from_filename.student_id, from_filename.id_source) == ("20240308", Source.FILENAME)


def test_only_the_first_few_paragraphs_are_searched():
    body = ["题目", "一、问题", "正文一", "正文二", "正文三", "正文四", "许嘉树 20240304"]

    identity = resolve_identity(filename="x.docx", body_lines=body, header_lines=[])

    assert identity.name is None and identity.student_id is None


def test_roster_fills_in_a_missing_id():
    identity = resolve_identity(
        filename="叶知秋的报告.docx", body_lines=["题目", "正文"], header_lines=[], roster=ROSTER
    )

    assert (identity.name, identity.name_source) == ("叶知秋", Source.FILENAME)
    assert (identity.student_id, identity.id_source) == ("20240313", Source.ROSTER)
    assert identity.problems == []


def test_roster_fills_in_a_missing_name():
    identity = resolve_identity(
        filename="x.docx", body_lines=["题目", "学号：20240322"], header_lines=[], roster=ROSTER
    )

    assert (identity.name, identity.name_source) == ("乔木", Source.ROSTER)


def test_nothing_found_means_nothing_guessed():
    identity = resolve_identity(
        filename="新建 Microsoft Word 文档.docx",
        body_lines=["题目", "正文"],
        header_lines=[],
        roster=ROSTER,
    )

    assert identity.name is None and identity.student_id is None


def test_filename_and_body_disagreeing_is_reported():
    identity = resolve_identity(
        filename="20240399-许嘉树.docx", body_lines=["题目", "许嘉树 20240304"], header_lines=[]
    )

    assert identity.student_id == "20240304"
    assert any("20240399" in p for p in identity.problems)


def test_name_and_id_not_matching_the_roster_is_reported():
    identity = resolve_identity(
        filename="x.docx", body_lines=["题目", "许嘉树 20240301"], header_lines=[], roster=ROSTER
    )

    assert any("名单" in p for p in identity.problems)


def test_id_missing_from_the_roster_is_reported():
    identity = resolve_identity(
        filename="x.docx", body_lines=["题目", "某某某 20249999"], header_lines=[], roster=ROSTER
    )

    assert any("名单" in p for p in identity.problems)
