"""用真实的 38 个文件做的检查。正确答案来自人工逐份查看（见 文档/2-设计方案.md 附录）。"""

import pytest

from conftest import REAL_INPUT, REAL_ROSTER, needs_real_roster, needs_real_samples
from homework_formatter.converter import find_soffice
from homework_formatter.identity import Source
from homework_formatter.inventory import Outcome, take_inventory
from homework_formatter.roster import load_roster
from homework_formatter.scanner import fingerprint

pytestmark = [
    needs_real_samples,
    pytest.mark.skipif(find_soffice() is None, reason="没有 LibreOffice，读不了旧版 .doc"),
]

# 原文件名 -> (姓名, 学号)，不用名单时应当认出的结果
EXPECTED = {
    "20240302周子衿.docx": ("周子衿", "20240302"),
    "20240304许嘉树.docx": ("许嘉树", "20240304"),
    "20240308-白鹿鸣.docx": ("白鹿鸣", "20240308"),
    "20240322-乔木.docx": ("乔木", "20240322"),
    "20240330-骆一鸣.docx": ("骆一鸣", "20240330"),
    "20240334-祝余.docx": ("祝余", "20240334"),
    "report_qinlang.docx": ("秦朗", "20240319"),
    "~$林知夏.docx": (None, None),
    "何其多.docx": ("何其多", "20240318"),
    "作业_柏舟_20240326.docx": ("柏舟", "20240326"),
    "作业_梁思齐_20240324.docx": ("梁思齐", "20240324"),
    "作业_梁思齐_20240324(1).docx": ("梁思齐", "20240324"),
    "关山月改.docx": ("关山月", "20240331"),
    "叶知秋的报告.docx": ("叶知秋", None),
    "唐小满改.doc": ("唐小满", "20240307"),
    "宋清和20240312.docx": ("宋清和", "20240312"),
    "尚可.docx": ("尚可", None),
    "文档1.docx": ("池鱼", "20240335"),
    "新建 Microsoft Word 文档 (2).docx": ("殷实", "20240329"),
    "新建 Microsoft Word 文档.docx": (None, None),
    "方可.docx": ("方可", "20240314"),
    "景行20240336.docx": ("景行", "20240336"),
    "林知夏.docx": ("林知夏", "20240301"),
    "江晚晴20240309.docx": ("江晚晴", "20240309"),
    "沈星河.docx": ("沈星河", "20240305"),
    "温言20240317.docx": ("温言", "20240317"),
    "研究性学习 孟想.docx": ("孟想", "20240320"),
    "研究性学习 苏禾.docx": ("苏禾", "20240315"),
    "研究性学习 齐悦.pdf": ("齐悦", None),
    "程予安20240310.docx": ("程予安", "20240310"),
    "邱实.doc": ("邱实", "20240327"),
    "阮青（最终版）.docx": ("阮青", "20240323"),
    "陆景明-研究性学习报告.docx": ("陆景明", "20240316"),
    "陈一诺20240303.docx": ("陈一诺", "20240303"),
    "顾念20240306.docx": ("顾念", "20240306"),
    "高二3班 夏以宁 小报告.docx": ("夏以宁", "20240311"),
    "高二3班 安然 小报告.docx": ("安然", "20240325"),
    "高二3班 谷雨 小报告.docx": ("谷雨", "20240328"),
}

# 名单能补上的学号
FILLED_BY_ROSTER = {
    "叶知秋的报告.docx": "20240313",
    "尚可.docx": "20240333",
    "研究性学习 齐悦.pdf": "20240332",
}

NOT_READY = {
    "~$林知夏.docx": Outcome.IGNORED,
    "研究性学习 齐悦.pdf": Outcome.UNPROCESSED,
    "尚可.docx": Outcome.UNPROCESSED,
    "作业_梁思齐_20240324.docx": Outcome.REVIEW,
    "作业_梁思齐_20240324(1).docx": Outcome.REVIEW,
    "新建 Microsoft Word 文档.docx": Outcome.REVIEW,
    "高二3班 安然 小报告.docx": Outcome.REVIEW,  # 没有题目
}


@pytest.fixture(scope="module")
def without_roster():
    return take_inventory(REAL_INPUT, None)


@pytest.fixture(scope="module")
def with_roster():
    if not REAL_ROSTER.is_file():
        pytest.skip("没有「名单.xlsx」")
    return take_inventory(REAL_INPUT, load_roster(REAL_ROSTER))


def identities(inventory):
    return {r.source.path.name: (r.name, r.student_id) for r in inventory.records}


def outcomes(inventory):
    return {r.source.path.name: r.outcome for r in inventory.records}


def test_every_file_is_identified_exactly_as_checked_by_hand(without_roster):
    assert identities(without_roster) == EXPECTED


def test_outcomes_without_roster(without_roster):
    expected = {name: Outcome.READY for name in EXPECTED}
    expected.update(NOT_READY)
    expected["叶知秋的报告.docx"] = Outcome.REVIEW  # 没有名单就查不到学号

    assert outcomes(without_roster) == expected


@needs_real_roster
def test_roster_fills_in_the_missing_ids_and_changes_nothing_else(with_roster):
    expected = dict(EXPECTED)
    for filename, student_id in FILLED_BY_ROSTER.items():
        expected[filename] = (expected[filename][0], student_id)

    assert identities(with_roster) == expected
    filled = {r.source.path.name for r in with_roster.records if r.id_source is Source.ROSTER}
    assert filled == set(FILLED_BY_ROSTER)


@needs_real_roster
def test_outcomes_with_roster(with_roster):
    expected = {name: Outcome.READY for name in EXPECTED}
    expected.update(NOT_READY)

    assert outcomes(with_roster) == expected
    assert with_roster.count(Outcome.READY) == 31
    assert with_roster.count(Outcome.REVIEW) == 4
    assert with_roster.count(Outcome.UNPROCESSED) == 2
    assert with_roster.count(Outcome.IGNORED) == 1


@needs_real_roster
def test_the_only_student_without_any_file(with_roster):
    assert with_roster.missing_students == [("20240321", "童年")]


@needs_real_roster
def test_the_two_versions_from_the_same_student_show_what_differs(with_roster):
    notes = {
        r.source.path.name: "\n".join(r.notes)
        for r in with_roster.records
        if r.student_id == "20240324"
    }

    assert "9 天" in notes["作业_梁思齐_20240324.docx"]
    assert "10 天" in notes["作业_梁思齐_20240324.docx"]
    assert "作业_梁思齐_20240324(1).docx" in notes["作业_梁思齐_20240324.docx"]


@needs_real_roster
def test_where_each_identity_came_from(with_roster):
    sources = {r.source.path.name: (r.name_source, r.id_source) for r in with_roster.records}

    assert sources["20240302周子衿.docx"] == (Source.HEADER, Source.HEADER)
    assert sources["陆景明-研究性学习报告.docx"] == (Source.HEADER, Source.HEADER)
    assert sources["report_qinlang.docx"] == (Source.BODY, Source.BODY)
    assert sources["唐小满改.doc"] == (Source.BODY, Source.BODY)
    assert sources["叶知秋的报告.docx"] == (Source.FILENAME, Source.ROSTER)


@needs_real_roster
def test_originals_are_untouched_by_taking_inventory():
    before = fingerprint(REAL_INPUT)

    take_inventory(REAL_INPUT, load_roster(REAL_ROSTER))

    assert fingerprint(REAL_INPUT) == before
    assert len(before) == 37  # 38 个文件里有 1 个是 Word 的临时文件，不算原件


# ---------- 第 2 步：认结构 ----------

from homework_formatter.structure import analyze  # noqa: E402

LIGHTS = "教室的灯，白天其实可以少开一半"
BIKES = "自行车棚为什么总是停不下"
DUTY = "值日表上，为什么总有人忘"
CANTEEN = "食堂排队太久，我想这样解决"
WATER = "课间十分钟，饮水机前的长队"
BOOKS = "图书角的书，为什么总是有借无还"

SECTIONS = ["我发现的问题", "原因分析", "我的办法"]
NUMBERING = {
    "一、": ["一、", "二、", "三、"],
    "1.": ["1. ", "2. ", "3. "],
    "（一）": ["（一）", "（二）", "（三）"],
    "": ["", "", ""],
}

# 原文件名 -> (题目, 小标题的编号写法, 有没有"参考资料")，人工逐份看过
EXPECTED_STRUCTURE = {
    "20240302周子衿.docx": (LIGHTS, "1.", True),
    "20240304许嘉树.docx": (BIKES, "", False),
    "20240308-白鹿鸣.docx": (LIGHTS, "一、", False),
    "20240322-乔木.docx": (BIKES, "一、", False),
    "20240330-骆一鸣.docx": (DUTY, "（一）", True),
    "20240334-祝余.docx": (BIKES, "一、", True),
    "report_qinlang.docx": (CANTEEN, "一、", False),
    "何其多.docx": (DUTY, "一、", False),
    "作业_柏舟_20240326.docx": (LIGHTS, "一、", True),
    "作业_梁思齐_20240324.docx": (DUTY, "1.", False),
    "作业_梁思齐_20240324(1).docx": (DUTY, "1.", False),
    "关山月改.docx": (CANTEEN, "（一）", False),
    "叶知秋的报告.docx": (CANTEEN, "一、", True),
    "唐小满改.doc": (CANTEEN, "一、", False),
    "宋清和20240312.docx": (DUTY, "（一）", False),
    "文档1.docx": (WATER, "", False),
    "新建 Microsoft Word 文档 (2).docx": (WATER, "一、", True),
    "新建 Microsoft Word 文档.docx": (BOOKS, "一、", False),
    "方可.docx": (LIGHTS, "一、", True),
    "景行20240336.docx": (DUTY, "一、", False),
    "林知夏.docx": (CANTEEN, "1.", True),
    "江晚晴20240309.docx": (BOOKS, "一、", False),
    "沈星河.docx": (WATER, "", False),
    "温言20240317.docx": (WATER, "1.", False),
    "研究性学习 孟想.docx": (LIGHTS, "一、", False),
    "研究性学习 苏禾.docx": (BOOKS, "（一）", False),
    "程予安20240310.docx": (BIKES, "一、", False),
    "邱实.doc": (BOOKS, "（一）", False),
    "阮青（最终版）.docx": (WATER, "一、", False),
    "陆景明-研究性学习报告.docx": (BIKES, "1.", True),
    "陈一诺20240303.docx": (BOOKS, "1.", True),
    "顾念20240306.docx": (DUTY, "一、", False),
    "高二3班 夏以宁 小报告.docx": (WATER, "一、", True),
    "高二3班 安然 小报告.docx": (None, "", False),
    "高二3班 谷雨 小报告.docx": (BIKES, "1.", True),
}


def expected_headings(numbering, has_references):
    headings = [prefix + name for prefix, name in zip(NUMBERING[numbering], SECTIONS)]
    return headings + (["参考资料"] if has_references else [])


def test_title_and_headings_of_every_assignment_match_what_a_person_sees(without_roster):
    found = {
        r.source.path.name: analyze(r.docx_bytes) for r in without_roster.records if r.has_content
    }

    assert set(found) == set(EXPECTED_STRUCTURE)
    wrong = {}
    for filename, (title, numbering, has_references) in EXPECTED_STRUCTURE.items():
        expected = (title, expected_headings(numbering, has_references))
        actual = (found[filename].title, found[filename].headings)
        if actual != expected:
            wrong[filename] = actual
    assert wrong == {}
