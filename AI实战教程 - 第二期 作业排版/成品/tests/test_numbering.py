import pytest

from homework_formatter.numbering import chinese_numeral, plan_numbering, split_number
from homework_formatter.structure import NUMBERED_PREFIX

UNNUMBERED = ("参考资料", "参考文献", "致谢", "附录")
SECTIONS = ["我发现的问题", "原因分析", "我的办法"]
UNIFIED = ["一、我发现的问题", "二、原因分析", "三、我的办法"]


@pytest.mark.parametrize(
    "text, style, core",
    [
        ("一、我发现的问题", "一、", "我发现的问题"),
        ("二．原因分析", "一、", "原因分析"),
        ("三. 我的办法", "一、", "我的办法"),
        ("1. 我发现的问题", "1.", "我发现的问题"),
        ("2、原因分析", "1.", "原因分析"),
        ("３．我的办法", "1.", "我的办法"),
        ("12.办法", "1.", "办法"),
        ("（一）我发现的问题", "（一）", "我发现的问题"),
        ("(二) 原因分析", "（一）", "原因分析"),
        ("（3）我的办法", "（一）", "我的办法"),
        ("第一部分 我发现的问题", "第一部分", "我发现的问题"),
        ("第二章　原因分析", "第一部分", "原因分析"),
        ("我发现的问题", None, "我发现的问题"),
        ("参考资料", None, "参考资料"),
        ("一些想法", None, "一些想法"),
        ("2024 年的情况", None, "2024 年的情况"),
        ("3.5 倍", None, "3.5 倍"),
    ],
)
def test_recognises_how_a_heading_is_numbered(text, style, core):
    assert split_number(text) == (style, core)


@pytest.mark.parametrize("text", ["一、问题", "（一）问题", "(1) 问题", "1. 问题", "2、问题", "第一部分 问题"])
def test_agrees_with_the_structure_module_about_what_counts_as_numbered(text):
    assert NUMBERED_PREFIX.match(text)
    assert split_number(text)[0] is not None


@pytest.mark.parametrize(
    "number, text",
    [(1, "一"), (3, "三"), (10, "十"), (11, "十一"), (19, "十九"), (20, "二十"), (21, "二十一"), (99, "九十九")],
)
def test_chinese_numerals(number, text):
    assert chinese_numeral(number) == text


@pytest.mark.parametrize(
    "prefixes",
    [["一、", "二、", "三、"], ["1. ", "2. ", "3. "], ["（一）", "（二）", "（三）"], ["", "", ""]],
)
def test_all_four_ways_students_wrote_become_the_same(prefixes):
    headings = [p + s for p, s in zip(prefixes, SECTIONS)]

    plan = plan_numbering(headings, UNNUMBERED)

    assert plan.headings == UNIFIED
    assert plan.problem is None


def test_already_unified_headings_need_no_change():
    plan = plan_numbering(UNIFIED + ["参考资料"], UNNUMBERED)

    assert plan.headings == UNIFIED + ["参考资料"]
    assert plan.note is None


@pytest.mark.parametrize(
    "prefixes, expected",
    [(["1. ", "2. ", "3. "], "1."), (["（一）", "（二）", "（三）"], "（一）"), (["", "", ""], "没有编号")],
)
def test_the_note_says_what_the_numbering_used_to_be(prefixes, expected):
    plan = plan_numbering([p + s for p, s in zip(prefixes, SECTIONS)], UNNUMBERED)

    assert "一、二、三" in plan.note and expected in plan.note


def test_references_and_the_like_are_never_numbered():
    plan = plan_numbering(["1. 问题", "2. 办法", "参考资料", "致谢："], UNNUMBERED)

    assert plan.headings == ["一、问题", "二、办法", "参考资料", "致谢："]


def test_a_number_in_front_of_references_is_removed_for_consistency():
    plan = plan_numbering(["一、问题", "二、办法", "三、参考资料"], UNNUMBERED)

    assert plan.headings == ["一、问题", "二、办法", "参考资料"]
    assert "参考资料" in plan.note


def test_a_skipped_number_is_closed_up():
    plan = plan_numbering(["一、问题", "三、原因", "四、办法"], UNNUMBERED)

    assert plan.headings == ["一、问题", "二、原因", "三、办法"]
    assert plan.note is not None


@pytest.mark.parametrize(
    "headings",
    [
        # 大标题和节内的小条目用了同一种写法：编号会往回数
        ["1. 背景", "2. 调查方法", "1. 问卷调查", "2. 个别访谈", "3. 结论"],
        ["一、问题", "三、原因", "三、办法"],
        ["2. 原因", "3. 办法"],  # 不从 1 开始，多半前面还有没认出来的
    ],
)
def test_numbers_that_go_backwards_or_do_not_start_at_one_are_left_alone(headings):
    plan = plan_numbering(headings, UNNUMBERED)

    assert plan.headings is None
    assert "不是从 1 开始" in plan.problem


@pytest.mark.parametrize(
    "headings",
    [
        ["一、二年级的情况", "三、四年级的情况", "五、六年级的情况"],
        ["七、八年级的对比"],
        ["（3）班同学的看法"],
    ],
)
def test_a_heading_that_merely_starts_with_a_number_is_not_renumbered(headings):
    assert plan_numbering(headings, UNNUMBERED).headings is None


def test_consecutive_numbers_are_trusted_even_when_the_words_start_with_a_number():
    plan = plan_numbering(["一、一些想法", "二、三点建议", "三、四月份的数据"], UNNUMBERED)

    assert plan.headings == ["一、一些想法", "二、三点建议", "三、四月份的数据"]


@pytest.mark.parametrize(
    "text, value",
    [("一、问题", 1), ("十、问题", 10), ("十二、问题", 12), ("二十一、问题", 21), ("3. 问题", 3),
     ("（四）问题", 4), ("(12) 问题", 12), ("１２．问题", 12), ("问题", None)],
)
def test_reads_the_value_of_a_heading_number(text, value):
    from homework_formatter.numbering import number_value

    assert number_value(text) == value


def test_more_than_ten_headings():
    plan = plan_numbering([f"{n}. 第{n}节的内容" for n in range(1, 13)], UNNUMBERED)

    assert plan.headings[9:] == ["十、第10节的内容", "十一、第11节的内容", "十二、第12节的内容"]


@pytest.mark.parametrize(
    "headings, reason",
    [
        (["一、问题", "1. 现象", "2. 数据", "二、办法"], "两种"),  # 大标题套小标题
        (["一、问题", "（一）现象", "二、办法"], "两种"),
        (["一、问题", "原因分析", "三、办法"], "有的"),  # 有的编了，有的没编
        (["第一部分 问题", "第二部分 办法"], "第一部分"),
    ],
)
def test_when_unsure_nothing_is_changed_and_the_reason_is_given(headings, reason):
    plan = plan_numbering(headings, UNNUMBERED)

    assert plan.headings is None
    assert reason in plan.problem


def test_no_headings_at_all_is_fine():
    plan = plan_numbering([], UNNUMBERED)

    assert plan.headings == [] and plan.problem is None and plan.note is None
