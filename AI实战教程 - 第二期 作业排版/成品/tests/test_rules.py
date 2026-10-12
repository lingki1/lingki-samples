import pytest

from conftest import PROJECT_ROOT
from homework_formatter.rules import RulesError, load_rules

RULES_FILE = PROJECT_ROOT / "规则.toml"


@pytest.fixture(scope="module")
def rules():
    return load_rules(RULES_FILE)


# 下面几条对着《格式要求.docx》逐条写：规则表被改错了会在这里发现


def test_rule_1_page(rules):
    assert (rules.page_width_twips, rules.page_height_twips) == (11906, 16838)  # A4
    assert rules.margin_top_bottom_twips == 1440  # 2.54 厘米
    assert rules.margin_left_right_twips == 1803  # 3.18 厘米
    assert rules.text_width_twips == 11906 - 2 * 1803


def test_rule_2_title(rules):
    assert (rules.title.east_asian_font, rules.title.size_pt) == ("黑体", 22)  # 二号
    assert rules.title.alignment == "center"
    assert rules.title.first_line_chars == 0


def test_rule_3_info_line(rules):
    assert (rules.info.east_asian_font, rules.info.size_pt) == ("楷体", 12)  # 小四
    assert rules.info.alignment == "center"
    assert rules.info_line("张三", "20240301") == "高二（3）班  张三  20240301"


def test_rule_4_heading(rules):
    assert (rules.heading.east_asian_font, rules.heading.size_pt) == ("黑体", 14)  # 四号
    assert rules.heading.alignment == "left"
    assert rules.heading.first_line_chars == 0


def test_teachers_addition_headings_are_renumbered_except_references(rules):
    assert rules.unify_heading_numbers is True
    assert rules.unnumbered_headings == ("参考资料", "参考文献", "致谢", "附录")


def test_rule_5_body(rules):
    assert rules.body.east_asian_font == "宋体"
    assert rules.body.western_font == "Times New Roman"
    assert rules.body.size_pt == 12  # 小四
    assert rules.body.alignment == "both"
    assert rules.body.first_line_chars == 2
    assert rules.line_spacing == 1.5


def test_rule_6_table_text_follows_the_body_font(rules):
    assert rules.table_text.east_asian_font == "宋体"
    assert rules.table_text.size_pt == 12
    assert rules.table_text.alignment == "center"
    assert rules.table_text.first_line_chars == 0


def test_rule_7_page_number(rules):
    assert rules.page_number.alignment == "center"
    assert rules.page_number.size_pt == 9  # 小五


def test_digits_in_title_heading_and_info_use_the_same_font_as_the_chinese(rules):
    for style in (rules.title, rules.heading, rules.info):
        assert style.western_font == style.east_asian_font


def write(tmp_path, text):
    path = tmp_path / "规则.toml"
    path.write_text(text, encoding="utf-8")
    return path


def test_size_can_be_given_in_points(tmp_path):
    text = RULES_FILE.read_text(encoding="utf-8").replace('"字号" = "二号"', '"字号" = 20')

    assert load_rules(write(tmp_path, text)).title.size_pt == 20


def test_unknown_size_name_is_explained(tmp_path):
    text = RULES_FILE.read_text(encoding="utf-8").replace('"字号" = "二号"', '"字号" = "特大号"')

    with pytest.raises(RulesError, match="特大号"):
        load_rules(write(tmp_path, text))


def test_unknown_alignment_is_explained(tmp_path):
    text = RULES_FILE.read_text(encoding="utf-8").replace('"对齐" = "两端对齐"', '"对齐" = "随便"')

    with pytest.raises(RulesError, match="随便"):
        load_rules(write(tmp_path, text))


def test_missing_entry_is_explained(tmp_path):
    text = RULES_FILE.read_text(encoding="utf-8").replace('"首行缩进_字" = 2', "")

    with pytest.raises(RulesError, match="首行缩进_字"):
        load_rules(write(tmp_path, text))


def test_broken_file_is_explained(tmp_path):
    with pytest.raises(RulesError, match="规则"):
        load_rules(write(tmp_path, "[页面\n纸张宽 = "))
