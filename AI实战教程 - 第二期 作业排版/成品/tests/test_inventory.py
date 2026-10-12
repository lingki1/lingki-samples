from conftest import make_docx
from homework_formatter.converter import ConversionResult
from homework_formatter.identity import Source
from homework_formatter.inventory import Outcome, take_inventory
from homework_formatter.roster import Roster

ROSTER = Roster([("20240301", "张三"), ("20240302", "李四"), ("20240303", "王五")])
BODY = ["一、问题", "这是正文的第一段，写了一些内容。", "这是正文的第二段。"]


def no_conversion(paths):
    return ConversionResult(converted={}, failed={p: "测试里不转换" for p in paths})


def by_name(inventory):
    return {r.source.path.name: r for r in inventory.records}


def test_a_normal_assignment_is_ready(tmp_path):
    make_docx(tmp_path / "随便.docx", ["题目", "高二（3）班  张三  20240301", *BODY])

    (record,) = take_inventory(tmp_path, ROSTER).records

    assert record.outcome is Outcome.READY
    assert (record.name, record.student_id) == ("张三", "20240301")
    assert record.new_filename == "20240301-张三.docx"
    assert record.notes == []


def test_title_and_headings_are_recorded(tmp_path):
    make_docx(tmp_path / "a.docx", ["我的题目", "高二（3）班  张三  20240301", *BODY])

    (record,) = take_inventory(tmp_path, ROSTER).records

    assert record.structure.title == "我的题目"
    assert record.structure.headings == ["一、问题"]


def test_assignment_without_a_title_needs_review(tmp_path):
    make_docx(tmp_path / "a.docx", ["高二（3）班  张三  20240301", *BODY])

    (record,) = take_inventory(tmp_path, ROSTER).records

    assert record.outcome is Outcome.REVIEW
    assert any("题目" in n for n in record.notes)
    assert (record.name, record.student_id) == ("张三", "20240301")


def test_identity_written_only_in_the_header_is_found(tmp_path):
    make_docx(tmp_path / "x.docx", ["题目", *BODY], header="高二（3）班  李四  20240302")

    (record,) = take_inventory(tmp_path, ROSTER).records

    assert record.outcome is Outcome.READY
    assert (record.name, record.name_source) == ("李四", Source.HEADER)


def test_anonymous_assignment_needs_review_and_gets_no_new_name(tmp_path):
    make_docx(tmp_path / "新建 Microsoft Word 文档.docx", ["题目", *BODY])

    (record,) = take_inventory(tmp_path, ROSTER).records

    assert record.outcome is Outcome.REVIEW
    assert record.new_filename is None
    assert any("姓名" in n and "学号" in n for n in record.notes)


def test_blank_document_is_not_processed_but_owner_is_still_identified(tmp_path):
    make_docx(tmp_path / "王五.docx", [])

    (record,) = take_inventory(tmp_path, ROSTER).records

    assert record.outcome is Outcome.UNPROCESSED
    assert (record.name, record.student_id) == ("王五", "20240303")
    assert any("空白" in n for n in record.notes)
    # 不会生成文件，就不该显示新文件名
    assert record.new_filename is None


def test_pdf_temp_and_unknown_files(tmp_path):
    (tmp_path / "研究性学习 李四.pdf").write_bytes(b"%PDF-1.7")
    (tmp_path / "~$张三.docx").write_bytes(b"\x05x" + b"\0" * 100)
    (tmp_path / "照片.jpg").write_bytes(b"\xff\xd8")

    records = by_name(take_inventory(tmp_path, ROSTER))

    assert records["研究性学习 李四.pdf"].outcome is Outcome.UNPROCESSED
    assert records["研究性学习 李四.pdf"].student_id == "20240302"
    assert any("PDF" in n for n in records["研究性学习 李四.pdf"].notes)
    assert records["~$张三.docx"].outcome is Outcome.IGNORED
    assert records["~$张三.docx"].name is None
    assert records["照片.jpg"].outcome is Outcome.UNPROCESSED


def test_two_submissions_from_one_student_both_need_review_and_show_the_difference(tmp_path):
    info = "高二（3）班  张三  20240301"
    make_docx(tmp_path / "张三.docx", ["题目", info, "我统计了 18 天，其中有 9 天有人忘了做值日。"])
    make_docx(tmp_path / "张三(1).docx", ["题目", info, "我统计了 18 天，其中有 10 天有人忘了做值日。"])

    records = by_name(take_inventory(tmp_path, ROSTER))

    for name, other, mine, theirs in [
        ("张三.docx", "张三(1).docx", "9 天", "10 天"),
        ("张三(1).docx", "张三.docx", "10 天", "9 天"),
    ]:
        record = records[name]
        assert record.outcome is Outcome.REVIEW
        note = "\n".join(record.notes)
        assert other in note
        assert mine in note and theirs in note


def test_difference_excerpt_does_not_cut_a_number_in_half(tmp_path):
    info = "高二（3）班  张三  20240301"
    make_docx(tmp_path / "a.docx", ["题目", info, "这个学期到现在，我统计了 18 天，其中有 9 天有人忘了。"])
    make_docx(tmp_path / "b.docx", ["题目", info, "这个学期到现在，我统计了 18 天，其中有 10 天有人忘了。"])

    note = "\n".join(by_name(take_inventory(tmp_path, ROSTER))["a.docx"].notes)

    assert "18 天" in note


def test_identical_duplicates_are_said_to_be_identical(tmp_path):
    paragraphs = ["题目", "高二（3）班  张三  20240301", *BODY]
    make_docx(tmp_path / "张三.docx", paragraphs)
    make_docx(tmp_path / "张三 - 副本.docx", paragraphs)

    records = take_inventory(tmp_path, ROSTER).records

    assert all(r.outcome is Outcome.REVIEW for r in records)
    assert all(any("完全相同" in n for n in r.notes) for r in records)


def test_identity_problems_send_a_file_to_review(tmp_path):
    make_docx(tmp_path / "x.docx", ["题目", "高二（3）班  张三  20240302", *BODY])

    (record,) = take_inventory(tmp_path, ROSTER).records

    assert record.outcome is Outcome.REVIEW
    assert any("对不上" in n for n in record.notes)


def test_doc_that_cannot_be_converted_is_reported_not_dropped(tmp_path):
    (tmp_path / "张三.doc").write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 64)

    (record,) = take_inventory(tmp_path, ROSTER, convert=no_conversion).records

    assert record.outcome is Outcome.UNPROCESSED
    assert record.student_id == "20240301"
    assert any("测试里不转换" in n for n in record.notes)


def test_doc_is_read_through_the_converter(tmp_path):
    (tmp_path / "x.doc").write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 64)
    converted = make_docx(tmp_path.parent / "converted.docx", ["题目", "李四 20240302", *BODY])

    def fake_convert(paths):
        return ConversionResult(converted={p: converted.read_bytes() for p in paths}, failed={})

    (record,) = take_inventory(tmp_path, ROSTER, convert=fake_convert).records

    assert record.outcome is Outcome.READY
    assert (record.name, record.student_id) == ("李四", "20240302")


def test_students_with_no_file_are_listed(tmp_path):
    make_docx(tmp_path / "a.docx", ["题目", "张三 20240301", *BODY])
    (tmp_path / "李四.pdf").write_bytes(b"%PDF-1.7")

    inventory = take_inventory(tmp_path, ROSTER)

    assert inventory.missing_students == [("20240303", "王五")]


def test_without_a_roster_nobody_is_listed_as_missing(tmp_path):
    make_docx(tmp_path / "a.docx", ["题目", "张三 20240301", *BODY])

    inventory = take_inventory(tmp_path, None)

    assert inventory.missing_students == []
    assert inventory.records[0].outcome is Outcome.READY


def test_subfolders_are_mentioned(tmp_path):
    (tmp_path / "一班").mkdir()

    assert take_inventory(tmp_path, None).skipped_folders == ["一班"]
