import hashlib

from conftest import make_docx
from homework_formatter.scanner import FileKind, fingerprint, scan_folder


def test_recognises_each_kind_of_file(tmp_path):
    make_docx(tmp_path / "张三.docx", ["题目"])
    (tmp_path / "李四.doc").write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 100)
    (tmp_path / "王五.pdf").write_bytes(b"%PDF-1.7\n...")
    (tmp_path / "~$张三.docx").write_bytes(b"\x05ling" + b"\0" * 150)
    (tmp_path / "照片.jpg").write_bytes(b"\xff\xd8\xff\xe0")

    kinds = {f.path.name: f.kind for f in scan_folder(tmp_path)}

    assert kinds == {
        "张三.docx": FileKind.DOCX,
        "李四.doc": FileKind.DOC,
        "王五.pdf": FileKind.PDF,
        "~$张三.docx": FileKind.TEMP,
        "照片.jpg": FileKind.OTHER,
    }


def test_docx_extension_with_broken_content_is_not_trusted(tmp_path):
    (tmp_path / "坏文件.docx").write_bytes(b"this is not a word file")

    (entry,) = scan_folder(tmp_path)

    assert entry.kind is FileKind.OTHER


def test_records_size_and_fingerprint(tmp_path):
    data = b"%PDF-1.7 hello"
    (tmp_path / "a.pdf").write_bytes(data)

    (entry,) = scan_folder(tmp_path)

    assert entry.size == len(data)
    assert entry.sha256 == hashlib.sha256(data).hexdigest()


def test_subfolders_are_not_entered(tmp_path):
    (tmp_path / "子文件夹").mkdir()
    make_docx(tmp_path / "子文件夹" / "里面.docx", ["x"])
    make_docx(tmp_path / "外面.docx", ["x"])

    assert [f.path.name for f in scan_folder(tmp_path)] == ["外面.docx"]


def test_fingerprint_changes_when_a_file_changes(tmp_path):
    target = tmp_path / "a.pdf"
    target.write_bytes(b"%PDF one")
    before = fingerprint(tmp_path)

    assert fingerprint(tmp_path) == before

    target.write_bytes(b"%PDF two")
    assert fingerprint(tmp_path) != before
