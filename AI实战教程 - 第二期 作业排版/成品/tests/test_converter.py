import pytest

from conftest import REAL_INPUT, needs_real_samples
from homework_formatter import converter
from homework_formatter.converter import convert_doc_files, find_soffice
from homework_formatter.reader import read_docx
from homework_formatter.scanner import fingerprint

needs_soffice = pytest.mark.skipif(find_soffice() is None, reason="没有 LibreOffice")


def test_nothing_to_convert_does_not_start_libreoffice(monkeypatch):
    monkeypatch.setattr(converter, "find_soffice", lambda: pytest.fail("不应该去找 LibreOffice"))

    result = convert_doc_files([])

    assert result.converted == {} and result.failed == {}


def test_without_libreoffice_every_file_fails_with_a_plain_explanation(tmp_path, monkeypatch):
    monkeypatch.setattr(converter, "find_soffice", lambda: None)
    path = tmp_path / "a.doc"
    path.write_bytes(b"x")

    result = convert_doc_files([path])

    assert result.converted == {}
    assert "LibreOffice" in result.failed[path]


@needs_soffice
@needs_real_samples
def test_real_doc_files_convert_with_their_text_intact_and_originals_untouched():
    paths = sorted(REAL_INPUT.glob("*.doc"))
    before = fingerprint(REAL_INPUT)

    result = convert_doc_files(paths)

    assert result.failed == {}
    texts = {p.name: "\n".join(read_docx(data).paragraphs) for p, data in result.converted.items()}
    assert "唐小满 学号:20240307" in texts["唐小满改.doc"]
    assert "高二（3）班  邱实  20240327" in texts["邱实.doc"]
    assert fingerprint(REAL_INPUT) == before
    # LibreOffice 不应在原件旁边留下锁文件之类的东西
    assert len(before) == 37  # 38 个文件里有 1 个是 Word 的临时文件，不算原件


@needs_soffice
@needs_real_samples
def test_no_sentence_of_a_real_doc_file_is_lost_in_conversion():
    """独立核对转换这一步：直接从旧版 .doc 的原始字节里找出每一句话（以句号结尾的），
    它们都必须出现在 LibreOffice 转出来的文字里。不经过 LibreOffice 读原文，所以它转丢了这里能发现。"""
    import re

    paths = sorted(REAL_INPUT.glob("*.doc"))
    result = convert_doc_files(paths)

    for path in paths:
        raw = path.read_bytes()
        decoded = raw.decode("utf-16-le", errors="ignore") + raw[1:].decode("utf-16-le", errors="ignore")
        sentences = set(re.findall(r"[一-鿿][一-鿿0-9A-Za-z ，、：；（）]{6,}。", decoded))
        converted = "".join("".join(read_docx(result.converted[path]).paragraphs).split())
        assert len(sentences) >= 6, path.name  # 确实从原始字节里读到了正文
        missing = [s for s in sentences if "".join(s.split()) not in converted]
        assert missing == [], path.name
