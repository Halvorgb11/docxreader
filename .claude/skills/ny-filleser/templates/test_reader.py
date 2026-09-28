"""Tester for <FORMAT> (.<endelse>). Testfilen lages av scripts/make_sample_<format>.py."""

from pathlib import Path

from docxreader.tools import document_outline, read_document, read_section, search_document

SAMPLE = Path(__file__).parent.parent / "samples" / "<fil>.<endelse>"


def _read() -> str:
    return read_document.invoke({"path": str(SAMPLE)})


def test_sections_become_headings():
    assert "## <Seksjon>" in _read()


def test_table():
    assert "| <a> | <b> |" in _read()


def test_reading_order():
    result = _read()
    assert result.index("<først>") < result.index("<sist>")


def test_planted_fact_is_searchable():
    result = search_document.invoke({"path": str(SAMPLE), "query": "<søkeord>"})
    assert "[<Seksjon>]" in result
    assert "<faktum>" in result


def test_read_one_section():
    result = read_section.invoke({"path": str(SAMPLE), "heading": "<Seksjon>"})
    assert result.startswith("## <Seksjon>")
    assert "<tekst fra neste seksjon>" not in result


def test_outline():
    assert "## <Seksjon>  (~" in document_outline.invoke({"path": str(SAMPLE)})


def test_corrupt_file_returns_error_text(tmp_path):
    bad = tmp_path / "ødelagt.<endelse>"
    bad.write_bytes(b"ikke en ekte fil")
    assert read_document.invoke({"path": str(bad)}).startswith("Feil:")
