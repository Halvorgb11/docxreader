"""Tester for verktøyene. Kjører uten Claude – bare ren Python."""

from pathlib import Path

from docxreader.tools import read_pdf

SAMPLE_PDF = Path(__file__).parent.parent / "samples" / "rapport.pdf"


def test_tool_metadata():
    # Dette er det modellen ser: navn, beskrivelse og argumentskjema.
    assert read_pdf.name == "read_pdf"
    assert "PDF" in read_pdf.description
    assert "path" in read_pdf.args
    assert "description" in read_pdf.args["path"]


def test_reads_all_pages_with_markers():
    # Verktøy kalles med .invoke() og et dict med argumentene.
    result = read_pdf.invoke({"path": str(SAMPLE_PDF)})

    assert "--- Side 1 av 2 ---" in result
    assert "--- Side 2 av 2 ---" in result
    assert "4,2 millioner kroner" in result
    assert "Kari Nordmann" in result


def test_missing_file_returns_error_text():
    result = read_pdf.invoke({"path": "finnes_ikke.pdf"})
    assert result.startswith("Feil:")


def test_non_pdf_returns_error_text(tmp_path):
    txt = tmp_path / "notat.txt"
    txt.write_text("hei")
    result = read_pdf.invoke({"path": str(txt)})
    assert result.startswith("Feil:")
    assert "ikke en PDF" in result


def test_corrupt_pdf_returns_error_text(tmp_path):
    bad = tmp_path / "ødelagt.pdf"
    bad.write_bytes(b"dette er ikke en ekte pdf")
    result = read_pdf.invoke({"path": str(bad)})
    assert result.startswith("Feil:")
