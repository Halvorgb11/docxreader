"""Tester for CSV (.csv). Testfilen lages av scripts/make_sample_csv.py."""

from pathlib import Path

from docxreader.tools import document_outline, read_document, read_section, search_document

SAMPLE = Path(__file__).parent.parent / "samples" / "reiseregning.csv"


def _write(tmp_path, content: str, encoding="utf-8") -> str:
    f = tmp_path / "data.csv"
    f.write_bytes(content.encode(encoding))
    return str(f)


def test_csv_in_windows_encoding_with_semicolon():
    result = read_document.invoke({"path": str(SAMPLE)})
    assert result.startswith("CSV-fil med 7 rader og 5 kolonner (skilletegn: semikolon).")
    assert "Kolonner: Dato, Ansatt, Reisemål, Beløp, Kategori" in result
    assert "| 5 | 2026-09-10 | Per Æsøy | Tromsø | 8 430,00 | Fly |" in result


def test_csv_quoted_cell_with_delimiter_inside():
    assert "| Hansen; Ola (konsulent) |" in read_document.invoke({"path": str(SAMPLE)})


def test_csv_row_numbers_are_line_numbers_and_empty_lines_are_skipped():
    result = read_document.invoke({"path": str(SAMPLE)})
    assert "| 6 |" not in result  # linje 6 er tom
    assert "| 7 | 2026-09-11 |" in result


def test_csv_search_and_section():
    result = search_document.invoke({"path": str(SAMPLE), "query": "tromsø fly"})
    assert "[Tabell]" in result and "| 5 | 2026-09-10 | Per Æsøy |" in result
    assert read_section.invoke({"path": str(SAMPLE), "heading": "Tabell"}).startswith("## Tabell")
    assert "## Tabell  (~" in document_outline.invoke({"path": str(SAMPLE)})


def test_csv_comma_and_tab_delimiters(tmp_path):
    assert "| 2 | Kari | 41 |" in read_document.invoke({"path": _write(tmp_path, "Navn,Alder\nKari,41\n")})
    assert "| 2 | Kari | 41 |" in read_document.invoke({"path": _write(tmp_path, "Navn\tAlder\nKari\t41\n")})


def test_csv_ragged_rows_and_blank_header(tmp_path):
    result = read_document.invoke({"path": _write(tmp_path, "Navn;\nKari;41;ekstra\nOla\n")})
    assert "| Rad | Navn | Kolonne 2 | Kolonne 3 |" in result
    assert "| 3 | Ola |  |  |" in result


def test_empty_csv(tmp_path):
    assert "ingen tekst" in read_document.invoke({"path": _write(tmp_path, "")})


def test_csv_title_line_above_header(tmp_path):
    f = _write(tmp_path, "Reiser september\nNavn;Beløp\nKari;1 250,50\n")
    result = read_document.invoke({"path": f})
    assert "## Tabell\n\nReiser september\n\n| Rad | Navn | Beløp |" in result
    assert "| 3 | Kari | 1 250,50 |" in result
