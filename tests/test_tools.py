"""Tester for verktøyene. Kjører uten Claude – bare ren Python."""

from pathlib import Path

from docx import Document

from docxreader.tools import read_docx

SAMPLE_DOCX = Path(__file__).parent.parent / "samples" / "prosjektplan.docx"


def test_docx_tool_metadata():
    assert read_docx.name == "read_docx"
    assert ".docx" in read_docx.description
    assert "description" in read_docx.args["path"]


def test_docx_headings_become_markdown():
    result = read_docx.invoke({"path": str(SAMPLE_DOCX)})
    assert "# Prosjektplan: Ny nettbutikk" in result
    assert "\n## Budsjett\n" in result
    assert "\n### Fase 1: Design\n" in result


def test_docx_lists():
    result = read_docx.invoke({"path": str(SAMPLE_DOCX)})
    assert "- Lansere ny nettbutikk i november\n- Halvere tiden" in result
    assert "1. Sett opp nettbutikkplattform\n2. Integrer betaling\n3. Test med ekte kunder" in result


def test_docx_table_is_markdown_and_in_order():
    result = read_docx.invoke({"path": str(SAMPLE_DOCX)})
    assert "| Post | Beløp (kr) | Ansvarlig |" in result
    assert "| Markedsføring | 150 000 | Per Æsøy |" in result
    # Tabellen skal stå MELLOM avsnittet før og avsnittet etter den.
    before = result.index("Tabellen under viser")
    table = result.index("| Post |")
    after = result.index("Totalt budsjett")
    assert before < table < after


def test_docx_header_and_footer():
    result = read_docx.invoke({"path": str(SAMPLE_DOCX)})
    assert "Topptekst: Fjellbekk AS – Internt" in result
    assert "Bunntekst: Konfidensielt" in result


def test_docx_numbering_restarts_after_other_content(tmp_path):
    doc = Document()
    doc.add_paragraph("A", style="List Number")
    doc.add_paragraph("B", style="List Number")
    doc.add_paragraph("Mellomtekst")
    doc.add_paragraph("C", style="List Number")
    f = tmp_path / "liste.docx"
    doc.save(f)

    result = read_docx.invoke({"path": str(f)})
    assert "1. A\n2. B\n\nMellomtekst\n\n1. C" in result


def test_docx_table_cell_with_pipe_and_newline(tmp_path):
    doc = Document()
    table = doc.add_table(rows=2, cols=1)
    table.rows[0].cells[0].text = "Kolonne"
    table.rows[1].cells[0].text = "a | b\nc"
    f = tmp_path / "tabell.docx"
    doc.save(f)

    result = read_docx.invoke({"path": str(f)})
    assert "| a \\| b c |" in result


def test_docx_empty_document(tmp_path):
    f = tmp_path / "tom.docx"
    Document().save(f)
    assert "ingen tekst" in read_docx.invoke({"path": str(f)})


def test_docx_missing_file_returns_error_text():
    assert read_docx.invoke({"path": "finnes_ikke.docx"}).startswith("Feil:")


def test_docx_wrong_extension_returns_error_text(tmp_path):
    txt = tmp_path / "notat.txt"
    txt.write_text("hei")
    result = read_docx.invoke({"path": str(txt)})
    assert result.startswith("Feil:")
    assert "ikke en .docx" in result


def test_docx_corrupt_file_returns_error_text(tmp_path):
    bad = tmp_path / "ødelagt.docx"
    bad.write_bytes(b"dette er ikke en ekte docx")
    result = read_docx.invoke({"path": str(bad)})
    assert result.startswith("Feil:")
    assert "ikke et gyldig Word-dokument" in result
