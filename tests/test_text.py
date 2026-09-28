"""Tester for Markdown (.md) og ren tekst (.txt)."""

from pathlib import Path

from docxreader.tools import document_outline, read_document, read_section, search_document

SAMPLE_MD = Path(__file__).parent.parent / "samples" / "retningslinjer.md"


def _write(tmp_path, name, content, encoding="utf-8") -> str:
    f = tmp_path / name
    f.write_bytes(content.encode(encoding))
    return str(f)


def test_markdown_outline():
    result = document_outline.invoke({"path": str(SAMPLE_MD)})
    assert "## Satser  (~19 ord, 1 tabell)" in result
    assert "### Bredbånd" in result


def test_markdown_hash_in_code_block_is_not_a_heading():
    result = document_outline.invoke({"path": str(SAMPLE_MD)})
    assert "Koble til VPN" not in result
    assert "vpn connect" in read_section.invoke({"path": str(SAMPLE_MD), "heading": "Sikkerhet"})


def test_markdown_section_includes_subsection():
    result = read_section.invoke({"path": str(SAMPLE_MD), "heading": "Utstyr"})
    assert "- Hodetelefoner med mikrofon" in result
    assert "### Bredbånd" in result
    assert "Satser" not in result


def test_markdown_search_table_row_and_sentence():
    result = search_document.invoke({"path": str(SAMPLE_MD), "query": "bredbånd"})
    assert "| Ordning | Beløp | Kommentar |" in result
    assert "| Bredbånd | 400 kr/mnd | Mot kvittering |" in result
    assert "- Kontaktperson" not in result
    assert "Lars Bakken" in search_document.invoke({"path": str(SAMPLE_MD), "query": "kontaktperson"})


def test_markdown_paragraph_lines_are_joined(tmp_path):
    # Et avsnitt brutt over flere linjer i filen er fortsatt ett avsnitt.
    f = _write(tmp_path, "a.md", "Første linje\nandre linje.\n\nNytt avsnitt.")
    assert read_document.invoke({"path": f}) == "Første linje andre linje.\n\nNytt avsnitt."


def test_markdown_heading_directly_followed_by_text(tmp_path):
    f = _write(tmp_path, "a.md", "## Mål\nØke salget.\n## Plan\nJobbe mer.")
    assert read_section.invoke({"path": f, "heading": "Mål"}) == "## Mål\n\nØke salget."


def test_markdown_several_top_level_headings_are_in_paths(tmp_path):
    # Flere "#"-kapitler: ingen av dem er "tittelen", så de er med i stien.
    f = _write(tmp_path, "a.md", "# Nord\n## Status\nBra i nord.\n# Sør\n## Status\nBra i sør.")
    result = read_section.invoke({"path": f, "heading": "Sør > Status"})
    assert "Bra i sør." in result


def test_plain_text_has_no_headings(tmp_path):
    f = _write(tmp_path, "notat.txt", "# ikke en overskrift\n\nEt avsnitt\nover to linjer.")
    assert read_document.invoke({"path": f}) == "# ikke en overskrift\n\nEt avsnitt over to linjer."
    assert "ingen overskrifter" in document_outline.invoke({"path": f})


def test_plain_text_in_windows_encoding(tmp_path):
    # Eldre norske filer fra Windows er ofte cp1252, ikke UTF-8.
    f = _write(tmp_path, "gammel.txt", "Blåbærsyltetøy på skjærgården.", encoding="cp1252")
    assert read_document.invoke({"path": f}) == "Blåbærsyltetøy på skjærgården."


def test_binary_file_with_text_suffix_gives_error(tmp_path):
    f = tmp_path / "bilde.txt"
    f.write_bytes(b"\x89PNG\x00\x00\x00")
    assert read_document.invoke({"path": str(f)}).startswith("Feil:")


def test_empty_text_file(tmp_path):
    assert "ingen tekst" in read_document.invoke({"path": _write(tmp_path, "tom.md", "")})
