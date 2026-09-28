"""Tester for menyen (readers.READERS) og det som gjelder alle filtyper."""

import pytest

from docxreader.readers import READERS, TABLE_READERS
from docxreader.tools import ALL_TOOLS, query_table, read_document

DOCUMENT_TOOLS = [t for t in ALL_TOOLS if t is not query_table]


def test_tools_have_names_and_argument_descriptions():
    assert [t.name for t in ALL_TOOLS] == [
        "read_document", "document_outline", "read_section", "search_document", "query_table",
    ]
    # parse_docstring=True legger Args-beskrivelsen inn i skjemaet modellen ser.
    for t in ALL_TOOLS:
        assert "description" in t.args["path"]


@pytest.mark.parametrize("suffix", READERS)
def test_every_supported_file_type_is_mentioned_in_every_tool(suffix):
    # Modellen vet bare det som står i beskrivelsene. Legger du til en filtype
    # i READERS uten å nevne den i docstringene, feiler denne testen.
    for t in DOCUMENT_TOOLS:
        assert suffix in t.description, f"{t.name} nevner ikke {suffix}"


@pytest.mark.parametrize("suffix", TABLE_READERS)
def test_every_table_file_type_is_mentioned_in_query_table(suffix):
    assert suffix in query_table.description
    assert suffix in READERS  # tabellfiler skal også kunne leses som dokumenter


def test_unsupported_file_type_lists_supported_types(tmp_path):
    f = tmp_path / "data.xyz"
    f.write_text("hei")
    result = read_document.invoke({"path": str(f)})
    assert result.startswith("Feil:")
    assert "støttes ikke" in result
    assert ".docx" in result and ".pptx" in result


def test_suffix_is_case_insensitive(tmp_path):
    f = tmp_path / "NOTAT.MD"
    f.write_text("# Hei\n\nTekst.")
    assert "# Hei" in read_document.invoke({"path": str(f)})


# --- Felles logikk i blocks.py ------------------------------------------------

from docxreader.blocks import split_sentences  # noqa: E402


@pytest.mark.parametrize("text, expected", [
    ("Første. Andre!", ["Første.", "Andre!"]),
    ("Møtet er kl. 10 i rom Fjorden. Ta med PC.", ["Møtet er kl. 10 i rom Fjorden.", "Ta med PC."]),
    ("Fristen er 8. oktober. Ikke glem.", ["Fristen er 8. oktober.", "Ikke glem."]),
    ("Gjelder fra 1. januar 2027. De erstatter reglene.", ["Gjelder fra 1. januar 2027.", "De erstatter reglene."]),
    ("Vi har f.eks. Oslo og Bergen. Ferdig.", ["Vi har f.eks. Oslo og Bergen.", "Ferdig."]),
])
def test_split_sentences_handles_norwegian_abbreviations_and_dates(text, expected):
    assert split_sentences(text) == expected


def test_top_level_heading_wins_when_same_name_is_nested(tmp_path):
    f = tmp_path / "a.md"
    f.write_text("## Innledning\nØverst.\n## Kapittel\n### Innledning\nNede.")
    from docxreader.tools import read_section

    assert read_section.invoke({"path": str(f), "heading": "Innledning"}) == "## Innledning\n\nØverst."
