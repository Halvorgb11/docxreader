"""Tester for menyen (readers.READERS) og det som gjelder alle filtyper."""

import pytest

from docxreader.readers import READERS
from docxreader.tools import ALL_TOOLS, read_document


def test_tools_have_names_and_argument_descriptions():
    assert [t.name for t in ALL_TOOLS] == [
        "read_document", "document_outline", "read_section", "search_document",
    ]
    # parse_docstring=True legger Args-beskrivelsen inn i skjemaet modellen ser.
    for t in ALL_TOOLS:
        assert "description" in t.args["path"]


@pytest.mark.parametrize("suffix", READERS)
def test_every_supported_file_type_is_mentioned_in_every_tool(suffix):
    # Modellen vet bare det som står i beskrivelsene. Legger du til en filtype
    # i READERS uten å nevne den i docstringene, feiler denne testen.
    for t in ALL_TOOLS:
        assert suffix in t.description, f"{t.name} nevner ikke {suffix}"


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
