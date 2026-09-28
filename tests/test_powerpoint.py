"""Tester for PowerPoint (.pptx). Testfilen lages av scripts/make_sample_pptx.py."""

from pathlib import Path

from docxreader.tools import document_outline, read_document, read_section, search_document

SAMPLE_PPTX = Path(__file__).parent.parent / "samples" / "salgsmote.pptx"


def _read(**kwargs) -> str:
    return read_document.invoke({"path": str(SAMPLE_PPTX), **kwargs})


def test_pptx_slides_become_sections_with_titles():
    result = _read()
    assert result.startswith("Tittel: Salgsmøte Q3\nPresentasjon med 5 lysbilder.")
    assert "## Lysbilde 2: Resultater" in result
    assert "## Lysbilde 5\n" in result  # lysbilde uten tittel


def test_pptx_title_is_not_repeated_in_body():
    assert _read().count("Resultater") == 1


def test_pptx_bullets_keep_indent_levels():
    assert (
        "- Omsetningen økte med 12 prosent\n"
        "  - Størst vekst i Bergen\n"
        "  - Svakere i Stavanger\n"
        "- Tre nye storkunder"
    ) in _read()


def test_pptx_table():
    assert "| Kari Nordmann | 8,1 | Ja |" in _read()


def test_pptx_text_boxes_in_reading_order_including_groups():
    result = _read()
    # Boksene ble laget i rekkefølgen "Til slutt", "Først", (gruppe) "Deretter".
    assert result.index("Først:") < result.index("Deretter:") < result.index("Til slutt:")


def test_pptx_speaker_notes_are_included_and_searchable():
    assert "Talenotater: Nevn at den største nye kunden er Havbruk Vest AS" in _read()
    result = search_document.invoke({"path": str(SAMPLE_PPTX), "query": "havbruk"})
    assert "[Lysbilde 2: Resultater]" in result
    assert "4,2 MNOK" in result


def test_pptx_outline_lists_slides():
    result = document_outline.invoke({"path": str(SAMPLE_PPTX)})
    assert "## Lysbilde 3: Salg per selger  (~" in result
    assert "1 tabell" in result
    assert "Kari Nordmann" not in result


def test_pptx_read_one_slide_by_title_or_number():
    by_title = read_section.invoke({"path": str(SAMPLE_PPTX), "heading": "Neste steg"})
    by_number = read_section.invoke({"path": str(SAMPLE_PPTX), "heading": "Lysbilde 4"})
    assert by_title == by_number
    assert by_title.startswith("## Lysbilde 4: Neste steg")
    assert "Spørsmål?" not in by_title


def test_pptx_corrupt_file_returns_error_text(tmp_path):
    bad = tmp_path / "ødelagt.pptx"
    bad.write_bytes(b"ikke en presentasjon")
    result = read_document.invoke({"path": str(bad)})
    assert result.startswith("Feil:")
    assert "ikke en gyldig PowerPoint-fil" in result
