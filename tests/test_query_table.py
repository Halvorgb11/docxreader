"""Tester for query_table (filtrer, sorter, tell, summer) på Excel og CSV."""

from pathlib import Path

import pytest

from docxreader.tables import parse_number
from docxreader.tools import query_table

SAMPLES = Path(__file__).parent.parent / "samples"
XLSX = str(SAMPLES / "budsjett.xlsx")
CSV = str(SAMPLES / "reiseregning.csv")


def q(path, **kwargs) -> str:
    return query_table.invoke({"path": path, **kwargs})


@pytest.mark.parametrize("text, expected", [
    ("1250", 1250), ("-3", -3), ("1250.5", 1250.5), ("1 250,50", 1250.5),
    ("8 430,00", 8430), ("2026-09-01", None), ("IT", None), ("", None),
])
def test_parse_number_handles_norwegian_numbers(text, expected):
    assert parse_number(text) == expected


def test_filter_and_sort_finds_largest_it_payment():
    result = q(XLSX, sheet="Transaksjoner", where=["Avdeling = IT"], sort_by="Beløp", descending=True, limit=1)
    assert result.startswith("'Transaksjoner': 89 av 250 rader der Avdeling = IT, sortert etter Beløp (kr) (synkende).")
    assert "| 180 | 2026-06-12 | Fjellsikring AS | 1250000 | IT |" in result
    assert "Viser 1 av 89 rader" in result


def test_group_by_sum_sorted_largest_first():
    result = q(XLSX, sheet="Transaksjoner", group_by="Avdeling", aggregate="sum", value_column="Beløp")
    lines = result.splitlines()
    assert "| Avdeling | Antall rader | Sum av Beløp (kr) |" in lines
    assert lines.index("| IT | 89 | 2109765 |") < lines.index("| Salg | 85 | 779367 |")


def test_count_with_date_range_and_contains():
    result = q(XLSX, sheet="Transaksjoner", where=["Dato >= 2026-06-01", "Dato < 2026-07-01", "Leverandør ~ AS"], aggregate="count")
    assert "| Alle | 7 |" in result


def test_sheet_name_is_forgiving():
    for sheet in ["Transaksjoner", "transaksjoner", "Ark: Transaksjoner", "trans"]:
        assert q(XLSX, sheet=sheet, aggregate="count").startswith("'Transaksjoner'")


def test_choose_columns():
    result = q(XLSX, sheet="Oversikt", columns=["Avdeling", "Forbruk"])
    assert "| Rad | Avdeling | Forbruk (kr) |" not in result
    assert "| Rad | Avdeling | Forbruk |" in result
    assert "Budsjett" not in result


def test_csv_norwegian_numbers_are_summed_and_empty_cells_skipped():
    result = q(CSV, group_by="Ansatt", aggregate="sum", value_column="Beløp")
    assert "| Per Æsøy | 2 | 10530 |" in result
    assert "| Kari Nordmann | 3 | 4120.50 |" in result
    assert "| Lise Øvrebø | 1 | – |" in result
    assert "ikke tall" not in result  # tom celle er ikke en feil


def test_csv_numeric_sort_not_text_sort():
    # Som tekst ville "890,00" kommet før "8 430,00".
    result = q(CSV, sort_by="Beløp", descending=True, limit=1)
    assert "| 5 | 2026-09-10 | Per Æsøy | Tromsø | 8 430,00 | Fly |" in result


def test_avg_min_max():
    assert "| Alle | 3 | 1373.50 |" in q(CSV, where=["Ansatt ~ kari"], aggregate="avg", value_column="Beløp")
    assert "| Alle | 7 | 450 |" in q(CSV, aggregate="min", value_column="Beløp")


def test_limit_is_capped():
    result = q(XLSX, sheet="Transaksjoner", limit=10_000)
    assert "Viser 100 av 250 rader" in result


def test_no_matching_rows():
    assert "(ingen rader)" in q(CSV, where=["Reisemål = Bodø"])


@pytest.mark.parametrize("kwargs, message", [
    ({"path": XLSX}, "flere ark"),
    ({"path": XLSX, "sheet": "Finnes ikke"}, "Ark med innhold: Oversikt, Transaksjoner, Hjelpetall"),
    ({"path": CSV, "where": ["Beløpet > 5"]}, "Kolonner: Dato, Ansatt"),
    ({"path": CSV, "where": ["bare tekst"]}, "forstår ikke betingelsen"),
    ({"path": CSV, "aggregate": "sum"}, "trenger value_column"),
    ({"path": str(SAMPLES / "prosjektplan.docx")}, "Støttede filtyper: .xlsx, .csv"),
])
def test_errors_explain_what_to_do(kwargs, message):
    result = query_table.invoke(kwargs)
    assert result.startswith("Feil:")
    assert message in result
