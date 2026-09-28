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
    ("8\u00a0430,00", 8430), ("2026-09-01", None), ("IT", None), ("", None),
    # engelsk og tysk/norsk med tusenskille, flere like skilletegn
    ("1,250,000.50", 1250000.5), ("1.250.000,50", 1250000.5), ("1,250.5", 1250.5),
    ("1.250.000", 1250000), ("1,250,000", 1250000),
    # valuta, prosent, norsk ",-", typografisk minus
    ("kr 1 250", 1250), ("1 250 kr", 1250), ("NOK 99,90", 99.9), ("$1,250.00", 1250),
    ("25 %", 25), ("1 250,-", 1250), ("\u2212500", -500),
    # ett enkelt skilletegn = desimaltegn; datoer og ugyldig gruppering er ikke tall
    ("1,250", 1.25), ("12.05.2026", None), ("1,2,3", None), ("1.2,3", None),
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


def test_or_alternatives_within_one_condition():
    result = q(CSV, where=["Reisemål = Bergen OR Reisemål = Oslo"], aggregate="count")
    assert "| Alle | 3 |" in result


def test_or_combined_with_and():
    # (Kategori = Fly ELLER Kategori = Tog) OG Ansatt ~ kari
    result = q(CSV, where=["Kategori = Fly OR Kategori = Tog", "Ansatt ~ kari"], aggregate="count")
    assert "| Alle | 2 |" in result


def test_sum_of_numbers_stored_as_text_in_different_formats():
    result = q(XLSX, sheet="Faktisk", aggregate="sum", value_column="faktisk")
    assert result.startswith("'Kvartal – Faktisk Q1'")
    assert "| Alle | 3 | 2525000.50 |" in result


def test_table_in_sheet_with_several_tables():
    assert "| 8 | HR | 100000 |" in q(XLSX, sheet="Kvartal – Budsjett Q1")
    result = q(XLSX, sheet="Kvartal")
    assert result.startswith("Feil:")
    assert "passer flere tabeller: Kvartal – Budsjett Q1, Kvartal – Faktisk Q1" in result


def test_limit_is_capped():
    result = q(XLSX, sheet="Transaksjoner", limit=10_000)
    assert "Viser 100 av 250 rader" in result


def test_no_matching_rows():
    assert "(ingen rader)" in q(CSV, where=["Reisemål = Bodø"])


@pytest.mark.parametrize("kwargs, message", [
    ({"path": XLSX}, "flere ark"),
    ({"path": XLSX, "sheet": "Finnes ikke"}, "Ark med innhold: Oversikt, Transaksjoner, Hjelpetall, Kvartal – Budsjett Q1"),
    ({"path": CSV, "where": ["Beløpet > 5"]}, "Kolonner: Dato, Ansatt"),
    ({"path": CSV, "where": ["bare tekst"]}, "forstår ikke betingelsen"),
    ({"path": CSV, "aggregate": "sum"}, "trenger value_column"),
    ({"path": str(SAMPLES / "prosjektplan.docx")}, "Støttede filtyper: .xlsx, .csv"),
])
def test_errors_explain_what_to_do(kwargs, message):
    result = query_table.invoke(kwargs)
    assert result.startswith("Feil:")
    assert message in result


# --- calculate: beregnede kolonner --------------------------------------------


def test_calculate_new_column_from_two_columns():
    result = q(XLSX, sheet="Oversikt", calculate=["Differanse = Budsjett - Forbruk"], columns=["Avdeling", "Differanse"])
    assert "| 2 | Salg | 250000 |" in result
    assert "| 3 | IT | -350000.50 |" in result


def test_calculate_respects_operator_precedence_and_chains():
    result = q(CSV, calculate=["X = Beløp + 10 * 2", "Y = X / 2"], columns=["Y"], limit=1)
    assert "| 2 | 635.25 |" in result  # (1250.5 + 20) / 2


def test_calculated_column_can_be_filtered_sorted_and_summed():
    result = q(XLSX, sheet="Oversikt", calculate=["Over = Forbruk - Budsjett"], where=["Over > 0"], aggregate="sum", value_column="Over")
    assert "| Alle | 1 | 350000.50 |" in result


def test_calculate_empty_or_invalid_values_give_empty_cell():
    result = q(CSV, calculate=["Null = Beløp / 0", "Dobbel = Beløp * 2"], where=["Ansatt ~ Lise"], columns=["Null", "Dobbel"])
    assert "| 9 |  |  |" in result  # Lise har tomt beløp


@pytest.mark.parametrize("expression, message", [
    ("bare tekst", "forstår ikke beregningen"),
    ("X = Beløp * (2 + 1)", "parenteser støttes ikke"),
    ("X = Finnesikke * 2", "fant ingen entydig kolonne"),
])
def test_calculate_errors(expression, message):
    result = q(CSV, calculate=[expression])
    assert result.startswith("Feil:") and message in result


# --- attachment: tabeller i e-postvedlegg -------------------------------------

EML = str(SAMPLES / "tilbud.eml")


def test_query_attachment_with_calculation():
    result = q(EML, attachment="tilbud.xlsx", calculate=["Sum = Antall * Pris per stk"], aggregate="sum", value_column="Sum")
    assert result.startswith("'Tilbud'")
    assert "| Alle | 2 | 192000 |" in result


def test_attachment_name_prefers_table_files_when_ambiguous():
    # "tilbud" passer både tilbud.xlsx og "Tilbud 2026-117.eml".
    assert "| 2 | Nordic Data AS | Server R750 | 4 | 45000 |" in q(EML, attachment="tilbud")


@pytest.mark.parametrize("kwargs, message", [
    ({"path": EML, "attachment": "avtaleutkast.docx"}, "er ikke et regneark eller en CSV-fil"),
    ({"path": EML, "attachment": "finnes.xlsx"}, "fant ikke vedlegget 'finnes.xlsx'. Vedlegg: tilbud.xlsx"),
    ({"path": EML, "attachment": "Tilbud 2026-117.eml > x.csv"}, "fant ikke vedlegget 'x.csv'. Vedlegg: (ingen)"),
    ({"path": CSV, "attachment": "x.xlsx"}, "attachment kan bare brukes når path er en e-post"),
])
def test_attachment_errors(kwargs, message):
    result = query_table.invoke(kwargs)
    assert result.startswith("Feil:") and message in result
