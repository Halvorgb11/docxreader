"""Leser for Excel-arbeidsbøker (.xlsx), med openpyxl.

To måter å lese på:
- load_tables(): hvert ark som en Table (rader og kolonner) – brukes av query_table.
- load():        blokker for de vanlige verktøyene. Hvert ark blir én seksjon
                 ("## Ark: Salg") med en markdown-tabell, laget fra load_tables().

Første ikke-tomme rad regnes som overskriftsrad. Hver rad får Excel-radnummeret
sitt i en egen "Rad"-kolonne, så modellen kan si "rad 57" og brukeren finne den.
Store ark deles i biter på tables.ROWS_PER_PART rader ("### Rad 2–101").
Cellekommentarer tas med under tabellen – der står ofte forklaringen på et tall.
"""

import datetime
import zipfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from docxreader.blocks import Block, DocumentError
from docxreader.tables import Table, make_table, table_blocks

SHEET_LEVEL = 2  # "## Ark: navn" – samme nivå som lysbilder og Heading 1


def _format(value, formula) -> str:
    """Én celleverdi -> tekst."""
    if value is None:
        # Formelen har ikke noe lagret resultat (filen er aldri åpnet i Excel
        # etter at formelen ble skrevet). Vis formelen i stedet for ingenting.
        return formula if isinstance(formula, str) and formula.startswith("=") else ""
    if isinstance(value, datetime.datetime) and value.time() == datetime.time(0):
        return value.date().isoformat()  # 2026-09-28, ikke 2026-09-28 00:00:00
    if isinstance(value, float) and value.is_integer():
        return str(int(value))  # 1200.0 -> 1200
    return str(value)


def _sheet_rows(values_sheet, formulas_sheet) -> list[tuple[int, list[str]]]:
    """Alle ikke-tomme rader som (Excel-radnummer, celletekster).
    Tomme kolonner helt til høyre fjernes."""
    rows = []
    for values, formulas in zip(values_sheet.iter_rows(), formulas_sheet.iter_rows()):
        cells = [_format(v.value, f.value) for v, f in zip(values, formulas)]
        if any(cells):
            rows.append((values[0].row, cells))
    if not rows:
        return []
    # ws.max_column kan være for stor (formatering teller); kutt tomme kolonner.
    width = max(max((i + 1 for i, c in enumerate(cells) if c), default=0) for _, cells in rows)
    return [(number, cells[:width]) for number, cells in rows]


def load_tables(path: Path) -> list[Table]:
    """Hvert ark som en Table – også tomme og skjulte ark."""
    try:
        # To utgaver av samme fil: én med lagrede resultater av formler
        # (data_only=True) og én med selve formlene.
        values_book = load_workbook(path, data_only=True)
        formulas_book = load_workbook(path)
    except (zipfile.BadZipFile, InvalidFileException, KeyError):
        # .xlsx er en zip-fil med XML inni; ødelagt zip eller manglende deler.
        raise DocumentError(f"'{path}' er ikke en gyldig Excel-fil (ødelagt eller feil format).")

    tables = []
    for values_sheet in values_book.worksheets:
        formulas_sheet = formulas_book[values_sheet.title]
        comments = [
            f"{cell.coordinate}: {' '.join(cell.comment.text.split())}"
            for row in formulas_sheet.iter_rows()
            for cell in row
            if cell.comment
        ]
        tables.append(make_table(
            values_sheet.title,
            _sheet_rows(values_sheet, formulas_sheet),
            hidden=values_sheet.sheet_state != "visible",
            comments=comments,
        ))
    return tables


def load(path: Path) -> tuple[list[Block], str]:
    blocks: list[Block] = []
    summary = []  # én linje per ark, med kolonnenavn – nyttig før query_table
    for table in load_tables(path):
        if not table.header and not table.comments:
            summary.append(f"- {table.name}: tomt")
            continue
        summary.append(f"- {table.name}: {len(table.rows)} rader; kolonner: {', '.join(table.header)}")
        title = f"Ark: {table.name}" + (" (skjult)" if table.hidden else "")
        blocks.extend(table_blocks(table, title, SHEET_LEVEL))

    meta = f"Arbeidsbok med {len(summary)} ark:\n" + "\n".join(summary)
    return blocks, meta
