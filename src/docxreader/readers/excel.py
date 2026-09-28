"""Leser for Excel-arbeidsbøker (.xlsx), med openpyxl.

Hvert ark blir én seksjon ("## Ark: Salg") med innholdet som en markdown-tabell.
Første ikke-tomme rad regnes som overskriftsrad. Hver rad får Excel-radnummeret
sitt i en egen "Rad"-kolonne, så modellen kan si "rad 57" og brukeren finne den.

Store ark deles i biter på ROWS_PER_PART rader ("### Rad 2–101"). Da virker
innholdsfortegnelsen og størrelsesgrensen i read_section som for kapitler i Word,
og modellen kan lese ett utsnitt om gangen. Overskriftsraden gjentas i hver bit.

Cellekommentarer tas med under tabellen – der står ofte forklaringen på et tall.
"""

import datetime
import zipfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from docxreader.blocks import Block, DocumentError, count_words, heading, table_block

SHEET_LEVEL = 2  # "## Ark: navn" – samme nivå som lysbilder og Heading 1
ROWS_PER_PART = 100


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


def _sheet_blocks(name: str, hidden: bool, rows, comments: list[str]) -> list[Block]:
    title = f"Ark: {name}" + (" (skjult)" if hidden else "")
    blocks = [heading(SHEET_LEVEL, title)]

    if rows:
        (_, header), data = rows[0], rows[1:]
        header_row = ["Rad", *header]

        def table(part):
            return table_block([header_row, *([str(n), *cells] for n, cells in part)])

        if len(data) <= ROWS_PER_PART:
            blocks.append(table(data))
        else:
            for start in range(0, len(data), ROWS_PER_PART):
                part = data[start : start + ROWS_PER_PART]
                blocks.append(heading(SHEET_LEVEL + 1, f"Rad {part[0][0]}–{part[-1][0]}"))
                blocks.append(table(part))

    if comments:
        text = "Kommentarer:\n" + "\n".join(f"- {c}" for c in comments)
        blocks.append(Block(text, count_words(text)))
    return blocks


def load(path: Path) -> tuple[list[Block], str]:
    try:
        # To utgaver av samme fil: én med lagrede resultater av formler
        # (data_only=True) og én med selve formlene.
        values_book = load_workbook(path, data_only=True)
        formulas_book = load_workbook(path)
    except (zipfile.BadZipFile, InvalidFileException, KeyError):
        # .xlsx er en zip-fil med XML inni; ødelagt zip eller manglende deler.
        raise DocumentError(f"'{path}' er ikke en gyldig Excel-fil (ødelagt eller feil format).")

    blocks: list[Block] = []
    summary = []
    for values_sheet in values_book.worksheets:
        formulas_sheet = formulas_book[values_sheet.title]
        rows = _sheet_rows(values_sheet, formulas_sheet)
        hidden = values_sheet.sheet_state != "visible"
        comments = [
            f"{cell.coordinate}: {' '.join(cell.comment.text.split())}"
            for row in formulas_sheet.iter_rows()
            for cell in row
            if cell.comment
        ]
        if not rows and not comments:
            summary.append(f"{values_sheet.title} (tomt)")
            continue
        summary.append(f"{values_sheet.title} ({max(len(rows) - 1, 0)} rader)")
        blocks.extend(_sheet_blocks(values_sheet.title, hidden, rows, comments))

    meta = f"Arbeidsbok med {len(summary)} ark: " + ", ".join(summary) + "."
    return blocks, meta
