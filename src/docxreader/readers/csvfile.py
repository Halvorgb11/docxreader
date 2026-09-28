"""Leser for CSV-filer (.csv), med `csv` fra standardbiblioteket.

En CSV-fil er én tabell i ren tekst. Den blir én seksjon ("## Tabell"), delt i
biter på tables.ROWS_PER_PART rader når den er stor – akkurat som et Excel-ark.
"Rad" er linjenummeret i filen (overskriftsraden er rad 1), så lenge ingen
celler inneholder linjeskift.

To vanlige norske fallgruver:
- Skilletegn: norsk Excel lagrer CSV med ";" (fordi "," er desimaltegn).
  Vi gjetter skilletegnet ut fra overskriftslinjen.
- Tegnkoding: eldre filer er ofte cp1252, ikke UTF-8 (se text.read_text).
"""

import csv
import io
from pathlib import Path

from docxreader.blocks import Block
from docxreader.readers.text import read_text
from docxreader.tables import Table, make_table, table_blocks

DELIMITERS = [";", ",", "\t", "|"]
DELIMITER_NAMES = {";": "semikolon", ",": "komma", "\t": "tabulator", "|": "loddrett strek"}


def _guess_delimiter(text: str) -> str:
    """Skilletegnet som forekommer oftest i første linje. (Overskriftslinjen
    inneholder sjelden desimaltall, så "1,5" forvirrer ikke gjettingen.)"""
    first_line = text.split("\n", 1)[0]
    return max(DELIMITERS, key=first_line.count)


def _read(path: Path) -> tuple[Table, str]:
    text = read_text(path)
    delimiter = _guess_delimiter(text)
    # csv-modulen håndterer anførselstegn: "Hansen; Ola" er én celle.
    records = csv.reader(io.StringIO(text), delimiter=delimiter)
    rows = [
        (number, [cell.strip() for cell in cells])
        for number, cells in enumerate(records, start=1)
        if any(cell.strip() for cell in cells)
    ]
    return make_table("Tabell", rows), delimiter


def load_tables(path: Path) -> list[Table]:
    return [_read(path)[0]]


def load(path: Path) -> tuple[list[Block], str]:
    table, delimiter = _read(path)
    if not table.header:
        return [], ""
    meta = (
        f"CSV-fil med {len(table.rows)} rader og {len(table.header)} kolonner "
        f"(skilletegn: {DELIMITER_NAMES[delimiter]}).\nKolonner: {', '.join(table.header)}"
    )
    return table_blocks(table, "Tabell"), meta
