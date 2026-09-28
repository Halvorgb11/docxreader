"""Tabelldata (rader og kolonner) for regneark og CSV – og spørringene i query_table.

Lesere for tabellformater har, i tillegg til `load()` (blokker for de vanlige
verktøyene), en `load_tables()` som gir `Table`-objekter. Spørringer (filtrer,
sorter, summer) gjøres her i Python, så modellen slipper å lese alle radene.

Alle celler er tekst. Tall tolkes når vi trenger dem (`parse_number`), slik at
Excel og CSV behandles likt – også norske tall som "1 250,50".
"""

import re
from dataclasses import dataclass, field

from docxreader.blocks import Block, DocumentError, count_words, heading, table_block

ROWS_PER_PART = 100  # store tabeller deles i biter på så mange rader ("### Rad 2–101")


@dataclass
class Table:
    name: str  # arknavn i Excel, "Tabell" i CSV
    header: list[str]
    rows: list[tuple[int, list[str]]]  # (radnummer i filen, celler)
    hidden: bool = False
    comments: list[str] = field(default_factory=list)  # "C3: tekst"


def make_table(name: str, numbered_rows: list[tuple[int, list[str]]], **kwargs) -> Table:
    """Første rad blir overskriftsrad. Tomme overskrifter får navnet "Kolonne N",
    og korte rader fylles ut så alle rader har like mange celler."""
    if not numbered_rows:
        return Table(name, [], [], **kwargs)
    width = max(len(cells) for _, cells in numbered_rows)
    header = numbered_rows[0][1] + [""] * (width - len(numbered_rows[0][1]))
    header = [h.strip() or f"Kolonne {i + 1}" for i, h in enumerate(header)]
    rows = [(n, cells + [""] * (width - len(cells))) for n, cells in numbered_rows[1:]]
    return Table(name, header, rows, **kwargs)


# --- Tabell -> blokker (for read_document, read_section, search_document) ---


def table_blocks(table: Table, title: str, level: int = 2) -> list[Block]:
    """Overskrift + tabell med "Rad"-kolonne. Store tabeller deles i biter med
    underoverskrifter, og overskriftsraden gjentas i hver bit."""
    blocks = [heading(level, title)]
    if table.header:
        header_row = ["Rad", *table.header]

        def part_block(part):
            return table_block([header_row, *([str(n), *cells] for n, cells in part)])

        if len(table.rows) <= ROWS_PER_PART:
            blocks.append(part_block(table.rows))
        else:
            for start in range(0, len(table.rows), ROWS_PER_PART):
                part = table.rows[start : start + ROWS_PER_PART]
                blocks.append(heading(level + 1, f"Rad {part[0][0]}–{part[-1][0]}"))
                blocks.append(part_block(part))
    if table.comments:
        text = "Kommentarer:\n" + "\n".join(f"- {c}" for c in table.comments)
        blocks.append(Block(text, count_words(text)))
    return blocks


# --- Tall og kolonner --------------------------------------------------------


def parse_number(text: str) -> float | None:
    """Tekst -> tall, eller None hvis det ikke er et tall.

    Godtar "1250", "1250.5", norsk "1 250,50" (mellomrom som tusenskille og
    komma som desimaltegn) og minus. Datoer som "2026-09-01" er IKKE tall.
    """
    t = text.strip().replace(" ", "").replace(" ", "")
    if re.fullmatch(r"-?\d+,\d+", t):
        t = t.replace(",", ".")
    if re.fullmatch(r"-?\d+(\.\d+)?", t):
        return float(t)
    return None


def format_number(value: float) -> str:
    """1200.0 -> "1200", 3.14159 -> "3.14"."""
    return str(int(value)) if float(value).is_integer() else f"{value:.2f}"


def find_column(table: Table, name: str) -> int:
    """Kolonneindeks for `name`: eksakt treff (uten hensyn til store/små
    bokstaver) først, deretter entydig delvis treff."""
    wanted = name.strip().casefold()
    names = [h.casefold() for h in table.header]
    if wanted in names:
        return names.index(wanted)
    partial = [i for i, n in enumerate(names) if wanted and wanted in n]
    if len(partial) == 1:
        return partial[0]
    raise DocumentError(
        f"fant ingen entydig kolonne '{name}' i '{table.name}'. "
        f"Kolonner: {', '.join(table.header)}."
    )


# --- Filtrering --------------------------------------------------------------

# "Kolonne OPERATOR verdi". Lengste operatorer først, så ">=" ikke leses som ">".
CONDITION = re.compile(r"^\s*(.+?)\s*(>=|<=|!=|=|>|<|~)\s*(.*?)\s*$")


def make_filter(table: Table, condition: str):
    """Gjør "Beløp (kr) > 10000" om til en funksjon som sier ja/nei for en rad.

    = og != sammenligner tekst uten hensyn til store/små bokstaver (eller tall).
    > < >= <= sammenligner tall hvis begge sider er tall, ellers tekst
    (fungerer for datoer på formen 2026-09-28). ~ betyr "inneholder".
    """
    m = CONDITION.match(condition)
    if not m:
        raise DocumentError(
            f"forstår ikke betingelsen '{condition}'. Bruk 'Kolonne operator verdi', "
            "f.eks. 'Avdeling = IT' eller 'Beløp > 1000'. Operatorer: = != > < >= <= ~"
        )
    column_name, op, value = m.groups()
    col = find_column(table, column_name)
    value = value.strip("'\"")
    wanted_number = parse_number(value)

    def keep(cells: list[str]) -> bool:
        cell = cells[col]
        number = parse_number(cell)
        if op == "~":
            return value.casefold() in cell.casefold()
        if number is not None and wanted_number is not None:
            a, b = number, wanted_number
        else:
            a, b = cell.casefold(), value.casefold()
            if op in (">", "<", ">=", "<=") and not cell:
                return False  # tomme celler er verken større eller mindre
        return {
            "=": a == b, "!=": a != b, ">": a > b, "<": a < b, ">=": a >= b, "<=": a <= b,
        }[op]

    return keep


# --- Spørringen ----------------------------------------------------------------


def _sort_key(cell: str):
    """Tall sorteres som tall, tekst som tekst; tall kommer før tekst."""
    number = parse_number(cell)
    return (0, number, "") if number is not None else (1, 0, cell.casefold())


def query(
    table: Table,
    where: list[str],
    sort_by: str,
    descending: bool,
    columns: list[str],
    limit: int,
    group_by: str,
    aggregate: str | None,
    value_column: str,
) -> str:
    rows = table.rows
    for condition in where:
        keep = make_filter(table, condition)
        rows = [(n, cells) for n, cells in rows if keep(cells)]

    description = f"'{table.name}': {len(rows)} av {len(table.rows)} rader"
    if where:
        description += " der " + " og ".join(where)

    if aggregate or group_by:
        return description + ".\n\n" + _aggregate(table, rows, group_by, aggregate or "count", value_column)

    if sort_by:
        col = find_column(table, sort_by)
        # Tomme celler legges alltid sist, også ved synkende sortering.
        filled = [r for r in rows if r[1][col].strip()]
        empty = [r for r in rows if not r[1][col].strip()]
        rows = sorted(filled, key=lambda r: _sort_key(r[1][col]), reverse=descending) + empty
        description += f", sortert etter {table.header[col]}" + (" (synkende)" if descending else "")

    shown_cols = [find_column(table, c) for c in columns] if columns else range(len(table.header))
    shown = rows[:limit]
    result = [["Rad", *(table.header[i] for i in shown_cols)]]
    result += [[str(n), *(cells[i] for i in shown_cols)] for n, cells in shown]

    text = description + ".\n\n" + (table_block(result).markdown if shown else "(ingen rader)")
    if len(rows) > len(shown):
        text += f"\n\nViser {len(shown)} av {len(rows)} rader. Øk limit, filtrer mer, eller bruk aggregate."
    return text


def _aggregate(table: Table, rows, group_by: str, aggregate: str, value_column: str) -> str:
    """Tell/summer/snitt/min/maks – totalt, eller per verdi i group_by-kolonnen."""
    if aggregate != "count" and not value_column:
        raise DocumentError(f"aggregate='{aggregate}' trenger value_column (kolonnen som skal regnes på).")
    value_col = find_column(table, value_column) if value_column else None
    group_col = find_column(table, group_by) if group_by else None

    # Samle verdiene per gruppe. dict husker rekkefølgen grupper dukker opp i.
    groups: dict[str, list[float]] = {}
    counts: dict[str, int] = {}
    skipped = 0
    for _, cells in rows:
        key = cells[group_col] if group_col is not None else "Alle"
        counts[key] = counts.get(key, 0) + 1
        groups.setdefault(key, [])
        if value_col is not None and cells[value_col].strip():  # tomme celler hoppes over
            number = parse_number(cells[value_col])
            if number is None:
                skipped += 1
            else:
                groups[key].append(number)

    def compute(key: str) -> float | None:
        values = groups[key]
        if aggregate == "count":
            return counts[key]
        if not values:
            return None
        return {"sum": sum, "min": min, "max": max, "avg": lambda v: sum(v) / len(v)}[aggregate](values)

    labels = {"count": "Antall", "sum": "Sum", "avg": "Snitt", "min": "Minste", "max": "Største"}
    label = labels[aggregate] + (f" av {table.header[value_col]}" if value_col is not None and aggregate != "count" else "")
    first = table.header[group_col] if group_col is not None else "Utvalg"

    results = [(key, compute(key)) for key in groups]
    # Største verdi først – det er som regel det man spør etter.
    results.sort(key=lambda kv: (kv[1] is None, -(kv[1] or 0)))
    table_rows = [[first, "Antall rader", label] if aggregate != "count" else [first, label]]
    for key, value in results:
        shown = "–" if value is None else format_number(value)
        table_rows.append([key or "(tom)", str(counts[key]), shown] if aggregate != "count" else [key or "(tom)", shown])

    text = table_block(table_rows).markdown
    if skipped:
        text += f"\n\n{skipped} celler i '{table.header[value_col]}' var ikke tall og ble hoppet over."
    return text
