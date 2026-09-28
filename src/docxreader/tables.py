"""Tabelldata (rader og kolonner) for regneark og CSV – og spørringene i query_table.

Lesere for tabellformater har, i tillegg til `load()` (blokker for de vanlige
verktøyene), en `load_tables()` som gir `Table`-objekter. Spørringer (filtrer,
sorter, summer) gjøres her i Python, så modellen slipper å lese alle radene.

Alle celler er tekst. Tall tolkes når vi trenger dem (`parse_number`), slik at
Excel og CSV behandles likt – også norske tall som "1 250,50".
`split_tables` finner tabellene i et ark (titler, flere tabeller, fotnoter).
"""

import re
from dataclasses import dataclass, field

from docxreader.blocks import Block, DocumentError, count_words, heading, table_block

ROWS_PER_PART = 100  # store tabeller deles i biter på så mange rader ("### Rad 2–101")


@dataclass
class Table:
    name: str  # arknavn i Excel ("Salg – Budsjett Q1" hvis arket har flere tabeller), "Tabell" i CSV
    header: list[str]
    rows: list[tuple[int, list[str]]]  # (radnummer i filen, celler)
    hidden: bool = False
    comments: list[str] = field(default_factory=list)  # "C3: tekst"
    sheet: str = ""  # arket tabellen står i
    title: str = ""  # tekst over tabellen, f.eks. "Budsjett Q1"
    notes: str = ""  # tekst under tabellen, f.eks. "Kilde: regnskapet"


def _filled(cells: list[str]) -> int:
    return sum(1 for c in cells if c.strip())


def make_table(name: str, numbered_rows: list[tuple[int, list[str]]], **kwargs) -> Table:
    """Første rad blir overskriftsrad. Tomme kolonner i kantene fjernes (tabellen
    kan starte i kolonne C), tomme overskrifter får navnet "Kolonne N", og korte
    rader fylles ut så alle rader har like mange celler."""
    if not numbered_rows:
        return Table(name, [], [], **kwargs)
    width = max(len(cells) for _, cells in numbered_rows)
    padded = [(n, cells + [""] * (width - len(cells))) for n, cells in numbered_rows]
    used = [i for i in range(width) if any(cells[i].strip() for _, cells in padded)]
    first, last = used[0], used[-1] + 1
    padded = [(n, cells[first:last]) for n, cells in padded]
    header = [h.strip() or f"Kolonne {i + 1}" for i, h in enumerate(padded[0][1])]
    return Table(name, header, padded[1:], **kwargs)


# --- Finne tabellene i et ark ------------------------------------------------
#
# Et regneark er et rutenett, ikke en tabell. Folk skriver ofte en tittel over
# tabellen, flere tabeller under hverandre og en kildehenvisning nederst.
# Vi leser oppsettet slik et menneske gjør:
#   - tomme rader skiller områder,
#   - rader med bare én utfylt celle over en overskriftsrad (ren tekst) er titler,
#   - et område som starter med data (tall/datoer) og ikke har tittel, er
#     fortsettelsen av forrige tabell (en tom rad midt i tabellen),
#   - tekst etter siste tabell er en fotnote.

ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}")


def _looks_like_data(cells: list[str]) -> bool:
    return any(parse_number(c) is not None or ISO_DATE.match(c.strip()) for c in cells)


def split_tables(name: str, numbered_rows, split_on_gaps: bool = True, **kwargs) -> list[Table]:
    """Del radene i et ark i én eller flere tabeller, med titler og fotnoter.

    split_on_gaps=False (CSV): tomme rader skiller ikke tabeller, men en
    tittellinje øverst gjenkjennes fortsatt.
    """
    if not numbered_rows:
        return [Table(name, [], [], sheet=name, **kwargs)]
    # Har ingen rad mer enn én utfylt celle, er det en tabell med én kolonne –
    # da finnes ingen titler å kjenne igjen.
    if max(_filled(cells) for _, cells in numbered_rows) <= 1:
        return [make_table(name, numbered_rows, sheet=name, **kwargs)]

    # 1. Områder: rader som følger rett etter hverandre.
    regions: list[list] = []
    for n, cells in numbered_rows:
        if regions and (not split_on_gaps or n == regions[-1][-1][0] + 1):
            regions[-1].append((n, cells))
        else:
            regions.append([(n, cells)])

    # 2. Hvert område: titler (rader med én celle) og så selve tabellen.
    parts: list[dict] = []
    pending_titles: list[str] = []
    for region in regions:
        i = 0
        while i < len(region) and _filled(region[i][1]) <= 1:
            i += 1
        if 0 < i < len(region) and _looks_like_data(region[i][1]):
            # Rett under kommer data, ikke en overskriftsrad: da var den korte
            # raden selve overskriftsraden (med tomme overskriftsceller).
            i = 0
        texts = [" ".join(c.strip() for c in cells if c.strip()) for _, cells in region[:i]]
        rest = region[i:]
        if not rest:  # bare tekst: tittel til neste tabell (eller fotnote)
            pending_titles.extend(texts)
            continue
        titles, pending_titles = pending_titles + texts, []
        if parts and not titles and _looks_like_data(rest[0][1]):
            parts[-1]["rows"].extend(rest)  # tom rad midt i en tabell
        else:
            parts.append({"titles": titles, "rows": rest})

    # 3. Lag Table-objektene.
    tables = []
    for number, part in enumerate(parts, start=1):
        if len(parts) == 1:
            table_name = name
        else:
            label = part["titles"][-1] if part["titles"] else f"tabell {number}"
            table_name = f"{name} – {label}"
        table = make_table(table_name, part["rows"], sheet=name, **kwargs)
        table.title = " – ".join(part["titles"])
        tables.append(table)
    tables[-1].notes = " ".join(pending_titles)  # tekst etter siste tabell
    return tables


# --- Tabell -> blokker (for read_document, read_section, search_document) ---


def table_blocks(table: Table, title: str, level: int = 2) -> list[Block]:
    """Overskrift + tabell med "Rad"-kolonne. Store tabeller deles i biter med
    underoverskrifter, og overskriftsraden gjentas i hver bit. Tittel over og
    fotnote under tabellen tas med som vanlige avsnitt."""
    blocks = [heading(level, title)]
    if table.title and table.title != title:
        blocks.append(Block(table.title, count_words(table.title)))
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
    if table.notes:
        blocks.append(Block(table.notes, count_words(table.notes)))
    if table.comments:
        text = "Kommentarer:\n" + "\n".join(f"- {c}" for c in table.comments)
        blocks.append(Block(text, count_words(text)))
    return blocks


# --- Tall og kolonner --------------------------------------------------------

CURRENCY_BEFORE = re.compile(r"^(kr\.?|nok|[$€£])\s*", re.IGNORECASE)
CURRENCY_AFTER = re.compile(r"\s*(kr\.?|nok|%|[$€£]|,-)$", re.IGNORECASE)


def parse_number(text: str) -> float | None:
    """Tekst -> tall, eller None hvis det ikke er et tall.

    Godtar:
    - "1250", "-3", "1250.5" og norsk "1 250,50" (mellomrom = tusenskille)
    - engelsk "1,250,000.50" og tysk/norsk "1.250.000,50": finnes både komma og
      punktum, er det SISTE desimaltegnet og det andre tusenskille
    - samme tegn flere ganger = tusenskille: "1.250.000", "1,250,000"
    - valuta og prosent: "kr 1 250", "1 250 kr", "25 %", norsk "1 250,-"
    Ett enkelt komma eller punktum leses som desimaltegn ("1,250" = 1,25).
    Datoer som "2026-09-01" og "12.05.2026" er IKKE tall.
    """
    t = text.strip().replace("−", "-")  # typografisk minus
    t = CURRENCY_AFTER.sub("", CURRENCY_BEFORE.sub("", t))
    t = re.sub(r"[\s  ']", "", t)  # mellomrom og ' som tusenskille

    if "," in t and "." in t:
        decimal = "," if t.rfind(",") > t.rfind(".") else "."
        thousands = "." if decimal == "," else ","
        if not re.fullmatch(rf"-?\d{{1,3}}(\{thousands}\d{{3}})*\{decimal}\d+", t):
            return None
        t = t.replace(thousands, "").replace(decimal, ".")
    elif t.count(",") > 1 or t.count(".") > 1:
        separator = "," if t.count(",") > 1 else "."
        if not re.fullmatch(rf"-?\d{{1,3}}(\{separator}\d{{3}})+", t):
            return None  # f.eks. datoen 12.05.2026
        t = t.replace(separator, "")
    else:
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
    """Gjør en betingelse om til en funksjon som sier ja/nei for en rad.

    En betingelse kan ha alternativer skilt med " OR ", og raden tas med hvis
    minst ett stemmer: "Avdeling = IT OR Avdeling = HR". (Flere elementer i
    where-listen må ALLE stemme – det er "og".)
    """
    checks = [_single_filter(table, part) for part in condition.split(" OR ")]
    return lambda cells: any(check(cells) for check in checks)


def _single_filter(table: Table, condition: str):
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


# --- Beregnede kolonner ------------------------------------------------------
#
# "Sum = Antall * Pris per stk" lager en ny kolonne "Sum". Vi bruker IKKE
# Pythons eval(): uttrykket kommer fra modellen, og eval kunne kjørt hva som
# helst. I stedet har vi en liten parser som bare forstår tall, kolonner,
# + - * / og parenteser:
#
#   1. Kolonnenavn byttes ut med plassholdere (lengste navn først). Da kan et
#      navn selv inneholde parenteser eller bindestrek: "Beløp (kr)".
#   2. Resten deles i biter ("tokens"): tall, kolonner, operatorer, parenteser.
#   3. Bitene settes sammen til et tre med vanlig regnerekkefølge
#      (parenteser først, så * og /, så + og -), og treet regnes ut per rad.

PLACEHOLDER = "\x00"  # tegn som aldri står i et uttrykk


def _tokenize(table: Table, formula: str) -> list[tuple[str, object]]:
    text = formula
    for index in sorted(range(len(table.header)), key=lambda i: -len(table.header[i])):
        pattern = re.compile(re.escape(table.header[index]), re.IGNORECASE)
        text = pattern.sub(f"{PLACEHOLDER}{index}{PLACEHOLDER}", text)

    tokens: list[tuple[str, object]] = []
    i = 0

    def binary_position() -> bool:  # står vi etter et tall/en kolonne/")"?
        return bool(tokens) and tokens[-1][0] in ("tall", "kolonne", ")")

    while i < len(text):
        char = text[i]
        if char.isspace():
            i += 1
        elif char == PLACEHOLDER:
            end = text.index(PLACEHOLDER, i + 1)
            tokens.append(("kolonne", int(text[i + 1 : end])))
            i = end + 1
        elif char in "()":
            tokens.append((char, None))
            i += 1
        elif char in "+-*/" and (binary_position() or char in "+-"):
            # Binær operator ("A - B"), eller fortegn foran noe ("-5", "-(A + B)").
            tokens.append(("op" if binary_position() else "fortegn", char))
            i += 1
        elif char in "*/":  # * eller / uten noe foran
            raise DocumentError(f"forstår ikke uttrykket '{formula}' (uventet '{char}').")
        else:
            # Et tall eller et (delvis) kolonnenavn. Det slutter ved parentes,
            # plassholder, eller en operator med mellomrom foran – så
            # bindestreken i "Pris-per-stk" ikke leses som minus.
            start = i
            while i < len(text) and text[i] not in "()" + PLACEHOLDER:
                if text[i] in "+-*/" and i > start and text[i - 1].isspace():
                    break
                i += 1
            operand = text[start:i].strip()
            number = parse_number(operand)
            tokens.append(("tall", number) if number is not None else ("kolonne", find_column(table, operand)))
    return tokens


def _parse(tokens: list, formula: str):
    """Tokens -> tre. uttrykk = ledd (+|- ledd)* ; ledd = faktor (*|/ faktor)* ;
    faktor = tall | kolonne | fortegn faktor | ( uttrykk )."""
    position = 0

    def fail(reason: str):
        raise DocumentError(f"forstår ikke uttrykket '{formula}' ({reason}).")

    def peek():
        return tokens[position] if position < len(tokens) else (None, None)

    def take():
        nonlocal position
        position += 1
        return tokens[position - 1]

    def factor():
        kind, value = peek()
        if kind in ("tall", "kolonne"):
            return take()
        if kind == "fortegn":
            take()
            inner = factor()
            return ("op", "-", ("tall", 0.0), inner) if value == "-" else inner
        if kind == "(":
            take()
            inner = expression()
            if peek()[0] != ")":
                fail("mangler )")
            take()
            return inner
        fail("mangler et tall eller en kolonne" if kind is None else f"uventet '{value or kind}'")

    def term():
        node = factor()
        while peek()[0] == "op" and peek()[1] in "*/":
            node = ("op", take()[1], node, factor())
        return node

    def expression():
        node = term()
        while peek()[0] == "op" and peek()[1] in "+-":
            node = ("op", take()[1], node, term())
        return node

    tree = expression()
    if position < len(tokens):
        fail("for mange )" if tokens[position][0] == ")" else f"uventet '{tokens[position][1] or tokens[position][0]}'")
    return tree


def _evaluate(node, cells: list[str]) -> float | None:
    """Regn ut treet for én rad. None hvis en verdi mangler/ikke er tall, eller ved deling på 0."""
    kind = node[0]
    if kind == "tall":
        return node[1]
    if kind == "kolonne":
        return parse_number(cells[node[1]])
    _, op, left, right = node
    a, b = _evaluate(left, cells), _evaluate(right, cells)
    if a is None or b is None or (op == "/" and b == 0):
        return None
    return {"+": a + b, "-": a - b, "*": a * b, "/": a / b if b else 0}[op]


def add_calculated_column(table: Table, expression: str) -> Table:
    """Ny tabell med én ekstra kolonne regnet ut fra "Navn = uttrykk".

    Uttrykket kan ha kolonnenavn, tall, + - * / og parenteser, f.eks.
    "Total = (Antall * Pris) * 1.25". Operatorer mellom (delvise) kolonnenavn
    trenger mellomrom rundt seg, så "Pris-per-stk" ikke leses som en minus.
    Mangler en verdi, eller er den ikke et tall, blir cellen tom. Deling på 0
    gir også tom celle.
    """
    name, sep, formula = expression.partition("=")
    name, formula = name.strip(), formula.strip()
    if not sep or not name or not formula:
        raise DocumentError(
            f"forstår ikke beregningen '{expression}'. Bruk 'Navn = uttrykk', "
            "f.eks. 'Sum = Antall * Pris' eller 'Med mva = (Pris + Frakt) * 1.25'."
        )
    tree = _parse(_tokenize(table, formula), formula)

    rows = []
    for n, cells in table.rows:
        value = _evaluate(tree, cells)
        rows.append((n, [*cells, "" if value is None else format_number(value)]))
    return Table(table.name, [*table.header, name], rows, table.hidden, table.comments,
                 table.sheet, table.title, table.notes)


# --- Spørringen ----------------------------------------------------------------


def _sort_key(cell: str):
    """Tall sorteres som tall, tekst som tekst; tall kommer før tekst."""
    number = parse_number(cell)
    return (0, number, "") if number is not None else (1, 0, cell.casefold())


def query(
    table: Table,
    calculate: list[str],
    where: list[str],
    sort_by: str,
    descending: bool,
    columns: list[str],
    limit: int,
    group_by: str,
    aggregate: str | None,
    value_column: str,
) -> str:
    # Beregnede kolonner først, så de kan brukes i where, sort_by og aggregate.
    for expression in calculate:
        table = add_calculated_column(table, expression)

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
