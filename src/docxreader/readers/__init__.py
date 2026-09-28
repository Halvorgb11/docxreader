"""Menyen over filtyper vi kan lese.

Hver leser er en funksjon `load(path) -> (blokker, metatekst)` som gjør en fil
om til `Block`-er. Verktøyene i tools.py kaller `load_document`, som velger
leser ut fra filendelsen.

Tabellformater (regneark, CSV) har i tillegg `load_tables(path) -> [Table]`,
som query_table bruker for å filtrere, sortere og summere rader.

Legge til en ny filtype (se også skillen .claude/skills/ny-filleser):
1. Lag readers/<format>.py med en `load(path)`-funksjon (og ev. `load_tables`).
2. Legg filendelsen inn i READERS (og ev. TABLE_READERS) under.
3. Nevn filtypen i docstringene i tools.py (en test sjekker at du ikke glemmer det).
"""

from pathlib import Path

from docxreader.blocks import Block, DocumentError
from docxreader.readers import csvfile, excel, mail, powerpoint, text, word
from docxreader.tables import Table

READERS = {
    ".docx": word.load,
    ".pptx": powerpoint.load,
    ".xlsx": excel.load,
    ".csv": csvfile.load,
    ".eml": mail.load_eml,
    ".msg": mail.load_msg,
    ".md": text.load_markdown,
    ".txt": text.load_plain,
}

TABLE_READERS = {
    ".xlsx": excel.load_tables,
    ".csv": csvfile.load_tables,
}


def _pick_reader(path: str, readers: dict):
    """Finn filen og riktig leser. Kaster DocumentError med en forklaring
    hvis filen mangler eller filtypen ikke støttes."""
    file = Path(path)
    if not file.is_file():
        raise DocumentError(f"fant ingen fil på '{path}'.")
    reader = readers.get(file.suffix.lower())
    if reader is None:
        supported = ", ".join(readers)
        raise DocumentError(
            f"filtypen til '{path}' støttes ikke ({file.suffix or 'ingen filendelse'}). "
            f"Støttede filtyper: {supported}."
        )
    return reader, file


def _run(reader, file: Path):
    """Kjør en leser med et sikkerhetsnett: et uventet unntak (en feil i et
    bibliotek eller i leseren) blir en DocumentError i stedet for å krasje
    agenten. Modellen får da en "Feil: …"-tekst den kan forholde seg til."""
    try:
        return reader(file)
    except DocumentError:
        raise
    except Exception as e:
        raise DocumentError(f"uventet feil ved lesing av '{file}' ({type(e).__name__}: {e}).") from e


def load_document(path: str) -> tuple[list[Block], str]:
    reader, file = _pick_reader(path, READERS)
    return _run(reader, file)


def load_table(path: str, sheet: str = "") -> Table:
    """Én tabell fra et regneark eller en CSV-fil.

    `sheet` velger ark (uten hensyn til store/små bokstaver, og "Ark: " foran
    er lov, siden modellen ofte kopierer overskriften fra innholdsfortegnelsen).
    Tomt `sheet` går bare når filen har ett ark med innhold.
    """
    reader, file = _pick_reader(path, TABLE_READERS)
    tables = [t for t in _run(reader, file) if t.header]
    if not tables:
        raise DocumentError(f"'{path}' inneholder ingen tabelldata.")

    names = ", ".join(t.name for t in tables)
    wanted = sheet.strip().removeprefix("Ark:").removeprefix("ark:").strip().casefold()
    wanted = wanted.removesuffix("(skjult)").strip()
    if len(tables) == 1:
        return tables[0]  # bare én tabell (f.eks. CSV): arknavnet spiller ingen rolle
    if not wanted:
        raise DocumentError(f"filen har flere ark; oppgi sheet. Ark: {names}.")
    for exact in (True, False):
        hits = [t for t in tables if (t.name.casefold() == wanted if exact else wanted in t.name.casefold())]
        if len(hits) == 1:
            return hits[0]
    if len(hits) > 1:
        options = ", ".join(t.name for t in hits)
        raise DocumentError(f"'{sheet}' passer flere tabeller: {options}. Oppgi tabellnavnet, f.eks. '{hits[0].name}'.")
    raise DocumentError(f"fant ikke ett entydig ark '{sheet}'. Ark med innhold: {names}.")
