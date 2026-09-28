"""Menyen over filtyper vi kan lese.

Hver leser er en funksjon `load(path) -> (blokker, metatekst)` som gjør en fil
om til `Block`-er. Verktøyene i tools.py kaller `load_document`, som velger
leser ut fra filendelsen.

Legge til en ny filtype:
1. Lag readers/<format>.py med en `load(path)`-funksjon.
2. Legg filendelsen inn i READERS under.
3. Nevn filtypen i docstringene i tools.py (en test sjekker at du ikke glemmer det).
"""

from pathlib import Path

from docxreader.blocks import Block, DocumentError
from docxreader.readers import excel, powerpoint, text, word

READERS = {
    ".docx": word.load,
    ".pptx": powerpoint.load,
    ".xlsx": excel.load,
    ".md": text.load_markdown,
    ".txt": text.load_plain,
}


def load_document(path: str) -> tuple[list[Block], str]:
    """Les en fil med riktig leser. Kaster DocumentError med en forklaring
    hvis filen mangler eller filtypen ikke støttes."""
    file = Path(path)
    if not file.is_file():
        raise DocumentError(f"fant ingen fil på '{path}'.")
    reader = READERS.get(file.suffix.lower())
    if reader is None:
        supported = ", ".join(READERS)
        raise DocumentError(
            f"filtypen til '{path}' støttes ikke ({file.suffix or 'ingen filendelse'}). "
            f"Støttede filtyper: {supported}."
        )
    return reader(file)
