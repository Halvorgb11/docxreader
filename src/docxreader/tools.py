"""Verktøy (tools) som en LangChain-agent kan kalle.

Et verktøy er en vanlig Python-funksjon pyntet med `@tool`. Dekoratoren gjør
funksjonen om til et `BaseTool`-objekt som LangChain kan gi til modellen:

- funksjonsnavnet blir verktøyets navn (`read_pdf`),
- docstringen blir beskrivelsen modellen leser for å avgjøre NÅR verktøyet
  skal brukes,
- typehintene (`path: str`) blir et skjema som forteller modellen HVILKE
  argumenter den må sende med.

Modellen kjører aldri koden selv. Den ber om et kall ("kall read_pdf med
path='x.pdf'"), agenten kjører funksjonen, og returverdien sendes tilbake til
modellen som tekst.
"""

from pathlib import Path

from langchain_core.tools import tool
from pypdf import PdfReader
from pypdf.errors import PdfReadError


# parse_docstring=True: LangChain leser "Args:"-seksjonen i docstringen og
# legger beskrivelsen av hvert argument inn i skjemaet modellen ser.
@tool(parse_docstring=True)
def read_pdf(path: str) -> str:
    """Les all tekst fra en PDF-fil og returner den, side for side.

    Bruk dette verktøyet når brukeren spør om innholdet i en PDF-fil.

    Args:
        path: Filsti til PDF-filen, for eksempel "samples/rapport.pdf".
    """
    pdf_path = Path(path)

    # Feil returneres som tekst i stedet for å kaste unntak. Da får modellen
    # vite hva som gikk galt og kan prøve igjen eller forklare det til brukeren.
    if not pdf_path.is_file():
        return f"Feil: fant ingen fil på '{path}'."
    if pdf_path.suffix.lower() != ".pdf":
        return f"Feil: '{path}' er ikke en PDF-fil."

    try:
        reader = PdfReader(pdf_path)
    except PdfReadError as e:
        return f"Feil: klarte ikke å lese '{path}' som PDF ({e})."

    pages = []
    for number, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        # PDF-tekst har ofte mye luft; fjern tomme linjer og etterfølgende mellomrom.
        lines = [line.rstrip() for line in text.splitlines() if line.strip()]
        body = "\n".join(lines) if lines else "(ingen tekst på denne siden)"
        # Sidemarkører lar modellen vise til sidenummer i svaret sitt.
        pages.append(f"--- Side {number} av {len(reader.pages)} ---\n{body}")

    return "\n\n".join(pages)
