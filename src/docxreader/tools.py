"""Verktøy (tools) som en LangChain-agent kan kalle.

Et verktøy er en vanlig Python-funksjon pyntet med `@tool`. Dekoratoren gjør
funksjonen om til et `BaseTool`-objekt som LangChain kan gi til modellen:

- funksjonsnavnet blir verktøyets navn (`read_docx`),
- docstringen blir beskrivelsen modellen leser for å avgjøre NÅR verktøyet
  skal brukes,
- typehintene (`path: str`) blir et skjema som forteller modellen HVILKE
  argumenter den må sende med.

Modellen kjører aldri koden selv. Den ber om et kall ("kall read_docx med
path='x.docx'"), agenten kjører funksjonen, og returverdien sendes tilbake til
modellen som tekst.
"""

from pathlib import Path

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from docx.table import Table
from docx.text.paragraph import Paragraph
from langchain_core.tools import tool


# --- .docx ---------------------------------------------------------------
#
# En .docx har ingen sider; det bestemmes først når Word viser dokumentet.
# Strukturen ligger i stilen til hvert avsnitt ("Heading 1", "List Bullet" …).
# Vi gjør dokumentet om til markdown, som Claude forstår godt.
#
# Funksjonene med _ foran er vanlige hjelpefunksjoner. Bare funksjoner med
# @tool blir verktøy som modellen kan se og kalle.


def _paragraph_to_markdown(paragraph: Paragraph, list_counter: list[int]) -> str:
    """Gjør ett avsnitt om til én linje markdown, basert på stilen.

    `list_counter` er en liste med ett tall som husker hvor langt vi er kommet
    i en nummerert liste. (En liste, fordi et tall ikke kan endres "inne i"
    funksjonen og bli husket utenfor.)
    """
    text = paragraph.text.strip()
    style = paragraph.style.name if paragraph.style is not None else ""

    # Alt som ikke er et nummerert listepunkt, avslutter en nummerert liste.
    if not style.startswith("List Number"):
        list_counter[0] = 0

    if not text:
        return ""
    if style == "Title":
        return f"# {text}"
    if style.startswith("Heading "):
        # "Heading 1" -> "##", "Heading 2" -> "###" … (# er reservert til tittelen).
        # Word lagrer innebygde stilnavn på engelsk, også i norske dokumenter.
        level = style.removeprefix("Heading ")
        if level.isdigit():
            return "#" * (int(level) + 1) + " " + text
    if style.startswith("List Bullet"):
        return f"- {text}"
    if style.startswith("List Number"):
        list_counter[0] += 1
        return f"{list_counter[0]}. {text}"
    return text


def _table_to_markdown(table: Table) -> str:
    """Gjør en Word-tabell om til en markdown-tabell. Første rad blir overskrift."""

    def clean(cell_text: str) -> str:
        # Linjeskift og | ville ødelagt markdown-tabellen.
        return cell_text.replace("\n", " ").replace("|", "\\|").strip()

    rows = [[clean(cell.text) for cell in row.cells] for row in table.rows]
    if not rows:
        return ""

    lines = ["| " + " | ".join(rows[0]) + " |"]
    lines.append("|" + " --- |" * len(rows[0]))
    for row in rows[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


# parse_docstring=True: LangChain leser "Args:"-seksjonen i docstringen og
# legger beskrivelsen av hvert argument inn i skjemaet modellen ser.
@tool(parse_docstring=True)
def read_docx(path: str) -> str:
    """Les et Word-dokument (.docx) og returner innholdet som markdown.

    Overskrifter blir #-overskrifter, lister blir markdown-lister og tabeller
    blir markdown-tabeller. Topp- og bunntekst tas med øverst.
    Bruk dette verktøyet når brukeren spør om innholdet i en .docx-fil.

    Args:
        path: Filsti til .docx-filen, for eksempel "samples/prosjektplan.docx".
    """
    docx_path = Path(path)

    if not docx_path.is_file():
        return f"Feil: fant ingen fil på '{path}'."
    if docx_path.suffix.lower() != ".docx":
        return f"Feil: '{path}' er ikke en .docx-fil."

    try:
        doc = Document(str(docx_path))
    except PackageNotFoundError:
        # python-docx sier "Package not found" også når filen finnes men er
        # ødelagt. Vi gir modellen en tydeligere melding.
        return f"Feil: '{path}' er ikke et gyldig Word-dokument (ødelagt eller feil format)."

    blocks = []

    # Topp- og bunntekst ligger i "sections", utenfor brødteksten.
    # set() fjerner duplikater når flere seksjoner har samme tekst.
    headers = {p.text.strip() for s in doc.sections for p in s.header.paragraphs if p.text.strip()}
    footers = {p.text.strip() for s in doc.sections for p in s.footer.paragraphs if p.text.strip()}
    if headers:
        blocks.append("Topptekst: " + " / ".join(sorted(headers)))
    if footers:
        blocks.append("Bunntekst: " + " / ".join(sorted(footers)))

    # iter_inner_content() gir avsnitt OG tabeller i den rekkefølgen de står i
    # dokumentet. (doc.paragraphs og doc.tables er separate lister og mister
    # rekkefølgen mellom dem.)
    list_counter = [0]
    prev_was_list = False
    for item in doc.iter_inner_content():
        if isinstance(item, Table):
            list_counter[0] = 0
            block, is_list = _table_to_markdown(item), False
        else:
            block = _paragraph_to_markdown(item, list_counter)
            is_list = item.style is not None and item.style.name.startswith("List")
        if not block:
            continue
        # Listepunkter som følger hverandre samles i én blokk (én per linje).
        if is_list and prev_was_list:
            blocks[-1] += "\n" + block
        else:
            blocks.append(block)
        prev_was_list = is_list

    if not blocks:
        return f"Dokumentet '{path}' inneholder ingen tekst."

    # Tom linje mellom blokker gir gyldig og lettlest markdown.
    return "\n\n".join(blocks)
