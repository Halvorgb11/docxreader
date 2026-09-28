"""Verktøy (tools) som en LangChain-agent kan kalle.

Et verktøy er en vanlig Python-funksjon pyntet med `@tool`. Dekoratoren gjør
funksjonen om til et `BaseTool`-objekt som LangChain kan gi til modellen:

- funksjonsnavnet blir verktøyets navn (`read_document`),
- docstringen blir beskrivelsen modellen leser for å avgjøre NÅR verktøyet
  skal brukes,
- typehintene (`path: str`) blir et skjema som forteller modellen HVILKE
  argumenter den må sende med.

Modellen kjører aldri koden selv. Den ber om et kall ("kall read_document med
path='x.docx'"), agenten kjører funksjonen, og returverdien sendes tilbake til
modellen som tekst.

Verktøyene fungerer for alle filtyper i readers.READERS. Vi har FÅ generelle
verktøy i stedet for ett sett per filtype: da slipper modellen å velge mellom
mange nesten like verktøy, og en ny filtype trenger ingen nye verktøy.

- read_document    – hele dokumentet (eller innholdsfortegnelsen hvis det er stort)
- document_outline – bare innholdsfortegnelsen, med størrelse per seksjon
- read_section     – én seksjon, valgt med overskriften
- search_document  – finn setninger, listepunkter og tabellrader som inneholder søkeord
"""

from langchain_core.tools import tool

from docxreader.blocks import (
    DocumentError,
    block_paths,
    find_section,
    outline,
    render,
    search_units,
    section_end,
)
from docxreader.readers import load_document

# Dokumenter (og seksjoner) over dette antallet ord sendes ikke i sin helhet.
# ~3000 ord er omtrent 4000–5000 tokens.
MAX_WORDS = 3000

# search_document viser høyst så mange treff.
MAX_SEARCH_HITS = 15


def _with_meta(meta: str, text: str) -> str:
    return (meta + "\n\n" if meta else "") + text


# parse_docstring=True: LangChain leser "Args:"-seksjonen i docstringen og
# legger beskrivelsen av hvert argument inn i skjemaet modellen ser.
@tool(parse_docstring=True)
def read_document(path: str) -> str:
    """Les et helt dokument og returner innholdet som markdown.

    Støttede filtyper: Word (.docx), PowerPoint (.pptx), Markdown (.md) og
    tekst (.txt). Overskrifter blir #-overskrifter, lister blir markdown-lister
    og tabeller blir markdown-tabeller. I PowerPoint er hvert lysbilde en
    seksjon ("Lysbilde 3: Tittel"), med talenotater.
    Er dokumentet stort, returneres innholdsfortegnelsen i stedet; bruk da
    read_section for å lese delene du trenger.

    Args:
        path: Filsti til dokumentet, for eksempel "samples/prosjektplan.docx".
    """
    try:
        blocks, meta = load_document(path)
    except DocumentError as e:
        return f"Feil: {e}"
    if not blocks:
        return f"Dokumentet '{path}' inneholder ingen tekst."

    total = sum(b.words for b in blocks)
    if total > MAX_WORDS:
        return (
            f"Dokumentet er stort (~{total} ord), så det returneres ikke i sin helhet.\n"
            "Her er innholdsfortegnelsen. Bruk read_section med en overskrift "
            "herfra for å lese den delen du trenger.\n\n"
            + _with_meta(meta, outline(blocks))
        )
    return _with_meta(meta, render(blocks))


@tool(parse_docstring=True)
def document_outline(path: str) -> str:
    """Vis innholdsfortegnelsen til et dokument, uten selve teksten.

    Støttede filtyper: Word (.docx), PowerPoint (.pptx), Markdown (.md) og
    tekst (.txt). Viser alle overskrifter (i PowerPoint: alle lysbilder) med
    omtrentlig antall ord og tabeller i hver seksjon. Bruk dette først for å
    finne ut hvor svaret står, og les deretter bare den delen med read_section.

    Args:
        path: Filsti til dokumentet, for eksempel "samples/prosjektplan.docx".
    """
    try:
        blocks, meta = load_document(path)
    except DocumentError as e:
        return f"Feil: {e}"

    total = sum(b.words for b in blocks)
    if not any(b.heading_level for b in blocks):
        return f"Dokumentet har ingen overskrifter (~{total} ord totalt). Bruk read_document for å lese det."
    return _with_meta(meta, f"Totalt ~{total} ord.\n\n" + outline(blocks))


@tool(parse_docstring=True)
def read_section(path: str, heading: str) -> str:
    """Les én seksjon av et dokument: en overskrift med alt innhold under den,
    inkludert underoverskrifter. I PowerPoint er en seksjon ett lysbilde.

    Støttede filtyper: Word (.docx), PowerPoint (.pptx), Markdown (.md) og
    tekst (.txt). Finn overskriftene med document_outline først. Hvis samme
    overskrift finnes flere steder, oppgi stien med " > ", f.eks. "Økonomi > Status".

    Args:
        path: Filsti til dokumentet, for eksempel "samples/prosjektplan.docx".
        heading: Overskriften til seksjonen, f.eks. "Budsjett", "Økonomi > Status" eller "Lysbilde 3". Store/små bokstaver spiller ingen rolle.
    """
    try:
        blocks, _ = load_document(path)
        start = find_section(blocks, heading)
    except DocumentError as e:
        return f"Feil: {e}"

    end = section_end(blocks, start)
    section = blocks[start:end]
    words = sum(b.words for b in section)
    has_subsections = any(b.heading_level for b in section[1:])

    # Også en seksjon kan være for stor. Har den underoverskrifter, viser vi
    # dem i stedet, så modellen kan gå et nivå dypere.
    if words > MAX_WORDS and has_subsections:
        return (
            f"Seksjonen '{blocks[start].heading_text}' er stor (~{words} ord). "
            "Her er underoverskriftene; les dem enkeltvis med read_section.\n\n"
            + outline(blocks, start, end)
        )
    return render(section)


@tool(parse_docstring=True)
def search_document(path: str, query: str) -> str:
    """Søk etter ord i et dokument og få bare de treffende setningene,
    listepunktene og tabellradene, med overskriften de står under.

    Støttede filtyper: Word (.docx), PowerPoint (.pptx), Markdown (.md) og
    tekst (.txt). Alle ordene i søket må finnes i samme setning/rad. Store/små
    bokstaver spiller ingen rolle, og deler av ord gir treff ("sikkerhet"
    finner "sikkerhetshendelser"). Bruk dette når du leter etter noe bestemt
    og ikke vet hvor det står. Les mer sammenheng med read_section.

    Args:
        path: Filsti til dokumentet, for eksempel "samples/prosjektplan.docx".
        query: Ett eller flere søkeord, f.eks. "driftsresultat" eller "kontaktperson nord". Korte ord og ordstammer gir flest treff.
    """
    words = query.casefold().split()
    if not words:
        return "Feil: søket kan ikke være tomt."
    try:
        blocks, _ = load_document(path)
    except DocumentError as e:
        return f"Feil: {e}"

    # Samle treff: (blokk, overskriftssti, tekst foran treffene, treffende biter).
    hits = []
    for block, block_path in zip(blocks, block_paths(blocks)):
        prefix, units = search_units(block)
        matched = [u for u in units if all(w in u.casefold() for w in words)]
        if matched:
            hits.append((block, block_path, prefix, matched))

    total = sum(len(matched) for *_, matched in hits)
    if total == 0:
        return (
            f"Ingen treff for '{query}'. Prøv færre eller kortere ord (en ordstamme "
            "finner også sammensatte ord), eller se innholdsfortegnelsen med document_outline."
        )

    lines = [f"{total} treff for '{query}':"]
    shown = 0
    previous_path = None
    for block, block_path, prefix, matched in hits:
        if shown >= MAX_SEARCH_HITS:
            break
        matched = matched[: MAX_SEARCH_HITS - shown]
        shown += len(matched)
        # Treff under samme overskrift grupperes under én [sti]-linje.
        if block_path != previous_path:
            lines.append(f"\n[{block_path}]")
            previous_path = block_path
        if prefix:  # tabell: overskriftsraden først, så tallene gir mening
            lines.append(prefix)
        # Setninger fra vanlige avsnitt får "- " foran. Tabellrader, listepunkter
        # og overskrifter har allerede sin egen markdown-form.
        is_plain_paragraph = not (block.is_table or block.heading_level or "\n" in block.markdown)
        lines.extend(f"- {u}" if is_plain_paragraph else u for u in matched)

    if shown < total:
        lines.append(
            f"\nViser {shown} av {total} treff. Gjør søket mer presist (flere ord), "
            "eller les en seksjon med read_section."
        )
    return "\n".join(lines)


# Alle verktøyene samlet, så agent.py (og tester) kan hente dem på ett sted.
ALL_TOOLS = [read_document, document_outline, read_section, search_document]
