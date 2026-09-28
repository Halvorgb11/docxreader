"""Leser for Markdown (.md) og ren tekst (.txt). Trenger ingen ekstra bibliotek.

Markdown er allerede formatet verktøyene returnerer, så vi trenger bare å
dele filen i blokker: overskrifter, avsnitt, lister, tabeller og kodeblokker.
"""

import re
from pathlib import Path

from docxreader.blocks import Block, DocumentError, count_words, heading

HEADING = re.compile(r"^(#{1,6})\s+(.*?)\s*#*\s*$")  # "## Tittel" (og "## Tittel ##")
LIST_ITEM = re.compile(r"^\s*([-*+]|\d+[.)])\s+")  # "- punkt", "* punkt", "1. punkt"


def read_text(path: Path) -> str:
    """Les filen som tekst. Prøv UTF-8 først, deretter Windows-koding (cp1252),
    som er vanlig i eldre norske filer. "utf-8-sig" tåler også BOM-tegnet
    som enkelte Windows-programmer legger først i filen."""
    raw = path.read_bytes()
    if b"\x00" in raw[:1024]:
        raise DocumentError(f"'{path}' ser ut til å være en binærfil, ikke tekst.")
    try:
        return raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        return raw.decode("cp1252", errors="replace")


def _chunks(text: str, split_on_headings: bool = False) -> list[list[str]]:
    """Del teksten i biter adskilt av tomme linjer. Kodeblokker (```) holdes
    samlet selv om de har tomme linjer inni seg.

    Med split_on_headings starter også hver "#"-overskrift en ny bit, siden
    Markdown tillater overskrifter rett under tekst uten tom linje mellom.
    """
    chunks: list[list[str]] = []
    current: list[str] = []
    in_code = False
    for line in text.splitlines():
        if line.strip().startswith("```"):
            in_code = not in_code
        is_heading = split_on_headings and not in_code and HEADING.match(line.strip())
        if (not line.strip() and not in_code) or is_heading:
            if current:
                chunks.append(current)
                current = []
            if is_heading:
                chunks.append([line])
            continue
        current.append(line.rstrip())
    if current:
        chunks.append(current)
    return chunks


def _markdown_chunk_to_block(lines: list[str]) -> Block:
    """Én bit markdown -> en blokk."""
    first = lines[0].strip()
    if first.startswith("```"):
        code = "\n".join(lines)
        return Block(code, count_words(code))
    if all(line.strip().startswith("|") for line in lines):
        table = "\n".join(line.strip() for line in lines)
        return Block(table, count_words(table), is_table=True)

    if m := HEADING.match(first):  # _chunks legger overskrifter i egne biter
        return heading(len(m.group(1)), m.group(2))
    if LIST_ITEM.match(lines[0]):
        # Lister beholder én linje per punkt (søket deler på linjeskift).
        text = "\n".join(lines)
        return Block(text, count_words(text))
    # Et avsnitt kan være brutt over flere linjer i filen; slå dem sammen.
    paragraph = " ".join(line.strip() for line in lines)
    return Block(paragraph, count_words(paragraph))


def markdown_blocks(text: str) -> list[Block]:
    """Markdown-tekst -> blokker. Brukes også for e-posttekst."""
    return [_markdown_chunk_to_block(chunk) for chunk in _chunks(text, split_on_headings=True)]


def load_markdown(path: Path) -> tuple[list[Block], str]:
    return markdown_blocks(read_text(path)), ""


def load_plain(path: Path) -> tuple[list[Block], str]:
    """Ren tekst har ingen overskrifter. Hvert avsnitt (adskilt av tom linje)
    blir én blokk. Søk og lesing fungerer; innholdsfortegnelse gir lite."""
    blocks = []
    for chunk in _chunks(read_text(path)):
        paragraph = " ".join(line.strip() for line in chunk)
        blocks.append(Block(paragraph, count_words(paragraph)))
    return blocks, ""
