"""Felles byggeklosser for alle filtyper.

Hver leser i `readers/` gjør en fil om til en liste med `Block`-er. Alt her
(innholdsfortegnelse, seksjoner, søk) jobber bare med blokkene og vet ikke
hvilken filtype de kom fra. Derfor får en ny filtype disse funksjonene gratis.
"""

import re
from dataclasses import dataclass


@dataclass
class Block:
    """Én del av dokumentet: en overskrift, et avsnitt, en liste eller en tabell."""

    markdown: str
    words: int
    # 0 = ikke en overskrift. 1 = #, 2 = ##, 3 = ### …
    # (Word: 1 = tittel, 2 = Heading 1. PowerPoint: 2 = lysbilde.)
    heading_level: int = 0
    is_table: bool = False

    @property
    def heading_text(self) -> str:
        return self.markdown.lstrip("#").strip() if self.heading_level else ""


class DocumentError(Exception):
    """En feil som skal vises til modellen som en "Feil: …"-tekst."""


def count_words(text: str) -> int:
    """Antall ord. Markdown-tegn som "|", "---" og "#" telles ikke."""
    return sum(1 for w in text.split() if any(c.isalnum() for c in w))


def heading(level: int, text: str) -> Block:
    """Lag en overskriftsblokk, f.eks. heading(2, "Budsjett") -> "## Budsjett"."""
    return Block("#" * level + " " + text, count_words(text), heading_level=level)


def rows_to_markdown(rows: list[list[str]]) -> str:
    """Tabell (liste av rader) -> markdown-tabell. Første rad blir overskrift."""

    def clean(cell_text: str) -> str:
        # Linjeskift og | ville ødelagt markdown-tabellen.
        return cell_text.replace("\n", " ").replace("|", "\\|").strip()

    rows = [[clean(cell) for cell in row] for row in rows]
    if not rows:
        return ""
    lines = ["| " + " | ".join(rows[0]) + " |"]
    lines.append("|" + " --- |" * len(rows[0]))
    for row in rows[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def table_block(rows: list[list[str]]) -> Block:
    return Block(rows_to_markdown(rows), sum(count_words(c) for r in rows for c in r), is_table=True)


def render(blocks: list[Block]) -> str:
    """Blokker -> markdown. Tom linje mellom blokker gir gyldig og lettlest markdown."""
    return "\n\n".join(b.markdown for b in blocks)


def section_end(blocks: list[Block], start: int) -> int:
    """En seksjon er overskriften på `start` og alt etter den, fram til neste
    overskrift på samme eller høyere nivå. Returnerer indeksen der den slutter."""
    level = blocks[start].heading_level
    for i in range(start + 1, len(blocks)):
        if 0 < blocks[i].heading_level <= level:
            return i
    return len(blocks)


def outline(blocks: list[Block], start: int = 0, end: int | None = None) -> str:
    """Innholdsfortegnelse: hver overskrift med størrelsen på seksjonen sin."""
    end = len(blocks) if end is None else end
    lines = []

    # Tekst før første overskrift ville ellers vært usynlig i oversikten.
    first_heading = next((i for i in range(start, end) if blocks[i].heading_level), end)
    if intro_words := sum(b.words for b in blocks[start:first_heading]):
        lines.append(f"(tekst før første overskrift: ~{intro_words} ord)")

    for i in range(first_heading, end):
        if not blocks[i].heading_level:
            continue
        section = blocks[i : section_end(blocks, i)]
        words = sum(b.words for b in section)
        tables = sum(b.is_table for b in section)
        size = f"~{words} ord"
        if tables:
            size += f", {tables} tabell" + ("er" if tables > 1 else "")
        lines.append(f"{blocks[i].markdown}  ({size})")
    return "\n".join(lines)


def heading_paths(blocks: list[Block]) -> dict[int, list[str]]:
    """For hver overskrift: stien av overskrifter ned til den.

    F.eks. {7: ["Økonomi", "Status"]}. Finnes det bare ÉN nivå 1-overskrift,
    regnes den som dokumentets tittel og tas ikke med i stiene til de andre
    overskriftene. (Markdown-filer har ofte flere "#"-kapitler; da er de med.)
    """
    single_title = sum(b.heading_level == 1 for b in blocks) == 1
    stack: list[tuple[int, str]] = []
    paths = {}
    for i, b in enumerate(blocks):
        if not b.heading_level:
            continue
        if b.heading_level == 1 and single_title:
            paths[i] = [b.heading_text]
            continue
        while stack and stack[-1][0] >= b.heading_level:
            stack.pop()
        stack.append((b.heading_level, b.heading_text))
        paths[i] = [text for _, text in stack]
    return paths


def find_section(blocks: list[Block], wanted_heading: str) -> int:
    """Finn indeksen til overskriften modellen ba om.

    `wanted_heading` kan være en enkelt overskrift ("Budsjett") eller en sti
    ("Økonomi > Status"). Store/små bokstaver spiller ingen rolle.
    Først prøves eksakt treff, deretter treff på deler av teksten.
    """
    wanted = [part.strip().casefold() for part in wanted_heading.split(">") if part.strip()]
    if not wanted:
        raise DocumentError("overskriften kan ikke være tom.")
    paths = heading_paths(blocks)

    def matches(path: list[str], exact: bool) -> bool:
        path = [p.casefold() for p in path]

        def same(w: str, p: str) -> bool:
            return w == p if exact else w in p

        if not same(wanted[-1], path[-1]):
            return False
        # De øvrige delene må finnes blant foreldrene, i riktig rekkefølge.
        # ("Transaksjoner > Rad 2–101" finner "Ark: Transaksjoner > Rad 2–101".)
        parents = iter(path[:-1])
        return all(any(same(w, p) for p in parents) for w in wanted[:-1])

    for exact in (True, False):
        hits = [i for i, path in paths.items() if matches(path, exact)]
        if len(hits) == 1:
            return hits[0]
        if len(hits) > 1:
            options = "\n".join("- " + " > ".join(paths[i]) for i in hits)
            raise DocumentError(
                f"flere overskrifter passer '{wanted_heading}':\n{options}\n"
                "Oppgi hele overskriften eller stien, f.eks. 'Kapittel > Underoverskrift'."
            )

    available = "\n".join("- " + " > ".join(p) for p in paths.values())
    raise DocumentError(
        f"fant ingen overskrift som passer '{wanted_heading}'. Tilgjengelige:\n{available}"
    )


def block_paths(blocks: list[Block]) -> list[str]:
    """For hver blokk: overskriftsstien den står under, f.eks. "Økonomi > Status"."""
    paths = heading_paths(blocks)
    current = "(før første overskrift)"
    result = []
    for i in range(len(blocks)):
        if i in paths:
            current = " > ".join(paths[i])
        result.append(current)
    return result


def search_units(block: Block) -> tuple[str, list[str]]:
    """Del en blokk i biter vi søker i, så et treff ikke returnerer et helt avsnitt.

    Returnerer (tekst som vises foran treffene, liste med biter):
    - tabell:  hver rad; tabellens overskriftsrad vises foran, så tallene gir mening
    - liste:   hvert listepunkt (eller linje)
    - avsnitt: hver setning
    """
    if block.is_table:
        lines = block.markdown.splitlines()
        return "\n".join(lines[:2]), lines[2:]
    if "\n" in block.markdown:
        return "", block.markdown.splitlines()
    # Del etter punktum/utrop/spørsmålstegn etterfulgt av mellomrom.
    return "", re.split(r"(?<=[.!?])\s+", block.markdown)
