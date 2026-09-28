"""Leser for <FORMAT> (.<endelse>), med <bibliotek>.

<Kort forklaring av hvordan formatet er bygd opp, og hva som blir en seksjon.>
"""

from pathlib import Path

# from <bibliotek> import <Klasse>
# from <bibliotek>.<modul> import <UnntakForØdelagtFil>

from docxreader.blocks import Block, DocumentError, count_words, heading, table_block


def load(path: Path) -> tuple[list[Block], str]:
    """Åpne filen og del den i blokker. Returnerer (blokker, metatekst)."""
    try:
        document = ...  # <Klasse>(str(path))
    except Exception:  # bytt til bibliotekets eget unntak for ødelagte filer
        raise DocumentError(f"'{path}' er ikke en gyldig <FORMAT>-fil (ødelagt eller feil format).")

    blocks: list[Block] = []
    for part in ...:  # i LESE-rekkefølge
        # Seksjon:  blocks.append(heading(2, "Tittel"))
        # Avsnitt:  blocks.append(Block(text, count_words(text)))
        # Liste:    text = "\n".join(f"- {p}" for p in punkter); blocks.append(Block(text, count_words(text)))
        # Tabell:   blocks.append(table_block([["Kolonne", …], ["verdi", …]]))
        pass

    meta = ""  # f.eks. tittel, forfatter, antall sider/ark/lysbilder
    return blocks, meta
