"""Lager samples/arsrapport.docx – et STORT test-dokument (~10 000 ord).

Kjør:  uv run python scripts/make_large_sample_docx.py

Brukes til å teste hvordan verktøyene håndterer store filer:
- 12 kapitler (Heading 1), hvert med de SAMME underoverskriftene
  ("Status", "Nøkkeltall", "Utfordringer") – tester stier som "Økonomi > Status".
- Fylltekst er generert med fast tilfeldighetsfrø, så filen blir lik hver gang.
- Tre "plantede" fakta som agenten skal kunne finne:
    * Region Nord > Status:          kontaktperson er Ingrid Solberg
    * IT og sikkerhet > Utfordringer: 3 alvorlige sikkerhetshendelser
    * Økonomi > Nøkkeltall (tabell): driftsresultat 38,5 MNOK
"""

import random
from pathlib import Path

from docx import Document

OUT = Path(__file__).parent.parent / "samples" / "arsrapport.docx"

CHAPTERS = [
    "Sammendrag", "Økonomi", "Salg og marked", "Kundeservice",
    "Drift og infrastruktur", "IT og sikkerhet", "HR og organisasjon",
    "Bærekraft", "Region Nord", "Region Sør", "Risiko", "Planer for 2027",
]
SUBSECTIONS = ["Status", "Nøkkeltall", "Utfordringer"]

PLANTED = {
    ("Region Nord", "Status"): "Kontaktperson for Region Nord er Ingrid Solberg.",
    ("IT og sikkerhet", "Utfordringer"): "Det ble registrert 3 alvorlige sikkerhetshendelser i 2026.",
}

# Byggeklosser for fylltekst.
SUBJECTS = ["Avdelingen", "Teamet", "Selskapet", "Ledelsen", "Prosjektgruppen", "Organisasjonen"]
VERBS = ["har forbedret", "har videreutviklet", "har evaluert", "har styrket", "har gjennomgått", "har prioritert"]
OBJECTS = [
    "rutinene for oppfølging", "samarbeidet på tvers av enheter", "kvaliteten i leveransene",
    "rapporteringen til styret", "kompetansen blant de ansatte", "dialogen med kundene",
    "arbeidet med kontinuerlig forbedring", "styringen av pågående prosjekter",
]
ENDINGS = [
    "gjennom året", "i tråd med strategien", "med gode resultater",
    "sammenlignet med fjoråret", "etter innspill fra ansatte", "som planlagt",
]

rng = random.Random(42)


def filler_paragraph(sentences: int = 7) -> str:
    return " ".join(
        f"{rng.choice(SUBJECTS)} {rng.choice(VERBS)} {rng.choice(OBJECTS)} {rng.choice(ENDINGS)}."
        for _ in range(sentences)
    )


doc = Document()
doc.sections[0].header.paragraphs[0].text = "Fjellbekk AS – Årsrapport"
doc.add_heading("Årsrapport 2026 – Fjellbekk AS", level=0)

for chapter in CHAPTERS:
    doc.add_heading(chapter, level=1)
    doc.add_paragraph(filler_paragraph(4))
    for sub in SUBSECTIONS:
        doc.add_heading(sub, level=2)
        for n in range(3):
            text = filler_paragraph()
            # Plant faktaet midt i andre avsnitt, så det ikke ligger i første setning.
            if n == 1 and (chapter, sub) in PLANTED:
                text = text + " " + PLANTED[(chapter, sub)] + " " + filler_paragraph(3)
            doc.add_paragraph(text)
        if sub == "Nøkkeltall":
            table = doc.add_table(rows=1, cols=3)
            table.style = "Table Grid"
            table.rows[0].cells[0].text = "Nøkkeltall"
            table.rows[0].cells[1].text = "2025"
            table.rows[0].cells[2].text = "2026"
            rows = [("Omsetning", "MNOK"), ("Kostnader", "MNOK"), ("Ansatte", "")]
            if chapter == "Økonomi":
                rows.insert(2, ("Driftsresultat", "MNOK"))
            for name, unit in rows:
                cells = table.add_row().cells
                cells[0].text = name
                if name == "Driftsresultat":
                    cells[1].text, cells[2].text = "31,2 MNOK", "38,5 MNOK"
                else:
                    a, b = rng.randint(10, 400), rng.randint(10, 400)
                    cells[1].text = f"{a} {unit}".strip()
                    cells[2].text = f"{b} {unit}".strip()

OUT.parent.mkdir(exist_ok=True)
doc.save(OUT)
print(f"Lagret {OUT}")
