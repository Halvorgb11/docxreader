"""Lager samples/budsjett.xlsx – en test-arbeidsbok.

Kjør:  uv run python scripts/make_sample_xlsx.py

Inneholder det Excel-leseren må håndtere:
- "Oversikt": liten tabell med dato, desimaltall og en formel som aldri er
  beregnet (filen er ikke åpnet i Excel) + et plantet faktum i en CELLEKOMMENTAR:
  forbruket i IT skyldes engangskjøp av servere fra Nordic Data AS.
  Radene skrives nederst-først for å vise at lagringsrekkefølgen ikke spiller noen rolle.
- "Transaksjoner": 250 rader (deles i biter på 100). Plantet faktum langt nede:
  rad 180 er en betaling på 1 250 000 kr til Fjellsikring AS.
- "Tomt": tomt ark.
- "Hjelpetall": skjult ark.
- "Kvartal": tittel over arket, to tabeller under hverandre med hver sin tittel,
  tabell som starter i kolonne B, tom rad midt i en tabell, tall lagret som
  tekst i engelsk/norsk/valuta-format, og en fotnote nederst.
"""

import datetime
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment

OUT = Path(__file__).parent.parent / "samples" / "budsjett.xlsx"
rng = random.Random(7)

wb = Workbook()

# --- Oversikt ---------------------------------------------------------------
ws = wb.active
ws.title = "Oversikt"
rows = [
    ("Avdeling", "Budsjett", "Forbruk", "Rest", "Sist oppdatert"),
    ("Salg", 1200000, 950000.0, "=B2-C2", datetime.datetime(2026, 9, 1)),
    ("IT", 800000, 1150000.5, "=B3-C3", datetime.datetime(2026, 9, 15)),
    ("HR", 400000, 310000, "=B4-C4", datetime.datetime(2026, 8, 30)),
]
for number in reversed(range(len(rows))):  # nederst først
    for col, value in enumerate(rows[number], start=1):
        ws.cell(row=number + 1, column=col, value=value)
ws["C3"].comment = Comment(
    "Overforbruket skyldes engangskjøp av servere fra Nordic Data AS i august.", "Kari Nordmann"
)

# --- Transaksjoner (stort ark) ------------------------------------------------
ws = wb.create_sheet("Transaksjoner")
ws.append(["Dato", "Leverandør", "Beløp (kr)", "Avdeling"])
suppliers = ["Kontorland AS", "Rema 1000", "Telenor", "Posten", "Elkjøp", "Staples"]
for n in range(250):
    row = ws.max_row + 1
    if row == 180:
        ws.append([datetime.datetime(2026, 6, 12), "Fjellsikring AS", 1250000, "IT"])
    else:
        date = datetime.datetime(2026, 1, 1) + datetime.timedelta(days=n)
        ws.append([date, rng.choice(suppliers), rng.randint(200, 20000), rng.choice(["Salg", "IT", "HR"])])

# --- Tomt og skjult ----------------------------------------------------------
wb.create_sheet("Tomt")
ws = wb.create_sheet("Hjelpetall")
ws.append(["Nøkkel", "Verdi"])
ws.append(["Momssats", 0.25])
ws.sheet_state = "hidden"

# --- Kvartal: tittel, to tabeller, tom rad i tabell, fotnote, tall som tekst ---
ws = wb.create_sheet("Kvartal")
ws["A1"] = "Kvartalsrapport 2026"
ws["B3"] = "Budsjett Q1"
for row, (avdeling, belop) in zip([4, 5, 6, 8], [("Avdeling", "Q1 budsjett"), ("Salg", 300000), ("IT", 200000), ("HR", 100000)]):
    ws.cell(row=row, column=2, value=avdeling)
    ws.cell(row=row, column=3, value=belop)
ws["B10"] = "Faktisk Q1"
for row, values in zip(range(11, 15), [
    ("Avdeling", "Q1 faktisk", "Kommentar"),
    ("Salg", "1,250,000.50", "Over plan"),  # engelsk format, lagret som tekst
    ("IT", "1.180.000,00", "Serverkjøp"),  # norsk/tysk format, lagret som tekst
    ("HR", "kr 95 000,-", None),
]):
    for col, value in enumerate(values, start=2):
        ws.cell(row=row, column=col, value=value)
ws["A16"] = "Kilde: regnskapssystemet, hentet 2026-09-28"

OUT.parent.mkdir(exist_ok=True)
wb.save(OUT)
print(f"Lagret {OUT}")
