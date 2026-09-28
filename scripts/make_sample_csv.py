"""Lager samples/reiseregning.csv – en CSV-fil slik norsk Excel lagrer den.

Kjør:  uv run python scripts/make_sample_csv.py

Inneholder fallgruvene CSV-leseren må håndtere:
- tegnkoding cp1252 (Windows), ikke UTF-8 – "æøå" må leses riktig
- semikolon som skilletegn, fordi komma er desimaltegn
- norske tall: "1 250,50" (mellomrom som tusenskille, komma som desimaltegn)
- en celle med semikolon inni, i anførselstegn: "Hansen; Ola (konsulent)"
- en tom linje midt i filen
Plantet faktum: den dyreste reisen er Tromsø-turen til Per Æsøy (8 430,00 kr).
"""

from pathlib import Path

OUT = Path(__file__).parent.parent / "samples" / "reiseregning.csv"

LINES = [
    "Dato;Ansatt;Reisemål;Beløp;Kategori",
    "2026-09-02;Kari Nordmann;Bergen;1 250,50;Fly",
    "2026-09-03;Kari Nordmann;Bergen;890,00;Hotell",
    '2026-09-05;"Hansen; Ola (konsulent)";Oslo;450,00;Tog',
    "2026-09-10;Per Æsøy;Tromsø;8 430,00;Fly",
    "",
    "2026-09-11;Per Æsøy;Tromsø;2 100,00;Hotell",
    "2026-09-15;Kari Nordmann;Trondheim;1 980,00;Fly",
    "2026-09-18;Lise Øvrebø;Stavanger;;Tog",
]

OUT.parent.mkdir(exist_ok=True)
OUT.write_bytes(("\r\n".join(LINES) + "\r\n").encode("cp1252"))
print(f"Lagret {OUT}")
