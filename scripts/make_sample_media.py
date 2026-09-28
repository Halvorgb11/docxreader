"""Lager bilde- og PDF-testfiler med Pillow (ingen tekstlag – modellen må SE dem).

Kjør:  uv run python scripts/make_sample_media.py

- samples/kvittering.png: en kvittering. Plantet faktum: totalt 1 487,50 kr,
  betalt med kort 12. september 2026.
- samples/moteinnkalling.pdf: to sider. Plantet faktum på side 2: styremøtet
  holdes i Ålesund 14. november 2026.
- skisse_png(): brukes av make_sample_eml.py som bildevedlegg. Plantet faktum:
  serverne skal stå i rack B3.
"""

import io
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

SAMPLES = Path(__file__).parent.parent / "samples"


def _page(lines: list[str], size=(800, 600), font_size=30) -> Image.Image:
    image = Image.new("RGB", size, "white")
    draw = ImageDraw.Draw(image)
    font = ImageFont.load_default(size=font_size)
    y = 40
    for line in lines:
        draw.text((40, y), line, fill="black", font=font)
        y += int(font_size * 1.6)
    return image


def skisse_png() -> bytes:
    image = _page(["Serverrom – skisse", "", "Rack A1: nettverk", "Rack B3: 4 nye servere (Nordic Data)",
                   "Rack C2: backup"], size=(700, 360))
    ImageDraw.Draw(image).rectangle((20, 20, 680, 340), outline="black", width=3)
    buffer = io.BytesIO()
    image.save(buffer, "PNG")
    return buffer.getvalue()


if __name__ == "__main__":
    SAMPLES.mkdir(exist_ok=True)
    _page(["Butikken Sentrum", "Kvittering nr. 88213", "", "Kaffe            89,90", "Brød             42,00",
           "Kontorrekvisita 1 355,60", "", "TOTALT     1 487,50 kr", "Betalt med kort 12.09.2026"],
          size=(600, 700)).save(SAMPLES / "kvittering.png")

    page1 = _page(["Innkalling til styremøte", "", "Saksliste:", "1. Godkjenning av protokoll",
                   "2. Budsjett 2027", "3. Serverinvestering"])
    page2 = _page(["Praktisk informasjon", "", "Sted: Ålesund, hotell Brosundet", "Dato: 14. november 2026",
                   "Tid: kl. 09.00–15.00"])
    page1.save(SAMPLES / "moteinnkalling.pdf", save_all=True, append_images=[page2])
    print("Lagret kvittering.png og moteinnkalling.pdf")
