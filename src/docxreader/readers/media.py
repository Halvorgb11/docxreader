"""PDF og bilder: gjøres ikke om til tekst, men sendes rett til Claude.

Claude leser PDF og ser bilder selv. Vi lager bare en LangChain-innholdsblokk
("content block") som verktøyet view_file returnerer:

    {"type": "image", "base64": "...", "mime_type": "image/png"}
    {"type": "file",  "base64": "...", "mime_type": "application/pdf"}

langchain-anthropic gjør dem om til Claudes bilde- og dokumentblokker.

Grenser vi passer på (ellers feiler API-kallet, og hele agenten stopper):
- Bilder: Claude tar PNG, JPEG, GIF og WebP. TIFF/BMP gjøres om til PNG.
  Store bilder skaleres ned til MAX_IMAGE_SIDE (Claude skalerer uansett ned,
  så det sparer bare tokens og båndbredde).
- PDF: maks MAX_PDF_BYTES og MAX_PDF_PAGES.
"""

import base64
import io
import re

from PIL import Image, UnidentifiedImageError

from docxreader.blocks import DocumentError

MIME_TYPES = {
    ".pdf": "application/pdf",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".webp": "image/webp",
    ".tif": "image/tiff",  # gjøres om til PNG
    ".tiff": "image/tiff",
    ".bmp": "image/bmp",
}
CLAUDE_IMAGE_FORMATS = {"PNG": "image/png", "JPEG": "image/jpeg", "GIF": "image/gif", "WEBP": "image/webp"}

MAX_IMAGE_SIDE = 1568
MAX_IMAGE_BYTES = 3_500_000  # API-grensen er 5 MB etter base64 (som gjør filen ~33 % større)
MAX_PDF_BYTES = 30_000_000
MAX_PDF_PAGES = 100


def _encode(data: bytes) -> str:
    return base64.b64encode(data).decode("ascii")


def _image_block(name: str, data: bytes) -> tuple[dict, str]:
    try:
        image = Image.open(io.BytesIO(data))
        image.load()
    except (UnidentifiedImageError, OSError):
        raise DocumentError(f"'{name}' er ikke et gyldig bilde (ødelagt eller feil format).")

    note = ""
    frames = getattr(image, "n_frames", 1)
    if frames > 1 and image.format != "GIF":
        note = f" Viser side 1 av {frames}."  # f.eks. TIFF med flere sider

    ok_as_is = (
        image.format in CLAUDE_IMAGE_FORMATS
        and max(image.size) <= MAX_IMAGE_SIDE
        and len(data) <= MAX_IMAGE_BYTES
    )
    if ok_as_is:
        return {"type": "image", "base64": _encode(data), "mime_type": CLAUDE_IMAGE_FORMATS[image.format]}, note

    # Gjør om (og skaler ned): TIFF/BMP, eller for stort bilde.
    original = image.size
    image = image.convert("RGBA" if "A" in image.getbands() or image.mode == "P" else "RGB")
    image.thumbnail((MAX_IMAGE_SIDE, MAX_IMAGE_SIDE))  # beholder forholdet mellom bredde og høyde
    buffer = io.BytesIO()
    if image.mode == "RGB":
        image.save(buffer, "JPEG", quality=85)  # bilder uten gjennomsiktighet: JPEG er mye mindre
        mime = "image/jpeg"
    else:
        image.save(buffer, "PNG")
        mime = "image/png"
    if image.size != original:
        note += f" Skalert ned fra {original[0]}×{original[1]} til {image.size[0]}×{image.size[1]} piksler."
    return {"type": "image", "base64": _encode(buffer.getvalue()), "mime_type": mime}, note


def _pdf_block(name: str, data: bytes) -> tuple[dict, str]:
    if not data.startswith(b"%PDF-"):
        raise DocumentError(f"'{name}' er ikke en gyldig PDF-fil.")
    if len(data) > MAX_PDF_BYTES:
        raise DocumentError(f"'{name}' er for stor ({len(data) // 1_000_000} MB; maks {MAX_PDF_BYTES // 1_000_000} MB).")
    # Omtrentlig sideantall: tell "/Type /Page" (men ikke "/Pages"). Komprimerte
    # PDF-er kan skjule sidene; da gir tellingen 0 og vi slipper filen gjennom.
    pages = len(re.findall(rb"/Type\s*/Page(?![s\w])", data))
    if pages > MAX_PDF_PAGES:
        raise DocumentError(f"'{name}' har {pages} sider; Claude kan lese maks {MAX_PDF_PAGES} sider om gangen.")
    note = f" {pages} sider." if pages else ""
    return {"type": "file", "base64": _encode(data), "mime_type": "application/pdf"}, note


def media_block(name: str, data: bytes) -> tuple[dict, str]:
    """Fil -> (LangChain-innholdsblokk, merknad om hva som ble gjort)."""
    suffix = "." + name.rsplit(".", 1)[-1].lower() if "." in name else ""
    if suffix not in MIME_TYPES:
        raise DocumentError(f"'{name}' er verken PDF eller bilde. view_file støtter: {', '.join(MIME_TYPES)}.")
    if suffix == ".pdf":
        return _pdf_block(name, data)
    return _image_block(name, data)
