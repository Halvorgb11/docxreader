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
- PDF: maks MAX_PDF_BYTES og MAX_PDF_PAGES sider per kall. Større PDF-er leses
  i deler med pages="21-40" (pypdf klipper ut sidene – selve lesingen gjør Claude).
"""

import base64
import io
import re

from PIL import Image, UnidentifiedImageError
from pypdf import PdfReader, PdfWriter
from pypdf.errors import PyPdfError

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
MAX_PDF_PAGES = 20  # per kall; større PDF-er leses i deler med pages


def _encode(data: bytes) -> str:
    return base64.b64encode(data).decode("ascii")


def _image_block(name: str, data: bytes, pages: str = "") -> tuple[dict, str]:
    try:
        image = Image.open(io.BytesIO(data))
        image.load()
    except (UnidentifiedImageError, OSError):
        raise DocumentError(f"'{name}' er ikke et gyldig bilde (ødelagt eller feil format).")

    note = ""
    must_convert = False
    frames = getattr(image, "n_frames", 1)
    if pages and (frames == 1 or image.format == "GIF"):
        raise DocumentError(f"'{name}' har bare én side; pages gjelder PDF og TIFF med flere sider.")
    if frames > 1 and image.format != "GIF":  # f.eks. skannet TIFF med flere sider
        selected = parse_pages(pages, frames) if pages else [1]
        if len(selected) > 1:
            raise DocumentError(f"'{name}' er et bilde; velg én side om gangen, f.eks. pages='{selected[0]}'.")
        image.seek(selected[0] - 1)
        note = f" Viser side {selected[0]} av {frames}."
        must_convert = True  # bare den valgte siden skal sendes

    ok_as_is = (
        not must_convert
        and image.format in CLAUDE_IMAGE_FORMATS
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


def parse_pages(spec: str, total: int) -> list[int]:
    """"3", "1-20", "5-" (til slutten), "1,3,7-9" -> sidenumre (fra 1)."""
    pages: list[int] = []
    for part in spec.replace(" ", "").split(","):
        if not part:
            continue
        match = re.fullmatch(r"(\d+)(?:-(\d*))?", part)
        if not match:
            raise DocumentError(f"forstår ikke pages='{spec}'. Bruk f.eks. '3', '1-20' eller '1,4,7-9'.")
        first = int(match.group(1))
        last = first if match.group(2) is None else int(match.group(2) or total)
        if not 1 <= first <= last <= total:
            raise DocumentError(f"sidene '{part}' finnes ikke; dokumentet har {total} sider.")
        pages.extend(n for n in range(first, last + 1) if n not in pages)
    if not pages:
        raise DocumentError(f"pages='{spec}' velger ingen sider.")
    return pages


def _describe_pages(pages: list[int]) -> str:
    """[1, 2, 3, 7] -> "1–3, 7"."""
    ranges, start = [], pages[0]
    for previous, current in zip(pages, pages[1:] + [None]):
        if current != previous + 1:
            ranges.append(f"{start}" if start == previous else f"{start}–{previous}")
            start = current
    return ", ".join(ranges)


def _pdf_block(name: str, data: bytes, pages: str = "") -> tuple[dict, str]:
    """PDF -> innholdsblokk. pypdf brukes bare til å telle og klippe ut sider;
    selve lesingen gjør Claude."""
    if not data.startswith(b"%PDF-"):
        raise DocumentError(f"'{name}' er ikke en gyldig PDF-fil.")
    try:
        reader = PdfReader(io.BytesIO(data))
        if reader.is_encrypted and not reader.decrypt(""):
            # Mange PDF-er er "kryptert" med tomt passord (bare kopisperre); de går fint.
            raise DocumentError(f"'{name}' er passordbeskyttet og kan ikke leses.")
        total = len(reader.pages)
    except PyPdfError:
        raise DocumentError(f"'{name}' er ikke en gyldig PDF-fil (ødelagt).")

    if not pages:
        if total > MAX_PDF_PAGES:
            raise DocumentError(
                f"'{name}' har {total} sider; maks {MAX_PDF_PAGES} sider per kall (hver side koster "
                f"ca. 1500–3000 tokens). Velg sider med pages, f.eks. pages='1-{MAX_PDF_PAGES}'."
            )
        if len(data) > MAX_PDF_BYTES:
            raise DocumentError(f"'{name}' er for stor ({len(data) // 1_000_000} MB; maks {MAX_PDF_BYTES // 1_000_000} MB).")
        return {"type": "file", "base64": _encode(data), "mime_type": "application/pdf"}, f" {total} sider."

    selected = parse_pages(pages, total)
    if len(selected) > MAX_PDF_PAGES:
        raise DocumentError(f"pages='{pages}' velger {len(selected)} sider; maks {MAX_PDF_PAGES} per kall.")
    # Lag en ny PDF med bare de valgte sidene.
    writer = PdfWriter()
    for number in selected:
        writer.add_page(reader.pages[number - 1])
    buffer = io.BytesIO()
    writer.write(buffer)
    part = buffer.getvalue()
    if len(part) > MAX_PDF_BYTES:
        raise DocumentError(f"de valgte sidene er for store ({len(part) // 1_000_000} MB); velg færre sider.")
    note = f" Side {_describe_pages(selected)} av {total}."
    return {"type": "file", "base64": _encode(part), "mime_type": "application/pdf"}, note


def media_block(name: str, data: bytes, pages: str = "") -> tuple[dict, str]:
    """Fil -> (LangChain-innholdsblokk, merknad om hva som ble gjort).
    `pages` velger sider i en PDF ("1-20") eller én side i en flersidig TIFF."""
    suffix = "." + name.rsplit(".", 1)[-1].lower() if "." in name else ""
    if suffix not in MIME_TYPES:
        raise DocumentError(f"'{name}' er verken PDF eller bilde. view_file støtter: {', '.join(MIME_TYPES)}.")
    if suffix == ".pdf":
        return _pdf_block(name, data, pages)
    return _image_block(name, data, pages)
