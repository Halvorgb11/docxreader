"""Tester for view_file: PDF og bilder sendes som innholdsblokker til modellen.
Testfilene lages av scripts/make_sample_media.py."""

import base64
import io
from pathlib import Path

import pytest
from langchain_core.messages import ToolMessage
from PIL import Image

from docxreader import _preview
from docxreader.readers import media
from docxreader.tools import read_document, view_file

SAMPLES = Path(__file__).parent.parent / "samples"


def _image_bytes(fmt: str, size=(100, 50), mode="RGB") -> bytes:
    buffer = io.BytesIO()
    Image.new(mode, size, "white").save(buffer, fmt)
    return buffer.getvalue()


def _view(path, **kwargs):
    return view_file.invoke({"path": str(path), **kwargs})


def _decoded(block) -> Image.Image:
    return Image.open(io.BytesIO(base64.b64decode(block["base64"])))


def test_png_is_passed_through_as_image_block():
    text, block = _view(SAMPLES / "kvittering.png")
    assert text == {"type": "text", "text": "Innholdet i 'kvittering.png' følger."}
    assert block["type"] == "image" and block["mime_type"] == "image/png"
    assert base64.b64decode(block["base64"]) == (SAMPLES / "kvittering.png").read_bytes()  # uendret


def test_pdf_is_passed_as_file_block_with_page_count():
    text, block = _view(SAMPLES / "moteinnkalling.pdf")
    assert text["text"] == "Innholdet i 'moteinnkalling.pdf' følger. 2 sider."
    assert block == {"type": "file", "base64": block["base64"], "mime_type": "application/pdf"}


def test_image_attachment_in_email():
    text, block = _view(SAMPLES / "tilbud.eml", attachment="skisse")
    assert "skisse.png" in text["text"]
    assert _decoded(block).size == (700, 360)


def test_tiff_is_converted_to_a_format_claude_accepts(tmp_path):
    f = tmp_path / "skann.tif"
    f.write_bytes(_image_bytes("TIFF"))
    _, block = _view(f)
    assert block["mime_type"] == "image/jpeg"
    assert _decoded(block).format == "JPEG"


def test_transparent_bmp_becomes_png(tmp_path):
    f = tmp_path / "logo.bmp"
    f.write_bytes(_image_bytes("BMP", mode="P"))
    _, block = _view(f)
    assert block["mime_type"] == "image/png"


def test_large_image_is_scaled_down(tmp_path):
    f = tmp_path / "stort.png"
    f.write_bytes(_image_bytes("PNG", size=(4000, 2000)))
    text, block = _view(f)
    assert "Skalert ned fra 4000×2000 til 1568×784 piksler." in text["text"]
    assert _decoded(block).size == (1568, 784)


@pytest.mark.parametrize("name, content, message", [
    ("ødelagt.png", b"ikke et bilde", "er ikke et gyldig bilde"),
    ("falsk.pdf", b"ikke en pdf", "er ikke en gyldig PDF-fil"),
    ("lang.pdf", b"%PDF-1.4\n" + b"<< /Type /Page >>\n" * 101 + b"<< /Type /Pages >>", "har 101 sider"),
    ("data.xlsx", b"x", "er verken PDF eller bilde"),
])
def test_errors_are_text(tmp_path, name, content, message):
    f = tmp_path / name
    f.write_bytes(content)
    result = _view(f)
    assert isinstance(result, str) and result.startswith("Feil:") and message in result


def test_too_large_pdf(tmp_path, monkeypatch):
    monkeypatch.setattr(media, "MAX_PDF_BYTES", 10)
    f = tmp_path / "stor.pdf"
    f.write_bytes(b"%PDF-1.4 " + b"x" * 100)
    assert "er for stor" in _view(f)


def test_attachment_that_is_a_document_is_rejected():
    assert "er verken PDF eller bilde" in _view(SAMPLES / "tilbud.eml", attachment="avtaleutkast.docx")


def test_read_document_points_to_view_file_for_images_and_pdf():
    for name in ["kvittering.png", "moteinnkalling.pdf"]:
        assert "PDF og bilder kan du se med view_file." in read_document.invoke({"path": str(SAMPLES / name)})


def test_cli_preview_does_not_print_base64():
    message = ToolMessage(content=_view(SAMPLES / "kvittering.png"), name="view_file", tool_call_id="1")
    preview = _preview(message.content)
    assert preview.startswith("Innholdet i 'kvittering.png' følger. [image: image/png, ~")
    assert len(preview) < 100
