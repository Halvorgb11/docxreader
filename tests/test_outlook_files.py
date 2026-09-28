"""Tester mot ekte Outlook-filer (samples/outlook/, MIT-lisens – se README der).

Falske objekter (test_mail.py) tester vår oversettelse; disse testene fanger
det falske objekter ikke kan: hvordan ekte .msg-filer faktisk ser ut. Det var
slik vi fant null-tegn, ufullstendige mottakerlister og kontaktkort."""

from pathlib import Path

from docxreader.tools import read_document, read_section, search_document, view_file

OUTLOOK = Path(__file__).parent.parent / "samples" / "outlook"


def _read(name: str) -> str:
    return read_document.invoke({"path": str(OUTLOOK / name)})


def test_nested_msg_with_pdf_attachments():
    result = _read("EmailWithInnerMailAndAttachments.msg")
    assert result.startswith("E-post\nEmne: Outer mail\nVedlegg: OUTER 1.pdf, OUTER 2.pdf, Inner mail.msg")
    assert "## Vedlegg: Inner mail.msg\n\nE-post\nEmne: Inner mail\nVedlegg: INNER 1.pdf, INNER 2.pdf" in result
    assert "### Vedlegg: INNER 1.pdf" in result


def test_pdf_inside_inner_msg_can_be_viewed():
    path = str(OUTLOOK / "EmailWithInnerMailAndAttachments.msg")
    text, block = view_file.invoke({"path": path, "attachment": "Inner mail.msg > INNER 1.pdf"})
    assert text["text"].startswith("Innholdet i 'INNER 1.pdf' følger.")
    assert block["mime_type"] == "application/pdf"


def test_french_subject_and_sender():
    result = _read("EmailWithSpecialCharsInSubject_2.msg")
    assert "Fra: Julie Clarebots <Julie.Clarebots@qbere.com>" in result
    assert "Emne: Un sujet très bien défini" in result
    assert "\x00" not in result


def test_russian_text():
    assert "«Имя пользователя»" in _read("RtfWithShortRussianString.msg")


def test_contact_card_shows_fields_instead_of_empty_mail():
    assert _read("kontakt-Swetlana.msg") == (
        "Kontakt (Outlook)\n"
        "Navn: Swetlana Novikova\n"
        "Firma: Scalabium Software\n"
        "E-post: support@scalabium.com\n"
        "Nettside: http://www.scalabium.com\n\n"
        "(Dokumentet inneholder ingen tekst utover dette.)"
    )


def test_html_mail_with_text_attachment():
    path = str(OUTLOOK / "mapi-sample.msg")
    assert "is working as expected" in read_section.invoke({"path": path, "heading": "Melding"})
    assert "plain text attachment" in read_section.invoke({"path": path, "heading": "attachment.txt"})
    assert "[Vedlegg: attachment.txt]" in search_document.invoke({"path": path, "query": "plain text"})
