"""Tester for kalender (.ics) og møteinnkallinger i e-post. samples/invitasjon.ics er skrevet for hånd."""

import datetime
from email.message import EmailMessage
from pathlib import Path

from docxreader.tools import document_outline, read_document, read_section, search_document

SAMPLE = str(Path(__file__).parent.parent / "samples" / "invitasjon.ics")


def _section(heading: str, path: str = SAMPLE) -> str:
    return read_section.invoke({"path": path, "heading": heading})


def _ics(tmp_path, body: str) -> str:
    f = tmp_path / "k.ics"
    f.write_text("BEGIN:VCALENDAR\nVERSION:2.0\n" + body + "\nEND:VCALENDAR\n")
    return str(f)


def test_meta_and_one_section_per_event():
    outline = document_outline.invoke({"path": SAMPLE})
    assert outline.startswith("Kalender med 3 hendelser (invitasjon): Serverprosjektet.")
    assert "## Hendelse: Leverandørmøte med Nordic Data  (~" in outline
    assert "## Hendelse: Serverinstallasjon  (~" in outline


def test_event_details_with_timezone_people_and_escapes():
    result = _section("Leverandørmøte")
    assert "- Start: 2026-10-08 10:00 (Europe/Oslo)\n- Slutt: 2026-10-08 11:30 (Europe/Oslo)" in result
    assert "- Sted: Rom Fjorden, 3. etasje" in result  # \, -> ,
    assert "- Arrangør: Per Æsøy <per@fjellbekk.example>" in result
    assert "Kari Nordmann <kari@fjellbekk.example> (godtatt)" in result
    assert "Hansen: Ola <ola@fjellbekk.example> (foreløpig)" in result  # kolon i anførselstegn


def test_folded_line_is_joined_and_alarm_is_skipped():
    result = _section("Leverandørmøte")
    assert "Parkering i P-hus Vest, kode 4471." in result
    assert "Påminnelse" not in result


def test_all_day_event_end_is_exclusive():
    result = _section("Serverinstallasjon")
    assert "- Start: 2026-11-23 (hele dagen)\n- Slutt: 2026-11-24" in result


def test_utc_time_and_recurrence():
    result = _section("Oppfølgingsmøte")
    assert "- Start: 2026-12-01 13:00 (UTC)" in result
    assert "- Gjentas: FREQ=MONTHLY;COUNT=3" in result


def test_planted_fact_is_searchable():
    result = search_document.invoke({"path": SAMPLE, "query": "parkering"})
    assert "[Hendelse: Leverandørmøte med Nordic Data]" in result and "4471" in result


def test_single_day_event_has_no_separate_end(tmp_path):
    f = _ics(tmp_path, "BEGIN:VEVENT\nSUMMARY:Fridag\nDTSTART;VALUE=DATE:20261224\nDTEND;VALUE=DATE:20261225\nEND:VEVENT")
    result = read_document.invoke({"path": f})
    assert "- Start: 2026-12-24 (hele dagen)" in result and "Slutt" not in result


def test_event_without_title_or_times(tmp_path):
    f = _ics(tmp_path, "BEGIN:VEVENT\nDESCRIPTION:Bare tekst\nEND:VEVENT")
    assert "## Hendelse: uten tittel\n\nBare tekst" in read_document.invoke({"path": f})


def test_empty_calendar(tmp_path):
    assert "ingen tekst" in read_document.invoke({"path": _ics(tmp_path, "")})


def test_not_a_calendar(tmp_path):
    f = tmp_path / "x.ics"
    f.write_text("Dette er ikke en kalender.")
    result = read_document.invoke({"path": str(f)})
    assert result.startswith("Feil:") and "ikke en gyldig kalenderfil" in result


# --- Møteinnkallinger i e-post ------------------------------------------------


def test_invitation_inside_eml_is_read_as_ics(tmp_path):
    m = EmailMessage()
    m["From"], m["Subject"] = "per@x.no", "Invitasjon: Leverandørmøte"
    m.set_content("Du er invitert.")
    m.add_alternative(Path(SAMPLE).read_text(), subtype="calendar", params={"method": "REQUEST"})
    f = tmp_path / "invitasjon.eml"
    f.write_bytes(m.as_bytes())
    result = read_document.invoke({"path": str(f)})
    assert "Vedlegg: invitasjon.ics" in result
    assert "## Vedlegg: invitasjon.ics\n\nKalender med 3 hendelser" in result
    assert "### Hendelse: Leverandørmøte med Nordic Data" in result  # flyttet under vedlegget


def test_outlook_appointment_fields(tmp_path, monkeypatch):
    from types import SimpleNamespace

    from docxreader.readers import mail

    fake = SimpleNamespace(
        sender="Per <per@x.no>", to="Kari <kari@x.no>", cc=None, date=None, subject="Styremøte",
        body="Saksliste følger.", htmlBody=None, attachments=[], close=lambda: None,
        startDate=datetime.datetime(2026, 11, 14, 9, 0), endDate=datetime.datetime(2026, 11, 14, 15, 0),
        location="Ålesund\x00", organizer="Per Æsøy",
    )
    monkeypatch.setattr(mail.extract_msg, "openMsg", lambda path: fake)
    f = tmp_path / "avtale.msg"
    f.write_bytes(b"")
    result = read_document.invoke({"path": str(f)})
    assert result.startswith("Møteinnkalling (Outlook)\nFra: Per <per@x.no>")
    assert "Start: 2026-11-14 09:00\nSlutt: 2026-11-14 15:00\nSted: Ålesund\nArrangør: Per Æsøy" in result
