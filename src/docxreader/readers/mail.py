"""Leser for e-post: .eml (standardformatet, `email` fra standardbiblioteket)
og .msg (Outlook, biblioteket `extract-msg`).

Begge formatene gjøres først om til samme mellomform, `Mail`. Deretter lages
blokkene likt:

    (metatekst: Fra, Til, Kopi, Dato, Emne, Vedlegg)
    ## Melding                 selve teksten
    ## Tidligere melding 1     sitert tråd ("Den … skrev:", "-----Original Message-----")
    ## Vedlegg: tilbud.xlsx    vedlegget lest med leseren for sin filtype
    ### Ark: …                 (overskriftene i vedlegget flyttes ned under vedlegget)

Vedleggene sendes videre til menyen i readers/__init__.py – et .xlsx-vedlegg
leses av Excel-leseren, en videresendt e-post av denne leseren igjen.
"""

import email
import email.policy
import email.utils
import mimetypes
import re
import tempfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import extract_msg
from extract_msg.enums import AttachmentType, RecipientType
from extract_msg.exceptions import ExMsgBaseException

from docxreader.blocks import Block, DocumentError, count_words, heading, shift_headings
from docxreader.readers.htmltext import html_to_text
from docxreader.readers.media import MIME_TYPES as MEDIA_TYPES
from docxreader.readers.text import markdown_blocks

SECTION_LEVEL = 2  # "## Melding", "## Vedlegg: x"
MAX_DEPTH = 3  # e-post i e-post i e-post … stopper her
MAX_ATTACHMENT_BYTES = 20 * 1024 * 1024


@dataclass
class Attachment:
    name: str
    data: bytes | None = None  # innholdet (vanlig fil)
    mail: "Mail | None" = None  # en e-post som vedlegg
    note: str = ""  # hvorfor innholdet mangler, hvis det mangler


@dataclass
class Mail:
    sender: str = ""
    to: str = ""
    cc: str = ""
    date: str = ""
    subject: str = ""
    text: str = ""  # brødteksten som markdown-lignende tekst
    attachments: list[Attachment] = field(default_factory=list)


def _format_date(value) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M")
    return str(value or "")


def _safe_name(name: str, content_type: str, number: int) -> str:
    """Bare filnavnet (aldri en sti), og en filendelse hvis den mangler."""
    name = Path((name or "").replace("\\", "/")).name.strip()
    if not name:
        name = f"vedlegg-{number}{mimetypes.guess_extension(content_type or '') or ''}"
    return name


# --- .eml ------------------------------------------------------------------


def _mail_from_eml(message) -> Mail:
    body = message.get_body(preferencelist=("plain", "html"))
    text = ""
    if body is not None:
        content = body.get_content()
        text = html_to_text(content) if body.get_content_type() == "text/html" else content

    attachments = []
    for number, part in enumerate(message.iter_attachments(), start=1):
        content_type = part.get_content_type()
        # Bilder som er "innebygd" i HTML-teksten (logoer o.l.) er ikke vedlegg.
        if content_type.startswith("image/") and part["Content-ID"] and part.get_content_disposition() != "attachment":
            continue
        if content_type == "message/rfc822":  # videresendt e-post som vedlegg
            inner = part.get_payload()[0]
            # Uten filnavn: bruk emnet ("Tilbud 2026-117.eml").
            subject = str(inner["Subject"] or "").replace("/", "-")
            name = _safe_name(part.get_filename() or (subject and subject + ".eml"), content_type, number)
            attachments.append(Attachment(name if name.endswith(".eml") else name + ".eml", mail=_mail_from_eml(inner)))
            continue
        name = _safe_name(part.get_filename(), content_type, number)
        attachments.append(Attachment(name, data=part.get_payload(decode=True) or b""))

    return Mail(
        sender=str(message["From"] or ""),
        to=str(message["To"] or ""),
        cc=str(message["Cc"] or ""),
        date=_format_date(email.utils.parsedate_to_datetime(message["Date"]) if message["Date"] else ""),
        subject=str(message["Subject"] or ""),
        text=text,
        attachments=attachments,
    )


def _load_eml_mail(path: Path) -> Mail:
    raw = path.read_bytes()
    if b"\x00" in raw[:1024]:
        raise DocumentError(f"'{path}' er ikke en gyldig e-postfil (binærdata).")
    # policy.default gir moderne EmailMessage med get_body(), og dekoder
    # overskrifter som "=?utf-8?q?M=C3=B8te?=" til "Møte" automatisk.
    message = email.message_from_bytes(raw, policy=email.policy.default)
    if not (message["From"] or message["Subject"] or message["Date"]):
        raise DocumentError(f"'{path}' ser ikke ut som en e-post (mangler Fra, Emne og Dato).")
    return _mail_from_eml(message)


def load_eml(path: Path) -> tuple[list[Block], str]:
    return _mail_blocks(_load_eml_mail(path))


# --- .msg (Outlook) --------------------------------------------------------


def _clean(value) -> str:
    """Outlook-tekst fylles ofte med null-tegn (\x00) til slutt; fjern dem."""
    return str(value or "").replace("\x00", "").strip()


def _field(obj, name: str, default=None):
    """Les ett felt uten å krasje. extract-msg dekoder feltene først når de
    leses, og ett felt med ødelagt tegnkoding skal ikke ødelegge resten."""
    try:
        return getattr(obj, name, default)
    except Exception:
        return default


def _recipients(msg, kind: RecipientType) -> str:
    """Mottakere av én type (Til, Kopi) fra mottakerlisten. (msg.to er ikke
    alltid komplett.) Faller tilbake på msg.to/msg.cc hvis listen mangler."""
    people = []
    for recipient in _field(msg, "recipients", None) or []:
        if _field(recipient, "type") != kind:
            continue
        name, address = _clean(_field(recipient, "name")), _clean(_field(recipient, "email"))
        person = f"{name} <{address}>" if name and address and name != address else (name or address)
        if person and person not in people:
            people.append(person)
    if people:
        return ", ".join(people)
    return _clean(_field(msg, "to" if kind == RecipientType.TO else "cc"))


def _mail_from_msg(msg) -> Mail:
    """Outlook-melding (fra extract-msg) -> Mail. Alle felt leses med _field,
    fordi .msg-filer også kan være kalenderinvitasjoner o.l. uten alle feltene,
    og fordi enkeltfelt kan ha ødelagt tegnkoding."""
    text = _clean(_field(msg, "body"))
    html = _field(msg, "htmlBody")
    if not text and html:
        text = html_to_text(html.decode("utf-8", errors="replace").replace("\x00", ""))

    attachments = []
    for number, att in enumerate(_field(msg, "attachments", []) or [], start=1):
        mimetype = getattr(att, "mimetype", "") or ""
        if getattr(att, "hidden", False) or (getattr(att, "contentId", None) and mimetype.startswith("image/")):
            continue  # innebygde bilder
        name = _safe_name(_clean(getattr(att, "longFilename", None) or getattr(att, "shortFilename", None)
                                 or getattr(att, "name", None)), mimetype, number)
        kind = getattr(att, "type", None)
        if kind == AttachmentType.MSG:
            attachments.append(Attachment(name if name.endswith(".msg") else name + ".msg", mail=_mail_from_msg(att.data)))
        elif kind == AttachmentType.DATA and isinstance(att.data, bytes):
            attachments.append(Attachment(name, data=att.data))
        else:
            attachments.append(Attachment(name, note="vedleggstypen kan ikke leses"))

    return Mail(
        sender=_clean(_field(msg, "sender")),
        to=_recipients(msg, RecipientType.TO),
        cc=_recipients(msg, RecipientType.CC),
        date=_format_date(_field(msg, "date")),
        subject=_clean(_field(msg, "subject")),
        text=text,
        attachments=attachments,
    )


def _load_msg_mail(path: Path) -> Mail:
    try:
        msg = extract_msg.openMsg(str(path))
    except (ExMsgBaseException, OSError):
        raise DocumentError(f"'{path}' er ikke en gyldig Outlook-fil (.msg) (ødelagt eller feil format).")
    try:
        return _mail_from_msg(msg)  # leser alt (også vedlegg) før filen lukkes
    finally:
        msg.close()


def load_msg(path: Path) -> tuple[list[Block], str]:
    return _mail_blocks(_load_msg_mail(path))


def load_mail(path: Path) -> Mail:
    """Les en .eml- eller .msg-fil til en Mail (uten å lage blokker)."""
    return _load_msg_mail(path) if path.suffix.lower() == ".msg" else _load_eml_mail(path)


def find_attachment(mail: Mail, wanted: str, prefer: tuple[str, ...] = ()) -> Attachment:
    """Finn et vedlegg ved navn. Uten hensyn til store/små bokstaver; eksakt
    treff før delvis. Vedlegg i en videresendt e-post nås med " > ":
    "Tilbud 2026-117.eml > data.csv". Passer flere, velges det ene (om bare
    ett) med en filendelse i `prefer`, f.eks. (".xlsx", ".csv") for query_table."""
    parts = [p.strip() for p in wanted.split(">") if p.strip()]
    if not parts:
        raise DocumentError("vedleggsnavnet kan ikke være tomt.")
    current = mail
    for number, part in enumerate(parts):
        names = [a.name for a in current.attachments]
        hits = [a for a in current.attachments if a.name.casefold() == part.casefold()]
        hits = hits or [a for a in current.attachments if part.casefold() in a.name.casefold()]
        last = number == len(parts) - 1
        preferred = [a for a in hits if Path(a.name).suffix.lower() in prefer]
        if last and len(hits) > 1 and len(preferred) == 1:
            hits = preferred
        if len(hits) != 1:
            problem = "flere vedlegg passer" if hits else "fant ikke vedlegget"
            raise DocumentError(f"{problem} '{part}'. Vedlegg: {', '.join(names) or '(ingen)'}.")
        found = hits[0]
        if last:
            return found
        if found.mail is None:
            raise DocumentError(f"'{found.name}' er ikke en e-post, så den har ingen vedlegg.")
        current = found.mail
    raise AssertionError("unreachable")


# --- Tråder: dele teksten i meldinger ----------------------------------------

REPLY_MARKERS = [
    # -----Original Message----- / -----Opprinnelig melding----- / ---------- Forwarded message ---------
    re.compile(r"^-{2,}\s*(original message|opprinnelig melding|videresendt melding|forwarded message)\s*-{2,}$", re.I),
    # "Den 25. sep. 2026 kl. 14:02 skrev Ola Hansen <ola@x.no>:" / "On Thu, … wrote:"
    # Navnet kan stå etter "skrev" (norsk) eller før "wrote" (engelsk).
    re.compile(r"^(den|on)\s.+\s(skrev|wrote)\b.*:$", re.I),
]
OUTLOOK_FROM = re.compile(r"^(fra|from)\s*:", re.I)
OUTLOOK_SENT = re.compile(r"^(sendt|sent|dato|date)\s*:", re.I)


def _find_reply_marker(lines: list[str], start: int) -> int | None:
    """Linjen der en sitert, tidligere melding begynner (fra og med `start`)."""
    for i in range(start, len(lines)):
        line = lines[i].strip().lstrip("> ").strip()
        if any(marker.match(line) for marker in REPLY_MARKERS):
            return i
        # Outlook skriver et lite hode: "Fra: …" og like under "Sendt: …".
        # (Står det rett under "-----Opprinnelig melding-----", hører det til den.)
        previous = lines[i - 1].strip().lstrip("> ").strip() if i > 0 else ""
        if any(marker.match(previous) for marker in REPLY_MARKERS):
            continue
        if OUTLOOK_FROM.match(line) and any(
            OUTLOOK_SENT.match(next_line.strip().lstrip("> ").strip()) for next_line in lines[i + 1 : i + 4]
        ):
            return i
    return None


HEADER_LINE = re.compile(r"^(fra|from|sendt|sent|dato|date|til|to|kopi|cc|emne|subject)\s*:", re.I)


def _header_lines_as_list(lines: list[str]) -> list[str]:
    """Et sitert e-posthode ("Fra: …", "Sendt: …", "Emne: …") blir en liste,
    så linjene ikke slås sammen til ett avsnitt."""
    result, in_header = [], False
    for line in lines:
        if HEADER_LINE.match(line.strip()):
            if not in_header and result and result[-1].strip():
                result.append("")  # tom linje før lista
            result.append("- " + line.strip())
            in_header = True
        else:
            if in_header and line.strip():
                result.append("")  # tom linje etter lista
            in_header = False
            result.append(line)
    return result


def _split_thread(text: str) -> list[tuple[str, str]]:
    """Del e-postteksten i (tittel, tekst): selve meldingen og hver sitert melding."""
    parts = []
    title, lines = "Melding", text.splitlines()
    number = 0
    while True:
        # Første del søkes fra linje 0; siterte deler starter MED markøren, så
        # der søker vi fra linje 1 (ellers ville vi funnet samme markør igjen).
        i = _find_reply_marker(lines, 0 if number == 0 else 1)
        if i is None:
            parts.append((title, "\n".join(lines)))
            return parts
        parts.append((title, "\n".join(lines[:i])))
        number += 1
        title, lines = f"Tidligere melding {number}", lines[i:]
        # Sitert med ">" foran hver linje? Fjern ett nivå.
        if all(line.startswith(">") or not line.strip() for line in lines):
            lines = [re.sub(r"^> ?", "", line) for line in lines]
        elif all(line.startswith(">") or not line.strip() for line in lines[1:]):
            lines = [lines[0]] + [re.sub(r"^> ?", "", line) for line in lines[1:]]
        # Markørlinjen ("Den … skrev:") er et eget avsnitt, også uten tom linje etter.
        # Er markøren selv et Outlook-hode ("Fra: …"), hører den med i hode-lista.
        if HEADER_LINE.match(lines[0].strip()):
            lines = _header_lines_as_list(lines)
        else:
            lines = [lines[0], ""] + _header_lines_as_list(lines[1:])


# --- Mail -> blokker ----------------------------------------------------------


def _meta(mail: Mail) -> str:
    lines = ["E-post"]
    for label, value in [("Fra", mail.sender), ("Til", mail.to), ("Kopi", mail.cc),
                         ("Dato", mail.date), ("Emne", mail.subject)]:
        if value:
            lines.append(f"{label}: {' '.join(str(value).split())}")
    if mail.attachments:
        lines.append("Vedlegg: " + ", ".join(a.name for a in mail.attachments))
    return "\n".join(lines)


def _nest(blocks: list[Block]) -> list[Block]:
    """Flytt overskriftene i et vedlegg ned, så den øverste havner på nivå 3 –
    rett under "## Vedlegg: …" (Word-tittel # -> ###, Excel-ark ## -> ###)."""
    levels = [b.heading_level for b in blocks if b.heading_level]
    return shift_headings(blocks, SECTION_LEVEL + 1 - min(levels)) if levels else blocks


def _attachment_blocks(attachment: Attachment, depth: int) -> list[Block]:
    """Les et vedlegg med leseren for sin filtype. Overskriftene i vedlegget
    flyttes ned, så de havner under "## Vedlegg: …"."""
    from docxreader.readers import READERS  # her, for å unngå sirkulær import

    def note(text: str) -> list[Block]:
        return [Block(f"({text})", count_words(text))]

    if attachment.mail is not None:
        if depth >= MAX_DEPTH:
            return note("videresendt e-post for dypt nøstet; ikke lest")
        blocks, meta = _mail_blocks(attachment.mail, depth + 1)
        return [Block(meta, count_words(meta)), *_nest(blocks)]
    if attachment.note:
        return note(attachment.note)

    data = attachment.data or b""
    size = f"{max(1, round(len(data) / 1024))} KB"
    suffix = Path(attachment.name).suffix.lower()
    reader = READERS.get(suffix)
    if reader is None and suffix in MEDIA_TYPES:
        return note(f"{size}; PDF/bilde – se innholdet med view_file og attachment=\"{attachment.name}\"")
    if reader is None:
        return note(f"{size}; filtypen {suffix or '(ingen)'} kan ikke leses her")
    if len(data) > MAX_ATTACHMENT_BYTES:
        return note(f"{size}; for stort til å leses")

    # Leserne tar en filsti, så vedlegget skrives til en midlertidig fil.
    with tempfile.TemporaryDirectory() as folder:
        file = Path(folder) / attachment.name
        file.write_bytes(data)
        try:
            blocks, meta = reader(file)
        except DocumentError as e:
            # Vis vedleggets navn, ikke stien til den midlertidige filen.
            return note(f"{size}; kunne ikke leses: {str(e).replace(str(file), attachment.name)}")
    result = [Block(meta, count_words(meta))] if meta else []
    return result + (_nest(blocks) if blocks else note("vedlegget er tomt"))


def _mail_blocks(mail: Mail, depth: int = 0) -> tuple[list[Block], str]:
    blocks: list[Block] = []
    for title, text in _split_thread(mail.text):
        if title != "Melding" and not text.strip():
            continue
        blocks.append(heading(SECTION_LEVEL, title))
        # Overskrifter inni teksten (fra HTML) flyttes under "## Melding".
        body = shift_headings(markdown_blocks(text), SECTION_LEVEL)
        blocks.extend(body or [Block("(ingen tekst)", 2)])

    for attachment in mail.attachments:
        blocks.append(heading(SECTION_LEVEL, f"Vedlegg: {attachment.name}"))
        blocks.extend(_attachment_blocks(attachment, depth))
    return blocks, _meta(mail)
