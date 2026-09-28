"""Leser for kalenderfiler (.ics / iCalendar). Trenger ingen ekstra bibliotek.

En .ics-fil er tekst med linjer på formen NAVN;PARAMETER=verdi:VERDI, gruppert
i BEGIN:VEVENT … END:VEVENT (én hendelse). Hver hendelse blir en seksjon:

    ## Hendelse: Leverandørmøte
    - Start: 2026-10-08 10:00 (Europe/Oslo)
    - Slutt: …   - Sted: …   - Arrangør: …   - Deltakere: …
    (beskrivelsen som avsnitt)

Fallgruver i formatet:
- Lange linjer brytes: en linje som starter med mellomrom fortsetter forrige.
- Tegn er "escapet": \\n = linjeskift, \\, = komma, \\; = semikolon.
- Tidspunkt: 20261008T100000 med TZID (lokal tid), 20261008T080000Z (UTC)
  eller VALUE=DATE:20261008 (hele dagen).
- Heldagshendelser har EKSKLUSIV slutt: DTEND er dagen ETTER siste dag.
- Gjentakelser (RRULE, f.eks. "FREQ=WEEKLY;BYDAY=MO") regnes om til konkrete
  datoer med python-dateutil, minus unntak (EXDATE), og beskrives på norsk.
"""

import datetime
import itertools
import re
from pathlib import Path

from dateutil.rrule import rruleset, rrulestr

from docxreader.blocks import Block, DocumentError, count_words, heading, shift_headings
from docxreader.readers.text import markdown_blocks, read_text

EVENT_LEVEL = 2
METHODS = {"REQUEST": "invitasjon", "CANCEL": "avlysning", "REPLY": "svar på invitasjon", "PUBLISH": "publisert kalender"}
ANSWERS = {"ACCEPTED": "godtatt", "DECLINED": "avslått", "TENTATIVE": "foreløpig", "NEEDS-ACTION": "ikke svart"}


def _unfold(text: str) -> list[str]:
    """Slå sammen brutte linjer (fortsettelseslinjer starter med mellomrom/tab)."""
    lines: list[str] = []
    for line in text.splitlines():
        if line[:1] in (" ", "\t") and lines:
            lines[-1] += line[1:]
        elif line.strip():
            lines.append(line)
    return lines


def _parse_line(line: str) -> tuple[str, dict[str, str], str]:
    """'DTSTART;TZID=Europe/Oslo:20261008T100000' -> ("DTSTART", {"TZID": …}, "2026…").
    Kolon inni anførselstegn i parametere (f.eks. CN="Hansen: Ola") hører ikke
    til skillet mellom navn og verdi."""
    in_quotes = False
    for i, char in enumerate(line):
        if char == '"':
            in_quotes = not in_quotes
        elif char == ":" and not in_quotes:
            head, value = line[:i], line[i + 1 :]
            break
    else:
        return line.upper(), {}, ""
    name, *params = head.split(";")
    parsed = {}
    for param in params:
        key, _, val = param.partition("=")
        parsed[key.upper()] = val.strip('"')
    return name.upper(), parsed, value


def _unescape(value: str) -> str:
    return (value.replace("\\n", "\n").replace("\\N", "\n").replace("\\,", ",")
            .replace("\\;", ";").replace("\\\\", "\\"))


def _parse_time(value: str, params: dict[str, str]) -> tuple[datetime.date | datetime.datetime | None, str]:
    """-> (dato/tidspunkt, tekst for visning)."""
    value = value.strip()
    try:
        if params.get("VALUE") == "DATE" or len(value) == 8:
            day = datetime.datetime.strptime(value, "%Y%m%d").date()
            return day, day.isoformat()
        moment = datetime.datetime.strptime(value.rstrip("Z"), "%Y%m%dT%H%M%S")
    except ValueError:
        return None, value  # ukjent format: vis som det står
    zone = "UTC" if value.endswith("Z") else params.get("TZID", "")
    return moment, moment.strftime("%Y-%m-%d %H:%M") + (f" ({zone})" if zone else "")


MAX_OCCURRENCES = 10
FREQUENCIES = {"DAILY": ("daglig", "dag"), "WEEKLY": ("ukentlig", "uke"), "MONTHLY": ("månedlig", "måned"),
               "YEARLY": ("årlig", "år"), "HOURLY": ("hver time", "time")}
DAYS = {"MO": "mandag", "TU": "tirsdag", "WE": "onsdag", "TH": "torsdag", "FR": "fredag", "SA": "lørdag", "SU": "søndag"}
ORDINALS = {1: "første", 2: "andre", 3: "tredje", 4: "fjerde", 5: "femte", -1: "siste", -2: "nest siste"}


def _describe_rule(rule: str) -> str:
    """"FREQ=MONTHLY;BYDAY=-1FR;COUNT=3" -> "månedlig på siste fredag, 3 ganger"."""
    parts = dict(p.split("=", 1) for p in rule.upper().split(";") if "=" in p)
    frequency = parts.get("FREQ", "")
    word, unit = FREQUENCIES.get(frequency, (frequency.lower(), ""))
    interval = int(parts.get("INTERVAL", "1") or 1)
    text = f"hver {interval}. {unit}" if interval > 1 and unit else word
    if "BYDAY" in parts:
        days = []
        for day in parts["BYDAY"].split(","):
            m = re.fullmatch(r"([+-]?\d+)?([A-Z]{2})", day)
            if not m:
                continue
            name = DAYS.get(m.group(2), m.group(2))
            days.append(f"{ORDINALS.get(int(m.group(1)), m.group(1) + '.')} {name}" if m.group(1) else name)
        text += " på " + (", ".join(days[:-1]) + " og " + days[-1] if len(days) > 1 else days[0]) if days else ""
    if "BYMONTHDAY" in parts:
        text += f" den {parts['BYMONTHDAY'].replace(',', '., ')}."
    if "COUNT" in parts:
        text += f", {parts['COUNT']} ganger"
    if "UNTIL" in parts:
        until, until_text = _parse_time(parts["UNTIL"], {})
        text += f", til og med {until_text.split(' ')[0] if until else parts['UNTIL']}"
    return text


def _occurrences(rule: str, start, whole_day: bool, exdates: list, rdates: list) -> str:
    """Konkrete datoer for en gjentakelse (maks MAX_OCCURRENCES vist)."""
    first = datetime.datetime.combine(start, datetime.time()) if whole_day else start
    # dateutil godtar ikke UNTIL i UTC ("…Z") når start er uten tidssone; vi
    # regner uten tidssoner (tidspunktene vises i samme sone som starten).
    rule = re.sub(r"(UNTIL=\d{8}(T\d{6})?)Z", r"\1", rule, flags=re.IGNORECASE)
    dates = rruleset()
    dates.rrule(rrulestr(rule, dtstart=first))
    for day in exdates:
        dates.exdate(datetime.datetime.combine(day, datetime.time()) if not isinstance(day, datetime.datetime) else day)
    for day in rdates:
        dates.rdate(datetime.datetime.combine(day, datetime.time()) if not isinstance(day, datetime.datetime) else day)

    shown = list(itertools.islice(dates, MAX_OCCURRENCES + 1))
    fmt = "%Y-%m-%d" if whole_day else "%Y-%m-%d %H:%M"
    text = ", ".join(d.strftime(fmt) for d in shown[:MAX_OCCURRENCES])
    if len(shown) > MAX_OCCURRENCES:
        endless = "COUNT" not in rule.upper() and "UNTIL" not in rule.upper()
        text += " … (fortsetter uten slutt)" if endless else f" … ({dates.count()} totalt)"
    return text


def _person(value: str, params: dict[str, str]) -> str:
    address = value.removeprefix("mailto:").removeprefix("MAILTO:")
    name = params.get("CN", "")
    person = f"{name} <{address}>" if name and name != address else address
    answer = ANSWERS.get(params.get("PARTSTAT", ""), "")
    return person + (f" ({answer})" if answer else "")


def _event_blocks(props: list[tuple[str, dict, str]]) -> list[Block]:
    first = {}
    attendees = []
    exdates, rdates = [], []  # unntak og ekstra datoer for gjentakelser (kan stå på flere linjer)
    for name, params, value in props:
        if name == "ATTENDEE":
            attendees.append(_person(value, params))
        elif name in ("EXDATE", "RDATE"):
            for item in value.split(","):
                moment, _ = _parse_time(item, params)
                if moment is not None:
                    (exdates if name == "EXDATE" else rdates).append(moment)
        else:
            first.setdefault(name, (params, value))

    def text(name: str) -> str:
        return _unescape(first[name][1]).strip() if name in first else ""

    lines = []
    start, start_text = _parse_time(first["DTSTART"][1], first["DTSTART"][0]) if "DTSTART" in first else (None, "")
    end, end_text = _parse_time(first["DTEND"][1], first["DTEND"][0]) if "DTEND" in first else (None, "")
    whole_day = isinstance(start, datetime.date) and not isinstance(start, datetime.datetime)
    if whole_day and isinstance(end, datetime.date) and end > start:
        end = end - datetime.timedelta(days=1)  # DTEND er dagen ETTER siste dag
        end_text = end.isoformat()
    if start_text:
        lines.append(f"Start: {start_text}" + (" (hele dagen)" if whole_day else ""))
    if end_text and end_text != start_text:
        lines.append(f"Slutt: {end_text}")
    for label, name in [("Sted", "LOCATION"), ("Status", "STATUS")]:
        if text(name):
            lines.append(f"{label}: {text(name)}")
    if "RRULE" in first:
        rule = first["RRULE"][1]
        try:  # regn ut alt før noe legges til, så en feil ikke gir halve linjer
            described = [f"Gjentas: {_describe_rule(rule)}"]
            if start is not None:
                described.append(f"Datoer: {_occurrences(rule, start, whole_day, exdates, rdates)}")
            lines.extend(described)
        except (ValueError, TypeError):
            lines.append(f"Gjentas: {rule} (regelen kunne ikke tolkes)")
    if "ORGANIZER" in first:
        lines.append(f"Arrangør: {_person(first['ORGANIZER'][1], first['ORGANIZER'][0])}")
    if attendees:
        lines.append("Deltakere: " + ", ".join(attendees))

    title = text("SUMMARY") or "uten tittel"
    blocks = [heading(EVENT_LEVEL, f"Hendelse: {title}")]
    if lines:
        listing = "\n".join(f"- {line}" for line in lines)
        blocks.append(Block(listing, count_words(listing)))
    description = text("DESCRIPTION")
    if description:
        blocks.extend(shift_headings(markdown_blocks(description), EVENT_LEVEL))
    return blocks


def calendar_blocks(text: str, name: str = "kalenderen") -> tuple[list[Block], str]:
    lines = _unfold(text)
    if not any(line.upper().startswith("BEGIN:VCALENDAR") for line in lines):
        raise DocumentError(f"'{name}' er ikke en gyldig kalenderfil (mangler BEGIN:VCALENDAR).")

    blocks: list[Block] = []
    events = 0
    method = calendar_name = ""
    current: list | None = None
    depth = 0  # VALARM (påminnelse) ligger inni VEVENT; dens linjer hopper vi over
    for line in lines:
        name_, params, value = _parse_line(line)
        if name_ == "BEGIN" and value.upper() == "VEVENT":
            current, depth = [], 0
        elif name_ == "BEGIN" and current is not None:
            depth += 1
        elif name_ == "END" and value.upper() == "VEVENT" and current is not None:
            blocks.extend(_event_blocks(current))
            events += 1
            current = None
        elif name_ == "END" and current is not None:
            depth -= 1
        elif current is not None and depth == 0:
            current.append((name_, params, value))
        elif name_ == "METHOD":
            method = METHODS.get(value.upper(), value.lower())
        elif name_ == "X-WR-CALNAME":
            calendar_name = _unescape(value)

    meta = f"Kalender med {events} hendelse" + ("r" if events != 1 else "")
    if method:
        meta += f" ({method})"
    if calendar_name:
        meta += f": {calendar_name}"
    return blocks, meta + "."


def load(path: Path) -> tuple[list[Block], str]:
    return calendar_blocks(read_text(path), str(path))
