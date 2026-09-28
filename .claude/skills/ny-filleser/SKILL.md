---
name: ny-filleser
description: Legg til støtte for en ny filtype (f.eks. .xlsx, .csv, .eml, .html) i docxreader-prosjektets leser-meny. Bruk når brukeren vil at agenten skal kunne lese en ny filtype, eller ber om en ny "leser"/"filleser".
argument-hint: "<filendelse, f.eks. xlsx>"
---

# Ny filleser: $ARGUMENTS

Repoet er et lager av fillesere. Hver filtype er én modul i `src/docxreader/readers/`
som gjør filen om til `Block`-er. De fire verktøyene i `tools.py` (`read_document`,
`document_outline`, `read_section`, `search_document`) virker da automatisk.

Brukeren er nybegynner i LangChain og vil lære: forklar kort hva du gjør og hvorfor
underveis, på norsk. Git og Claude API trenger ikke forklares.

## 0. Vurder formatet før du koder

Les `CLAUDE.md` (konvensjoner + flermålsanalysen) og svar kort for brukeren:

- **Passer formatet i overskrift/seksjon-modellen?** Word, PowerPoint, Markdown og HTML gjør det
  (overskrift/lysbilde = seksjon). Regneark, CSV og JSON gjør det dårlig.
  - Passer det: bare en leser (steg 1–5). Ingen nye verktøy.
  - Passer det dårlig: leser for grunnleggende lesing/søk **pluss** egne verktøy
    (f.eks. `list_sheets`, `read_range`). Stopp og avklar verktøydesignet med brukeren
    før du lager nye verktøy – få generelle verktøy er bedre enn mange like.
- **Hva blir en seksjon?** (Word: overskrift. PowerPoint: `## Lysbilde N: tittel`.) Hva går i metateksten?
- **Hva kan Claude allerede lese selv?** PDF og bilder leser Claude API direkte – ikke lag lesere for dem.

## 1. Bibliotek

- Foretrekk standardbiblioteket. Ellers ett godt vedlikeholdt bibliotek: `uv add <pakke>`.
- **Inspiser API-et før du skriver koden** i stedet for å gjette:
  `uv run python -c "import pakke; help(pakke.X)"` – særlig navnet på unntaket for ødelagte filer.

## 2. Leseren – `src/docxreader/readers/<format>.py`

Start fra [templates/reader.py](templates/reader.py). Kontrakt:

```python
def load(path: Path) -> tuple[list[Block], str]:  # (blokker, metatekst)
```

- Bruk hjelperne i `docxreader.blocks`: `heading(level, text)`, `table_block(rows)`,
  `Block(text, count_words(text))`. Ikke lag egen markdown for tabeller.
- `heading_level`: 1 = `#` (tittel), 2 = `##` (kapittel/lysbilde), 3 = `###` …
- Lister: én linje per punkt i SAMME blokk (`"- a\n- b"`) – søket deler på linjeskift.
  Avsnitt: én linje (slå sammen linjeskift) – søket deler på setninger.
- Ødelagt fil: fang bibliotekets unntak og kast `DocumentError("… er ikke en gyldig …-fil …")`.
- Ikke sjekk om filen finnes eller filendelsen – det gjør `load_document`.
- Kommenter på norsk, lesbart for en nybegynner.

Fallgruver vi har truffet:
- **Rekkefølge:** biblioteker gir ofte elementer i lagret rekkefølge, ikke lese-rekkefølge
  (Word: `iter_inner_content()`; PowerPoint: sorter former på `top`, `left`).
- **Nøstede elementer** (grupper, tabeller i tabeller) må leses rekursivt.
- **Tegnkoding** for tekstformater: `utf-8-sig`, så `cp1252` (se `readers/text.py`).
- **Skjult innhold** som ofte har svaret: talenotater, topp-/bunntekst, kommentarer, alt-tekst.
- Tittel i overskriftsstier: nøyaktig ÉN nivå 1-overskrift regnes som tittel.
- Test stier slik modellen sannsynligvis skriver dem, også uten prefiks (`"Transaksjoner > Rad 2–101"` for `Ark: Transaksjoner`).
- Talltabeller er token-tette: `MAX_WORDS` teller ord, ikke tokens. Del store tabeller i biter (se `readers/excel.py`).
- Se på verktøykallene i agentkjøringen: leser Claude alt for å svare på noe som egentlig er en spørring (summer, største, filtrer), trengs det kanskje et eget verktøy.

## 3. Registrer i menyen

- Legg filendelsen(e) i `READERS` i `src/docxreader/readers/__init__.py`.
- Nevn filendelsen i **"Støttede filtyper"-linjen i docstringen til alle fire verktøy** i `tools.py`.
  Modellen vet bare det som står der. `test_every_supported_file_type_is_mentioned_in_every_tool`
  feiler hvis du glemmer det.
- Oppdater `SYSTEM_PROMPT` i `agent.py` bare hvis formatet har et eget begrep for "seksjon".

## 4. Testfil og tester

- Testfil i `samples/`, laget av et skript `scripts/make_sample_<format>.py`
  (se `make_sample_pptx.py`) så den kan lages på nytt. Skriv filen for hånd bare for rene tekstformater.
- Legg inn minst ett **plantet faktum** på et sted som er lett å overse (notater, fotnote, gruppe),
  og elementer i "feil" lagret rekkefølge.
- `tests/test_<format>.py` fra [templates/test_reader.py](templates/test_reader.py). Dekk minst:
  seksjoner/overskrifter, tabeller, lister, rekkefølge, plantet faktum via `search_document`,
  `read_section`, `document_outline`, ødelagt fil gir `Feil:`, tom fil.
- `uv run pytest -q` – ALLE tester skal bestå, også de gamle.

## 5. Prøv med agenten

Kjør ett spørsmål der svaret bare finnes via det plantede faktumet, f.eks.
`uv run docxreader "Hva står det om X i samples/<fil>?"`. Vis brukeren hvilke verktøy
Claude valgte og hvorfor (beskrivelsene styrer valget). Kjøringen koster litt API-bruk.

## 6. Dokumenter og lever

- `CLAUDE.md`: stack (nytt bibliotek), mappestruktur (leser, testfil, skript), konvensjoner
  (hvordan formatet blir til blokker), plan og en datert beslutning.
- `README.md`: listen over støttede filtyper.
- Commit med beskrivende melding og push (`git push`).
- Oppsummer for brukeren: hva formatet blir til, hva som hoppes over, begrensninger,
  og hva som eventuelt trengs av egne verktøy senere.
