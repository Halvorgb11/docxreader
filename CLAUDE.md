# docxreader

## Om prosjektet
LangChain-verktøy (`@tool`) som lar en Claude-agent lese dokumenter. Støtter nå .docx, .pptx, .xlsx, .csv, .eml, .msg, .ics, .md og .txt, pluss spørringer i regneark/CSV (`query_table`) og PDF/bilder sendt rett til Claude (`view_file`).

**Langsiktig mål (2026-09-28):** repoet skal være en meny/et lager av fillesere for mange filtyper. Ny filtype = ny modul i `readers/` + én linje i `READERS`. Alle filtyper fra toppen av flermålsanalysen (2026-09-28) er nå med; neste kandidater: JSON/XML, HTML (`readers/htmltext.py` finnes allerede). PDF og bilder får ingen tekstleser: `view_file` sender dem som innholdsblokker, og Claude leser/ser dem selv.

Kursendring 2026-09-27: fokus flyttet fra PDF til .docx, fordi Claude API leser PDF direkte, men ikke .docx. `read_pdf` ble fjernet samme dag (finnes i git-historikken, commit `94f61b4`).

## Arbeidsregler for Claude
- **Hold denne filen oppdatert.** Når noe viktig bestemmes eller endres (formål, stack, mappestruktur, kommandoer, konvensjoner, kjente fallgruver), oppdater relevant seksjon her i samme arbeidsøkt.
- Skriv kort og konkret – dette er en referanse, ikke en logg. Fjern det som ikke lenger stemmer.
- Kommuniser med brukeren på norsk.
- **Brukeren er nybegynner i LangChain og vil lære det.** Git og Claude API kan brukeren fra før. Forklar LangChain-konsepter og hvorfor vi gjør som vi gjør, med enkle ord. Innfør ett begrep om gangen. Bygg i små steg brukeren kan følge, og skriv lesbar, godt kommentert kode fremfor smarte snarveier.

## Teknologi / stack
- LangChain (verktøy definert med `@tool`).
- Python 3.12, prosjekt og avhengigheter styres med `uv` (`pyproject.toml` + `uv.lock`).
- `langchain` 1.x (`create_agent`, `@tool` fra `langchain_core.tools`).
- `langchain-anthropic` – Claude API som modell.
- `python-docx` – lesing av .docx (importeres som `import docx`).
- `python-pptx` – lesing av .pptx (importeres som `import pptx`).
- `openpyxl` – lesing av .xlsx.
- `pypdf` – bare til å telle og klippe ut PDF-sider for `view_file` (Claude leser selve PDF-en).
- `Pillow` (via python-pptx) – omgjøring/nedskalering av bilder for `view_file`.
- `extract-msg` – lesing av Outlook .msg (kan bare lese, ikke lage .msg). `.eml` leses med `email` fra standardbiblioteket.
- `python-dotenv` – leser `ANTHROPIC_API_KEY` fra `.env`.
- `pytest` (dev) – tester.

## Mappestruktur
- `CLAUDE.md` – denne filen, prosjektkontekst for Claude.
- `pyproject.toml` / `uv.lock` / `.python-version` – Python-prosjektet.
- `src/docxreader/` – pakken (navnet er historisk; kan døpes om senere).
  - `tools.py` – de fire generelle `@tool`-verktøyene `read_document`, `document_outline`, `read_section`, `search_document`, tabellverktøyet `query_table`, `view_file` (PDF/bilder), og `ALL_TOOLS`. Ren Python, testbar uten Claude.
  - `tables.py` – tabelldata: `Table`, `make_table`, `table_blocks` (tabell → blokker, deling i `ROWS_PER_PART`), `parse_number`, `find_column`, `make_filter`, `query` (logikken bak `query_table`).
  - `blocks.py` – filtype-uavhengig: `Block`, `DocumentError`, `outline`, `section_end`, `find_section`, `heading_paths`, `block_paths`, `search_units`, `render`, `rows_to_markdown`/`table_block`, `heading()`, `count_words`.
  - `readers/__init__.py` – menyene `READERS` (filendelse → `load`) og `TABLE_READERS` (filendelse → `load_tables`), `load_document(path)` og `load_table(path, sheet)`.
  - `readers/word.py` (.docx), `readers/powerpoint.py` (.pptx), `readers/excel.py` (.xlsx), `readers/csvfile.py` (.csv), `readers/mail.py` (.eml via `load_eml`, .msg via `load_msg`), `readers/htmltext.py` (`html_to_text`, HTML → markdown-lignende tekst), `readers/media.py` (`media_block`: PDF/bilde → LangChain-innholdsblokk), `readers/calendar.py` (.ics; `calendar_blocks`), `readers/text.py` (.md via `load_markdown`, .txt via `load_plain`).
  - `agent.py` = Claude + `create_agent` (`build_agent()`, `ask()`), `__init__.py` = CLI (`main`, `print_messages`).
- `tests/` – pytest-tester: `test_readers.py` (meny + felles), `test_word.py`, `test_powerpoint.py`, `test_excel.py`, `test_csv.py`, `test_query_table.py`, `test_mail.py`, `test_media.py`, `test_calendar.py`, `test_text.py`.
- `samples/` – testfiler: `prosjektplan.docx` (lite), `arsrapport.docx` (~8000 ord, plantede fakta), `salgsmote.pptx` (faktum i talenotater: Havbruk Vest AS), `budsjett.xlsx` (kommentar i Oversikt!C3: Nordic Data AS; Transaksjoner rad 180: Fjellsikring AS 1 250 000 – bevisst selvmotsigende, 250 rader, skjult ark), `reiseregning.csv` (cp1252, `;`, norske tall, tom linje, `;` i anførselstegn; mest brukt: Per Æsøy 10 530 kr), `retningslinjer.md`. `tilbud.eml` (tråd i to nivåer med plantet møtetid «8. oktober kl. 10 i rom Fjorden», vedlegg .xlsx/.docx/.png/videresendt e-post med «gjelder til 15. oktober 2026», innebygd logo). `invitasjon.ics` (skrevet for hånd: 3 hendelser – TZID, UTC, heldag 23.–24. nov, brutt linje med parkeringskode 4471, VALARM, CN med kolon). `kvittering.png` (totalt 1 487,50 kr) og `moteinnkalling.pdf` (2 sider, side 2: Ålesund 14. november 2026) – begge uten tekstlag. `tilbud.eml` har bildevedlegget `skisse.png` (Rack B3). `budsjett.xlsx` har også arket `Kvartal` (tittel, to tabeller fra kolonne B, tom rad i tabell, tall som tekst i tre formater, fotnote; faktisk Q1 = 2 525 000,50).
- `scripts/make_sample_docx.py`, `make_large_sample_docx.py`, `make_sample_pptx.py`, `make_sample_xlsx.py`, `make_sample_csv.py`, `make_sample_eml.py`, `make_sample_media.py` – lager testfilene på nytt. `retningslinjer.md` er skrevet for hånd.
- `.env.example` – mal for `.env` (API-nøkkel).
- `.claude/skills/` – prosjektspesifikke skills. Hver skill ligger i egen mappe med en `SKILL.md`. Egne skills legges direkte her.
  - `ny-filleser/` – egen skill: oppskrift for å legge til en ny filtype (`/ny-filleser xlsx`), med maler i `templates/` for leser og test. Hold den i takt med konvensjonene under.
- `.agents/skills/` – skills installert med `npx skills` (felles for flere AI-agenter). Symlenket inn i `.claude/skills/`.
- `.gitignore` – holder `.env`, virtuelle miljøer og cache utenfor git.
- `skills-lock.json` – låsfil for skills installert med `npx skills`.

## Kommandoer
- `uv sync` – installer avhengigheter. `uv add <pakke>` – legg til avhengighet.
- `uv run docxreader "spørsmål"` – kjør agenten (krever `.env`). Skriver ut hvert steg i agentløkken.
- `uv run pytest` – kjør tester. `uv run python ...` – kjør kode i prosjektets miljø.
- Repo: https://github.com/Halvorgb11/docxreader (privat, branch `main`). Commit og push med vanlig `git`.
- Node.js er installert via nvm (v24 LTS). I nye skall: `export NVM_DIR="$HOME/.nvm"; . "$NVM_DIR/nvm.sh"` hvis `node` ikke finnes.
- GitHub CLI: `~/.local/bin/gh` (v2.101.0, installert uten Homebrew).
- Legge til skills: `npx skills add <github-repo> --skill '<navn|*>' --yes`

## Plan og fremdrift
1. [x] Oppsett: uv-prosjekt, avhengigheter.
2. [x] `read_pdf`-verktøy (senere fjernet til fordel for `.docx`).
3. [x] Agent i `agent.py` med `create_agent` + Claude, kjørbar fra kommandolinjen.
4. [x] `read_docx`-verktøy (python-docx): overskrifter, avsnitt og tabeller som markdown-lignende tekst. Test-.docx i `samples/`, tester.
5. [x] Koble `read_docx` til agenten – testet mot Claude (tabell, lister og overskrifter leses riktig).
6. [x] Store filer: `docx_outline` + `read_docx_section` + grense (`MAX_WORDS`) i `read_docx`.
7. [x] Søk: `search_docx`.
8. [x] Flere filtyper: omstrukturert til `blocks.py` + `readers/` + generelle verktøy; lagt til .md/.txt og .pptx.
9. [x] Excel (.xlsx)-leser (via skillen `ny-filleser`).
10. [x] `query_table` (filtrer/sorter/tell/summer, godkjent av brukeren) + CSV-leser. Agenten bruker nå `query_table` i stedet for å lese alle radene.
11. [x] E-post (.eml/.msg) med vedlegg lest av de andre leserne.
12. [ ] Ideer: samtaleminne (checkpointer) + `ContextEditingMiddleware`, liste dokumenter i en mappe, kommentarer/sporede endringer, RAG ved mange dokumenter.

## Konvensjoner
- Verktøy defineres med `@tool(parse_docstring=True)` og Google-stil docstring (`Args:`), slik at argumentbeskrivelsene havner i skjemaet modellen ser.
- Verktøy returnerer feil som tekst som begynner med `Feil:` i stedet for å kaste unntak, så modellen kan forstå og håndtere feilen.
- Verktøy testes med `tool.invoke({...})` i `tests/`, uten å kalle Claude.
- Systemprompten har arbeidsregler for agenten (f.eks. «regn ikke i hodet – bruk query_table»); verktøybeskrivelser sier hva verktøyet kan. Uten regelen regnet Claude små tabeller selv.
- Modell settes i `agent.py` som `MODEL = "anthropic:claude-sonnet-5"` (streng-format `leverandør:modell`).
- .docx har ingen sider; struktur hentes fra avsnittsstiler (`Title`, `Heading N`, `List Bullet`, `List Number`) og gjøres om til markdown.
- Fallgruve: `doc.paragraphs` og `doc.tables` er separate lister. Bruk `doc.iter_inner_content()` for avsnitt og tabeller i riktig rekkefølge.
- **Leser-kontrakt:** `load(path: Path) -> (list[Block], metatekst)`. Kast `DocumentError` ved ødelagt fil. Filsjekk og filtype-valg gjøres i `load_document`, ikke i leseren.
- **Ny filtype – oppskrift** (utførlig i skillen `ny-filleser`): 1) `readers/<format>.py` med `load()`, 2) legg filendelsen i `READERS`, 3) nevn filendelsen i docstringen til ALLE fire verktøy (`test_every_supported_file_type_is_mentioned_in_every_tool` sjekker det), 4) testfil i `samples/` (helst via skript i `scripts/`) + `tests/test_<format>.py`, 5) oppdater CLAUDE.md og README. Passer formatet dårlig i overskrift/seksjon-modellen (f.eks. regneark), lag heller egne verktøy.
- Få generelle verktøy fremfor ett sett per filtype: modellen velger ut fra beskrivelser, og mange nesten like verktøy gir feilvalg.
- `Block`: `markdown`, `words`, `heading_level` (0 = ikke overskrift, 1 = `#`, 2 = `##` …), `is_table`.
- Word: topp-/bunntekst i metateksten, `Title` → `#`, `Heading N` → N+1 `#`, lister → `-`/`1.`, tabeller → markdown-tabeller (første rad = overskrift).
- PowerPoint: hvert lysbilde = `## Lysbilde N: <tittel>` (nivå 2). Former sorteres etter plassering (topp, venstre), grupper leses rekursivt, punkter får innrykk etter `p.level`, diagrammer blir `[Diagram: tittel]`, bilder hoppes over, talenotater blir `Talenotater: …`. Metatekst: tittel + antall lysbilder.
- `split_tables` (i `tables.py`) finner tabellene i et ark: tomme rader skiller områder; rader med én utfylt celle over en ren tekst-overskriftsrad er titler (står det data rett under, er den korte raden selve overskriftsraden); et område uten tittel som starter med data (tall/dato) er fortsettelsen av forrige tabell; tekst etter siste tabell blir fotnote (`Table.notes`); tomme kolonner i kantene fjernes. Ark med flere tabeller: `## Ark: X` + `### <tittel>` per tabell, tabellnavn `"X – <tittel>"` (eller `"X – tabell N"`). Ett-kolonne-ark: ingen tittelgjenkjenning. CSV: `split_on_gaps=False` (bare tittellinje øverst gjenkjennes). Kommentarer havner hos tabellen som dekker raden.
- `view_file(path, attachment)` returnerer en LISTE med innholdsblokker: `{"type": "text", …}` + `{"type": "image", "base64", "mime_type"}` eller `{"type": "file", "base64", "mime_type": "application/pdf"}` (LangChain-standardblokker; langchain-anthropic gjør dem om til Claudes image/document i tool_result). Bilder: PNG/JPEG/GIF/WebP sendes uendret hvis ≤ 1568 px og ≤ 3,5 MB; ellers (og TIFF/BMP) gjøres om med Pillow (JPEG uten gjennomsiktighet, ellers PNG) og skaleres ned; flersidig TIFF → side 1. PDF: må starte med `%PDF-`, sider telles med `pypdf`, maks `MAX_PDF_PAGES = 20` sider per kall (tokenkostnad), større PDF-er leses i deler med `pages` (`"1-20"`, `"21-"`, `"1,4,7-9"`; `parse_pages`) – pypdf lager en ny PDF med bare de sidene; passordbeskyttet (ikke tomt passord) → `Feil:`; maks 30 MB. Flersidig TIFF: `pages` velger én side. Feil → `Feil:`-tekst (streng). `read_document` på PDF/bilde og bildevedlegg i e-post peker til `view_file`. CLI-en (`_preview`) viser `[image: image/png, ~27 KB]` i stedet for base64.
- Tabellformater har både `load()` (blokker) og `load_tables()` (`Table`-er); `load()` bygges fra `load_tables()` + `table_blocks()`. Metateksten viser kolonnenavn (så modellen kan bruke `query_table` direkte).
- `query_table(path, sheet, attachment, calculate: list[str], where: list[str], sort_by, descending, columns: list[str], limit, group_by, aggregate: Literal[count|sum|avg|min|max], value_column)`. `attachment`: regneark/CSV-vedlegg i en e-post (`"x.eml > data.csv"` for videresendt e-post; ved tvetydig navn foretrekkes tabellfiler) – `readers.load_table` henter det via `mail.load_mail` + `mail.find_attachment`. `calculate`: `"Navn = uttrykk"` med kolonner/tall, `+ - * /`, fortegn og parenteser (egen parser i `tables.py`: kolonnenavn → plassholdere lengste først, så tokens → tre → regnes per rad; IKKE `eval`; operatorer etter delvise kolonnenavn trenger mellomrom foran, så `Pris-per-stk` ikke blir minus), regnes før where/sort/aggregate; tom/ugyldig verdi eller deling på 0 → tom celle. Betingelser `"Kolonne op verdi"`, op = `= != > < >= <= ~` (~ = inneholder). Et element i `where` kan ha alternativer med ` OR ` (ELLER); flere elementer = OG. Tall sammenlignes som tall (`parse_number`: norsk `1 250,50`, engelsk `1,250.50`, `1.250.000,50`, `kr`/`NOK`/`$`/`€`/`%`, `,-`, typografisk minus; finnes både `,` og `.`, er det siste desimaltegnet; samme tegn flere ganger = tusenskille; ETT enkelt `,`/`.` = desimaltegn, så `1,250` = 1,25), ellers tekst uten hensyn til store/små bokstaver (ISO-datoer fungerer). Kolonner og ark: eksakt, så entydig delvis treff; `"Ark: "` foran arknavnet er lov; én tabell i filen → `sheet` ignoreres. Sortering: tomme celler sist. Grupper sorteres med største verdi først. Tomme celler hoppes over i summer; ikke-tall rapporteres. `MAX_QUERY_ROWS = 100`.
- CSV: én seksjon `## Tabell`; skilletegn gjettes fra første linje (`; , tab |`, flest vinner); tekst via `text.read_text` (UTF-8, ellers cp1252); `Rad` = linjenummer; tomme linjer hoppes over; korte rader fylles ut, tomme overskrifter → `Kolonne N`.
- Excel: hvert ark = `## Ark: <navn>` (+ ` (skjult)`), innhold som tabell med ekstra første kolonne `Rad` (Excel-radnummer). Første ikke-tomme rad = overskriftsrad. Over `ROWS_PER_PART = 100` datarader deles arket i `### Rad 2–101` osv., med overskriftsraden gjentatt. Formler: lastes både med `data_only=True` (lagret resultat) og uten; mangler resultat (fil aldri åpnet i Excel) vises formelen. Datoer uten klokkeslett → `2026-09-28`, heltall-float → heltall. Tomme rader og tomme kolonner til høyre fjernes; tomme ark nevnes bare i metateksten. Cellekommentarer → `Kommentarer:\n- C3: …`. Kjente begrensninger: tabeller side om side (horisontalt) i samme ark skilles ikke; en ny tabell uten tittel der overskriftsraden inneholder tall (f.eks. årstall 2025/2026) tolkes som fortsettelse av forrige tabell; diagrammer hoppes over; `MAX_WORDS` undervurderer token-mengden i talltabeller; `25 %` i tekst blir 25 (Excel-prosent lagret som tall blir 0.25).
- E-post: `.eml` og `.msg` gjøres om til `Mail` (mellomform) → `_mail_blocks`. Metatekst: `E-post` + Fra/Til/Kopi/Dato/Emne/Vedlegg. Seksjoner (nivå 2): `Melding`, `Tidligere melding N` (sitert tråd, delt ved «Den/On … skrev/wrote …:», «-----Original Message-----/Opprinnelig melding/Forwarded message», eller Outlook-hode «Fra:» + «Sendt:/Dato:» innen 3 linjer; ett nivå `>` fjernes; siterte hodelinjer blir liste), `Vedlegg: navn`. Ren tekst foretrekkes framfor HTML. Vedlegg leses med `READERS` via midlertidig fil (bare filnavnet brukes, aldri stien), overskriftene flyttes så vedleggets øverste nivå blir 3 (`_nest`), vedleggets metatekst står først. Videresendt e-post (`message/rfc822` / `.msg`-type MSG) leses rekursivt, maks `MAX_DEPTH = 3`, navn = emnet. Innebygde bilder (Content-ID) er ikke vedlegg. `.msg`: alle felt leses med `_field` (unntak → tomt), tekst renses for `\x00` (`_clean`), Til/Kopi hentes fra `msg.recipients` (fallback `msg.to`/`msg.cc`). Uleselige vedlegg får en merknad i parentes, aldri `Feil:`. Maks 20 MB per vedlegg.
- Kalender (.ics): én seksjon `## Hendelse: <SUMMARY>` per VEVENT med liste (Start, Slutt, Sted, Status, Gjentas = rå RRULE, Arrangør, Deltakere med svarstatus) + beskrivelse. Linjer foldes sammen, `\n \, \;` tolkes, kolon i anførselstegn i parametere ignoreres, VALARM hoppes over. Tid: TZID → `2026-10-08 10:00 (Europe/Oslo)`, `Z` → `(UTC)`, `VALUE=DATE` → hele dagen med eksklusiv DTEND (minus én dag). Metatekst: antall hendelser, METHOD (invitasjon/avlysning …), X-WR-CALNAME.
- Møteinnkallinger i e-post: `text/calendar`-deler inni meldingen legges til som vedlegget `invitasjon.ics`. Outlook-avtaler (.msg) får Start/Slutt/Sted/Arrangør i metateksten, som da starter med `Møteinnkalling (Outlook)`.
- Markdown: `#`-overskrifter starter alltid ny blokk (også uten tom linje foran); `#` i kodeblokker er ikke overskrifter; avsnitt over flere linjer slås sammen. `.txt` har ingen overskrifter. Tekst leses som UTF-8 (med BOM), ellers cp1252.
- Tittel i overskriftsstier: finnes det nøyaktig ÉN nivå 1-overskrift, regnes den som tittel og utelates fra stiene; ellers er `#`-overskrifter med.
- Seksjon = overskrift + alt fram til neste overskrift på samme eller høyere nivå (`_section_end`).
- Passer flere overskrifter, men bare én ligger øverst (kortest sti), velges den (`"Melding"` = e-postens, ikke vedleggets).
- Søket deler avsnitt i setninger med `split_sentences`: ikke etter forkortelser (`ABBREVIATIONS`: kl., f.eks., nr. …) eller når neste ord starter med liten bokstav/tall (`8. oktober`, `kl. 10`).
- `read_section` tar overskrift eller sti med ` > ` (f.eks. `"Økonomi > Status"`); eksakt treff (alle deler) før delvis treff (alle deler, også foreldre: `"Transaksjoner > Rad 2–101"` finner `Ark: Transaksjoner > …`), uavhengig av store/små bokstaver. Tvetydig eller ukjent overskrift → `Feil:` med liste over alternativer.
- `MAX_WORDS = 3000` (i `tools.py`): over dette gir `read_document` innholdsfortegnelsen, og `read_section` gir underoverskriftene. Tester endrer den med `monkeypatch`.
- `search_document`: alle søkeord (delstrenger, uavhengig av store/små bokstaver) må finnes i samme bit – setning (avsnitt), listepunkt eller tabellrad (`search_units`). Treff grupperes under `[overskriftssti]` (`block_paths`); tabelltreff vises med overskriftsraden. Maks `MAX_SEARCH_HITS = 15`.
- Interne feil kastes som `DocumentError` og gjøres om til `Feil: …` i verktøyene. Sikkerhetsnett: `readers._run` gjør ethvert uventet unntak fra en leser om til `DocumentError` – et verktøy skal aldri krasje agenten.
- Ødelagt .docx/.pptx gir `PackageNotFoundError` (fra hhv. `docx.opc.exceptions` og `pptx.exc`); fanges i leseren og gjøres om til `DocumentError`.

## Beslutninger
- 2026-09-27: Prosjektet opprettet med `CLAUDE.md` og `.claude/skills/`.
- 2026-09-27: Installerte Node.js via nvm (ingen sudo/Homebrew på maskinen) og alle 23 skills fra `langchain-ai/langchain-skills`.
- 2026-09-27: Initialiserte git og opprettet privat GitHub-repo `Halvorgb11/docxreader`.
- 2026-09-27: Fjernet `read_pdf`, `pypdf` og `samples/rapport.pdf`.
- 2026-09-28: Flermålsanalyse av filtyper (relevans 40 %, enkelhet 30 %, utvidbarhet 30 %): xlsx 4,4 > pptx 3,7 > csv 3,6 > md/txt 3,3 = e-post 3,3 > json/xml 2,9 > html 2,6 > odt 2,3 > epub/rtf 1,9.
- 2026-09-28: Laget skillen `.claude/skills/ny-filleser` for å legge til filtyper på en ensartet måte.
- 2026-09-28: `pypdf` tilbake (bare sidetelling/utklipp), `view_file(pages=…)`; `calculate` med parenteser.
- 2026-09-28: Lagt til .ics-leser og kalenderfelt/invitasjoner i e-post; regel i systemprompten om å bruke query_table til utregninger.
- 2026-09-28: Lagt til `view_file` (PDF/bilder som innholdsblokker) – verifisert mot Claude: leste kvittering, side 2 i PDF og bildevedlegg uten tekstlag.
- 2026-09-28: Lagt til e-post (.eml/.msg) med vedlegg via menyen.
- 2026-09-28: `.msg` verifisert manuelt mot extract-msg sine eksempelfiler (GitHub TeamMsgExtractor/msg-extractor, `example-msg-files/`; GPL, derfor ikke i repoet). Fant og rettet: krasj ved ødelagt tegnkoding, `\x00`-fyll, ufullstendig `msg.to`. Regresjonstester med falske objekter.
- 2026-09-28: Fikset begrensninger: tabeller/titler/fotnoter i ark (`split_tables`), flere tallformater, ` OR ` i `where`.
- 2026-09-28: Lagt til `query_table` (ett verktøy med valgfrie argumenter i stedet for flere små) og CSV-leser.
- 2026-09-28: Lagt til Excel-leser (.xlsx) med skillen `ny-filleser`. Delvis treff tillatt også i foreldre-deler av overskriftsstier.
- 2026-09-28: Omstrukturert til `blocks.py` + `readers/`; `read_docx`/`docx_outline`/`read_docx_section`/`search_docx` erstattet av generelle `read_document`/`document_outline`/`read_section`/`search_document`. Lagt til .md, .txt og .pptx.
