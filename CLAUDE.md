# docxreader

## Om prosjektet
LangChain-verktøy (`@tool`) som lar en Claude-agent lese dokumenter. Støtter nå .docx, .pptx, .md og .txt.

**Langsiktig mål (2026-09-28):** repoet skal være en meny/et lager av fillesere for mange filtyper. Ny filtype = ny modul i `readers/` + én linje i `READERS`. Prioritert rekkefølge videre (fra flermålsanalyse 2026-09-28): Excel (.xlsx) + CSV med egne tabellverktøy, deretter e-post (.eml/.msg). PDF og bilder utelates fordi Claude API leser dem direkte.

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
- `python-dotenv` – leser `ANTHROPIC_API_KEY` fra `.env`.
- `pytest` (dev) – tester.

## Mappestruktur
- `CLAUDE.md` – denne filen, prosjektkontekst for Claude.
- `pyproject.toml` / `uv.lock` / `.python-version` – Python-prosjektet.
- `src/docxreader/` – pakken (navnet er historisk; kan døpes om senere).
  - `tools.py` – de fire generelle `@tool`-verktøyene `read_document`, `document_outline`, `read_section`, `search_document` + `ALL_TOOLS`. Ren Python, testbar uten Claude.
  - `blocks.py` – filtype-uavhengig: `Block`, `DocumentError`, `outline`, `section_end`, `find_section`, `heading_paths`, `block_paths`, `search_units`, `render`, `rows_to_markdown`/`table_block`, `heading()`, `count_words`.
  - `readers/__init__.py` – menyen `READERS` (filendelse → `load`-funksjon) og `load_document(path)`.
  - `readers/word.py` (.docx), `readers/powerpoint.py` (.pptx), `readers/text.py` (.md via `load_markdown`, .txt via `load_plain`).
  - `agent.py` = Claude + `create_agent` (`build_agent()`, `ask()`), `__init__.py` = CLI (`main`, `print_messages`).
- `tests/` – pytest-tester: `test_readers.py` (meny + felles), `test_word.py`, `test_powerpoint.py`, `test_text.py`.
- `samples/` – testfiler: `prosjektplan.docx` (lite), `arsrapport.docx` (~8000 ord, plantede fakta), `salgsmote.pptx` (faktum i talenotater: Havbruk Vest AS), `retningslinjer.md`.
- `scripts/make_sample_docx.py`, `make_large_sample_docx.py`, `make_sample_pptx.py` – lager testfilene på nytt. `retningslinjer.md` er skrevet for hånd.
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
9. [ ] Neste filtyper: Excel (.xlsx) + CSV (egne tabellverktøy: ark, områder, filtrering), deretter e-post.
10. [ ] Ideer: samtaleminne (checkpointer) + `ContextEditingMiddleware`, liste dokumenter i en mappe, kommentarer/sporede endringer, RAG ved mange dokumenter.

## Konvensjoner
- Verktøy defineres med `@tool(parse_docstring=True)` og Google-stil docstring (`Args:`), slik at argumentbeskrivelsene havner i skjemaet modellen ser.
- Verktøy returnerer feil som tekst som begynner med `Feil:` i stedet for å kaste unntak, så modellen kan forstå og håndtere feilen.
- Verktøy testes med `tool.invoke({...})` i `tests/`, uten å kalle Claude.
- Modell settes i `agent.py` som `MODEL = "anthropic:claude-sonnet-5"` (streng-format `leverandør:modell`).
- .docx har ingen sider; struktur hentes fra avsnittsstiler (`Title`, `Heading N`, `List Bullet`, `List Number`) og gjøres om til markdown.
- Fallgruve: `doc.paragraphs` og `doc.tables` er separate lister. Bruk `doc.iter_inner_content()` for avsnitt og tabeller i riktig rekkefølge.
- **Leser-kontrakt:** `load(path: Path) -> (list[Block], metatekst)`. Kast `DocumentError` ved ødelagt fil. Filsjekk og filtype-valg gjøres i `load_document`, ikke i leseren.
- **Ny filtype – oppskrift** (utførlig i skillen `ny-filleser`): 1) `readers/<format>.py` med `load()`, 2) legg filendelsen i `READERS`, 3) nevn filendelsen i docstringen til ALLE fire verktøy (`test_every_supported_file_type_is_mentioned_in_every_tool` sjekker det), 4) testfil i `samples/` (helst via skript i `scripts/`) + `tests/test_<format>.py`, 5) oppdater CLAUDE.md og README. Passer formatet dårlig i overskrift/seksjon-modellen (f.eks. regneark), lag heller egne verktøy.
- Få generelle verktøy fremfor ett sett per filtype: modellen velger ut fra beskrivelser, og mange nesten like verktøy gir feilvalg.
- `Block`: `markdown`, `words`, `heading_level` (0 = ikke overskrift, 1 = `#`, 2 = `##` …), `is_table`.
- Word: topp-/bunntekst i metateksten, `Title` → `#`, `Heading N` → N+1 `#`, lister → `-`/`1.`, tabeller → markdown-tabeller (første rad = overskrift).
- PowerPoint: hvert lysbilde = `## Lysbilde N: <tittel>` (nivå 2). Former sorteres etter plassering (topp, venstre), grupper leses rekursivt, punkter får innrykk etter `p.level`, diagrammer blir `[Diagram: tittel]`, bilder hoppes over, talenotater blir `Talenotater: …`. Metatekst: tittel + antall lysbilder.
- Markdown: `#`-overskrifter starter alltid ny blokk (også uten tom linje foran); `#` i kodeblokker er ikke overskrifter; avsnitt over flere linjer slås sammen. `.txt` har ingen overskrifter. Tekst leses som UTF-8 (med BOM), ellers cp1252.
- Tittel i overskriftsstier: finnes det nøyaktig ÉN nivå 1-overskrift, regnes den som tittel og utelates fra stiene; ellers er `#`-overskrifter med.
- Seksjon = overskrift + alt fram til neste overskrift på samme eller høyere nivå (`_section_end`).
- `read_section` tar overskrift eller sti med ` > ` (f.eks. `"Økonomi > Status"`); eksakt treff før delvis, uavhengig av store/små bokstaver. Tvetydig eller ukjent overskrift → `Feil:` med liste over alternativer.
- `MAX_WORDS = 3000` (i `tools.py`): over dette gir `read_document` innholdsfortegnelsen, og `read_section` gir underoverskriftene. Tester endrer den med `monkeypatch`.
- `search_document`: alle søkeord (delstrenger, uavhengig av store/små bokstaver) må finnes i samme bit – setning (avsnitt), listepunkt eller tabellrad (`search_units`). Treff grupperes under `[overskriftssti]` (`block_paths`); tabelltreff vises med overskriftsraden. Maks `MAX_SEARCH_HITS = 15`.
- Interne feil kastes som `DocumentError` og gjøres om til `Feil: …` i verktøyene.
- Ødelagt .docx/.pptx gir `PackageNotFoundError` (fra hhv. `docx.opc.exceptions` og `pptx.exc`); fanges i leseren og gjøres om til `DocumentError`.

## Beslutninger
- 2026-09-27: Prosjektet opprettet med `CLAUDE.md` og `.claude/skills/`.
- 2026-09-27: Installerte Node.js via nvm (ingen sudo/Homebrew på maskinen) og alle 23 skills fra `langchain-ai/langchain-skills`.
- 2026-09-27: Initialiserte git og opprettet privat GitHub-repo `Halvorgb11/docxreader`.
- 2026-09-27: Fjernet `read_pdf`, `pypdf` og `samples/rapport.pdf`.
- 2026-09-28: Flermålsanalyse av filtyper (relevans 40 %, enkelhet 30 %, utvidbarhet 30 %): xlsx 4,4 > pptx 3,7 > csv 3,6 > md/txt 3,3 = e-post 3,3 > json/xml 2,9 > html 2,6 > odt 2,3 > epub/rtf 1,9.
- 2026-09-28: Laget skillen `.claude/skills/ny-filleser` for å legge til filtyper på en ensartet måte.
- 2026-09-28: Omstrukturert til `blocks.py` + `readers/`; `read_docx`/`docx_outline`/`read_docx_section`/`search_docx` erstattet av generelle `read_document`/`document_outline`/`read_section`/`search_document`. Lagt til .md, .txt og .pptx.
