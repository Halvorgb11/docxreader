# docxreader

## Om prosjektet
LangChain-verktøy (`@tool`) som lar en Claude-agent lese Word-dokumenter (.docx). Kursendring 2026-09-27: fokus flyttet fra PDF til .docx, fordi Claude API leser PDF direkte, men ikke .docx. `read_pdf` ble fjernet samme dag (finnes i git-historikken, commit `94f61b4`).

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
- `python-dotenv` – leser `ANTHROPIC_API_KEY` fra `.env`.
- `pytest` (dev) – tester.

## Mappestruktur
- `CLAUDE.md` – denne filen, prosjektkontekst for Claude.
- `pyproject.toml` / `uv.lock` / `.python-version` – Python-prosjektet.
- `src/docxreader/` – pakken. `tools.py` = `@tool`-funksjoner (ren Python, testbar uten Claude), `agent.py` = Claude + `create_agent` (`build_agent()`, `ask()`), `__init__.py` = CLI (`main`, `print_messages`).
- `tests/` – pytest-tester. `samples/` – testfiler (`prosjektplan.docx`).
- `scripts/make_sample_docx.py` – lager `samples/prosjektplan.docx` på nytt.
- `.env.example` – mal for `.env` (API-nøkkel).
- `.claude/skills/` – prosjektspesifikke skills. Hver skill ligger i egen mappe med en `SKILL.md`. Egne skills legges direkte her.
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
6. [ ] Senere: lese én seksjon, søk i dokument, liste dokumenter i en mappe, kommentarer/sporede endringer.

## Konvensjoner
- Verktøy defineres med `@tool(parse_docstring=True)` og Google-stil docstring (`Args:`), slik at argumentbeskrivelsene havner i skjemaet modellen ser.
- Verktøy returnerer feil som tekst som begynner med `Feil:` i stedet for å kaste unntak, så modellen kan forstå og håndtere feilen.
- Verktøy testes med `tool.invoke({...})` i `tests/`, uten å kalle Claude.
- Modell settes i `agent.py` som `MODEL = "anthropic:claude-sonnet-5"` (streng-format `leverandør:modell`).
- .docx har ingen sider; struktur hentes fra avsnittsstiler (`Title`, `Heading N`, `List Bullet`, `List Number`) og gjøres om til markdown.
- Fallgruve: `doc.paragraphs` og `doc.tables` er separate lister. Bruk `doc.iter_inner_content()` for avsnitt og tabeller i riktig rekkefølge.
- `read_docx`-format: topp-/bunntekst øverst, `Title` → `#`, `Heading N` → N+1 `#`, lister → `-`/`1.`, tabeller → markdown-tabeller (første rad = overskrift). Hjelpefunksjoner `_paragraph_to_markdown`, `_table_to_markdown`.
- Ødelagt .docx gir `PackageNotFoundError` fra python-docx; fanges og gjøres om til `Feil:`-tekst.

## Beslutninger
- 2026-09-27: Prosjektet opprettet med `CLAUDE.md` og `.claude/skills/`.
- 2026-09-27: Installerte Node.js via nvm (ingen sudo/Homebrew på maskinen) og alle 23 skills fra `langchain-ai/langchain-skills`.
- 2026-09-27: Initialiserte git og opprettet privat GitHub-repo `Halvorgb11/docxreader`.
- 2026-09-27: Fjernet `read_pdf`, `pypdf` og `samples/rapport.pdf`; prosjektet handler bare om .docx.
