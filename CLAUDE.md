# docxreader

## Om prosjektet
Et LangChain-verktøy (`@tool`) som lar en agent lese filer. Første mål er at agenten skal kunne lese PDF-filer. Flere filtyper, for eksempel .docx, kan komme senere. Detaljer fylles inn etter hvert.

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
- `pypdf` – lesing av PDF (ren Python).
- `python-dotenv` – leser `ANTHROPIC_API_KEY` fra `.env`.
- `pytest` (dev) – tester.

## Mappestruktur
- `CLAUDE.md` – denne filen, prosjektkontekst for Claude.
- `pyproject.toml` / `uv.lock` / `.python-version` – Python-prosjektet.
- `src/docxreader/` – pakken. `tools.py` = `@tool`-funksjoner (ren Python, testbar uten Claude), `agent.py` = Claude + `create_agent`.
- `tests/` – pytest-tester. `samples/` – test-PDF-er.
- `.env.example` – mal for `.env` (API-nøkkel).
- `.claude/skills/` – prosjektspesifikke skills. Hver skill ligger i egen mappe med en `SKILL.md`. Egne skills legges direkte her.
- `.agents/skills/` – skills installert med `npx skills` (felles for flere AI-agenter). Symlenket inn i `.claude/skills/`.
- `.gitignore` – holder `.env`, virtuelle miljøer og cache utenfor git.
- `skills-lock.json` – låsfil for skills installert med `npx skills`.

## Kommandoer
- `uv sync` – installer avhengigheter. `uv add <pakke>` – legg til avhengighet.
- `uv run pytest` – kjør tester. `uv run python ...` – kjør kode i prosjektets miljø.
- Repo: https://github.com/Halvorgb11/docxreader (privat, branch `main`). Commit og push med vanlig `git`.
- Node.js er installert via nvm (v24 LTS). I nye skall: `export NVM_DIR="$HOME/.nvm"; . "$NVM_DIR/nvm.sh"` hvis `node` ikke finnes.
- GitHub CLI: `~/.local/bin/gh` (v2.101.0, installert uten Homebrew).
- Legge til skills: `npx skills add <github-repo> --skill '<navn|*>' --yes`

## Plan og fremdrift
1. [x] Oppsett: uv-prosjekt, avhengigheter.
2. [x] `read_pdf`-verktøy i `tools.py`, testet uten agent.
3. [ ] Agent i `agent.py` med `create_agent` + Claude, kjørbar fra kommandolinjen.
4. [ ] Forbedringer: sidevalg, store PDF-er, gode feilmeldinger til modellen, senere `.docx`.

## Konvensjoner
- Verktøy defineres med `@tool(parse_docstring=True)` og Google-stil docstring (`Args:`), slik at argumentbeskrivelsene havner i skjemaet modellen ser.
- Verktøy returnerer feil som tekst som begynner med `Feil:` i stedet for å kaste unntak, så modellen kan forstå og håndtere feilen.
- `read_pdf` markerer sider som `--- Side N av M ---`, så modellen kan vise til sidenummer.
- Verktøy testes med `tool.invoke({...})` i `tests/`, uten å kalle Claude.
- `samples/rapport.pdf` er en 2-siders test-PDF laget med macOS `cupsfilter`.

## Beslutninger
- 2026-09-27: Prosjektet opprettet med `CLAUDE.md` og `.claude/skills/`.
- 2026-09-27: Installerte Node.js via nvm (ingen sudo/Homebrew på maskinen) og alle 23 skills fra `langchain-ai/langchain-skills`.
- 2026-09-27: Initialiserte git og opprettet privat GitHub-repo `Halvorgb11/docxreader`.
