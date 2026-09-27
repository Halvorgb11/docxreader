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
- Python.
- Claude API (Anthropic) som modell for agenten.
- PDF-bibliotek er ikke bestemt ennå.

## Mappestruktur
- `CLAUDE.md` – denne filen, prosjektkontekst for Claude.
- `.claude/skills/` – prosjektspesifikke skills. Hver skill ligger i egen mappe med en `SKILL.md`. Egne skills legges direkte her.
- `.agents/skills/` – skills installert med `npx skills` (felles for flere AI-agenter). Symlenket inn i `.claude/skills/`.
- `.gitignore` – holder `.env`, virtuelle miljøer og cache utenfor git.
- `skills-lock.json` – låsfil for skills installert med `npx skills`.

## Kommandoer
- Repo: https://github.com/Halvorgb11/docxreader (privat, branch `main`). Commit og push med vanlig `git`.
- Node.js er installert via nvm (v24 LTS). I nye skall: `export NVM_DIR="$HOME/.nvm"; . "$NVM_DIR/nvm.sh"` hvis `node` ikke finnes.
- GitHub CLI: `~/.local/bin/gh` (v2.101.0, installert uten Homebrew).
- Legge til skills: `npx skills add <github-repo> --skill '<navn|*>' --yes`

## Konvensjoner
_Ingen ennå._

## Beslutninger
- 2026-09-27: Prosjektet opprettet med `CLAUDE.md` og `.claude/skills/`.
- 2026-09-27: Installerte Node.js via nvm (ingen sudo/Homebrew på maskinen) og alle 23 skills fra `langchain-ai/langchain-skills`.
- 2026-09-27: Initialiserte git og opprettet privat GitHub-repo `Halvorgb11/docxreader`.
