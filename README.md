# docxreader

LangChain-verktøy (`@tool`) som lar en Claude-agent lese dokumenter: Word (.docx), PowerPoint (.pptx), Markdown (.md) og tekst (.txt).

Målet er et lager av fillesere: hver filtype er én modul i `src/docxreader/readers/`, og alle deler de samme fire verktøyene (`read_document`, `document_outline`, `read_section`, `search_document`).

## Oppsett

```bash
uv sync                 # installer avhengigheter i .venv
cp .env.example .env    # legg inn ANTHROPIC_API_KEY
uv run docxreader "Hva står det i samples/prosjektplan.docx?"
uv run pytest           # tester (uten Claude)
```
