# docxreader

LangChain-verktøy (`@tool`) som lar en Claude-agent lese dokumenter: Word (.docx), PowerPoint (.pptx), Excel (.xlsx), CSV (.csv), e-post (.eml, .msg – med vedlegg), Markdown (.md) og tekst (.txt).

Målet er et lager av fillesere: hver filtype er én modul i `src/docxreader/readers/`, og alle deler de samme fire verktøyene (`read_document`, `document_outline`, `read_section`, `search_document`). Regneark og CSV har i tillegg `query_table` for å filtrere, sortere, telle, summere og regne ut rader (også i e-postvedlegg), og `view_file` sender PDF-er og bilder rett til Claude.

## Oppsett

```bash
uv sync                 # installer avhengigheter i .venv
cp .env.example .env    # legg inn ANTHROPIC_API_KEY
uv run docxreader "Hva står det i samples/prosjektplan.docx?"
uv run pytest           # tester (uten Claude)
```
