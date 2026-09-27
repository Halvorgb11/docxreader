# docxreader

LangChain-verktøy (`@tool`) som lar en Claude-agent lese Word-dokumenter (.docx).

## Oppsett

```bash
uv sync                 # installer avhengigheter i .venv
cp .env.example .env    # legg inn ANTHROPIC_API_KEY
uv run docxreader "Hva står det i samples/prosjektplan.docx?"
```
