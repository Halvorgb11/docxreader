"""docxreader: la en Claude-agent lese dokumenter (.docx, .pptx, .md, .txt) via LangChain-verktøy.

Kjør fra kommandolinjen:
    uv run docxreader "Hva står det i samples/prosjektplan.docx?"
"""

import sys

from langchain_core.messages import AIMessage, HumanMessage, ToolMessage


def _preview(content) -> str:
    """Kort forhåndsvisning av et verktøysvar. Svaret er vanligvis tekst, men
    view_file gir en liste med blokker (tekst + bilde/PDF i base64) – da viser
    vi bare hva slags blokker det er, ikke tusenvis av base64-tegn."""
    if isinstance(content, list):
        parts = []
        for block in content:
            if isinstance(block, dict) and block.get("type") == "text":
                parts.append(block["text"])
            elif isinstance(block, dict):
                size = len(block.get("base64", "")) * 3 // 4 // 1024
                parts.append(f"[{block.get('type')}: {block.get('mime_type', '?')}, ~{size} KB]")
        content = " ".join(parts)
    return content if len(content) < 300 else content[:300] + " …"


def print_messages(messages) -> None:
    """Skriv ut hvert steg i agentløkken så vi ser hva som skjedde."""
    for msg in messages:
        if isinstance(msg, HumanMessage):
            print(f"\n👤 Bruker: {msg.content}")
        elif isinstance(msg, AIMessage) and msg.tool_calls:
            # Claude svarer ikke ennå – den ber om å få kjøre et verktøy.
            for call in msg.tool_calls:
                print(f"\n🤖 Claude vil kalle verktøy: {call['name']}({call['args']})")
        elif isinstance(msg, ToolMessage):
            print(f"\n🔧 Verktøyet {msg.name} returnerte:\n{_preview(msg.content)}")
        elif isinstance(msg, AIMessage):
            print(f"\n🤖 Claude svarer:\n{msg.text}")


def main() -> None:
    if len(sys.argv) < 2:
        print('Bruk: uv run docxreader "spørsmål om et dokument"')
        sys.exit(1)

    # Importeres her så `--help`-lignende bruk ikke trenger API-nøkkel.
    from docxreader.agent import ask

    print_messages(ask(" ".join(sys.argv[1:])))
