"""docxreader: la en Claude-agent lese dokumenter (.docx, .pptx, .md, .txt) via LangChain-verktøy.

Kjør fra kommandolinjen:
    uv run docxreader "Hva står det i samples/prosjektplan.docx?"
"""

import sys

from langchain_core.messages import AIMessage, HumanMessage, ToolMessage


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
            preview = msg.content if len(msg.content) < 300 else msg.content[:300] + " …"
            print(f"\n🔧 Verktøyet {msg.name} returnerte:\n{preview}")
        elif isinstance(msg, AIMessage):
            print(f"\n🤖 Claude svarer:\n{msg.text}")


def main() -> None:
    if len(sys.argv) < 2:
        print('Bruk: uv run docxreader "spørsmål om et dokument"')
        sys.exit(1)

    # Importeres her så `--help`-lignende bruk ikke trenger API-nøkkel.
    from docxreader.agent import ask

    print_messages(ask(" ".join(sys.argv[1:])))
