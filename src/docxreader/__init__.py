"""docxreader: la en Claude-agent lese dokumenter via LangChain-verktøy.

Kjør fra kommandolinjen:
    uv run docxreader "Hva står det i samples/prosjektplan.docx?"   (ett spørsmål)
    uv run docxreader                                               (chat med minne)
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


def print_messages(messages, show_user: bool = True) -> None:
    """Skriv ut hvert steg i agentløkken så vi ser hva som skjedde."""
    for msg in messages:
        if isinstance(msg, HumanMessage):
            if show_user:
                print(f"\n👤 Bruker: {msg.content}")
        elif isinstance(msg, AIMessage) and msg.tool_calls:
            # Claude svarer ikke ennå – den ber om å få kjøre et verktøy.
            for call in msg.tool_calls:
                print(f"\n🤖 Claude vil kalle verktøy: {call['name']}({call['args']})")
        elif isinstance(msg, ToolMessage):
            print(f"\n🔧 Verktøyet {msg.name} returnerte:\n{_preview(msg.content)}")
        elif isinstance(msg, AIMessage):
            print(f"\n🤖 Claude svarer:\n{msg.text}")


CHAT_HELP = """💬 Samtale med minne – agenten husker det du har spurt om.
   Skriv spørsmål om filer, f.eks. "Hva står det i samples/prosjektplan.docx?"
   /ny = ny samtale (glem alt)   /avslutt eller Ctrl-D = avslutt"""


def chat(conversation=None, read=input) -> None:
    """Chat i terminalen. Hele økten er én samtale (én thread_id) til /ny.
    `read` kan byttes ut i tester (i stedet for å lese fra tastaturet)."""
    from docxreader.agent import Conversation  # trenger API-nøkkel først her

    conversation = conversation or Conversation()
    print(CHAT_HELP)
    while True:
        try:
            question = read("\n👤 Du: ").strip()
        except (EOFError, KeyboardInterrupt):  # Ctrl-D / Ctrl-C
            print()
            return
        if not question:
            continue
        if question in ("/avslutt", "/exit", "/quit"):
            return
        if question == "/ny":
            conversation.new_thread()
            print("🆕 Ny samtale – tidligere spørsmål er glemt.")
            continue
        try:
            print_messages(conversation.ask(question), show_user=False)
        except Exception as e:  # f.eks. nettverksfeil – ikke mist samtalen
            print(f"\n⚠️ Feil: {type(e).__name__}: {e}")


def main() -> None:
    # Uten spørsmål: chat med minne. Med spørsmål: ett svar, uten minne.
    if len(sys.argv) < 2:
        chat()
        return

    # Importeres her så `--help`-lignende bruk ikke trenger API-nøkkel.
    from docxreader.agent import ask

    print_messages(ask(" ".join(sys.argv[1:])))
