"""Tester for samtaleminne (checkpointer + thread_id) og kontekstrydding.

Kjører uten Claude: vi lager falske chatmodeller ved å arve fra LangChains
BaseChatModel. Alt som oppfører seg som en chatmodell kan settes inn i agenten.
"""

from pathlib import Path

from langchain_core.language_models import BaseChatModel
from langchain_core.messages import AIMessage, HumanMessage, ToolMessage
from langchain_core.outputs import ChatGeneration, ChatResult

from docxreader import agent, chat
from docxreader.agent import CLEARED_PLACEHOLDER, Conversation

SAMPLE = str(Path(__file__).parent.parent / "samples" / "prosjektplan.docx")


def _result(message: AIMessage) -> ChatResult:
    return ChatResult(generations=[ChatGeneration(message=message)])


class CountingModel(BaseChatModel):
    """Svarer med hvor mange brukerspørsmål den får se – viser om minnet virker."""

    @property
    def _llm_type(self) -> str:
        return "counting"

    def bind_tools(self, tools, **kwargs):  # agenten gir modellen verktøyene
        return self

    def _generate(self, messages, stop=None, run_manager=None, **kwargs):
        questions = [m.content for m in messages if isinstance(m, HumanMessage)]
        return _result(AIMessage(f"Jeg ser {len(questions)} spørsmål: {' | '.join(questions)}"))


SEEN: list[list] = []  # meldingene ToolReadingModel fikk se, per kall


class ToolReadingModel(BaseChatModel):
    """Ber alltid om å lese prosjektplanen, og svarer "ferdig" når svaret kommer."""

    @property
    def _llm_type(self) -> str:
        return "tool-reading"

    def bind_tools(self, tools, **kwargs):
        return self

    def _generate(self, messages, stop=None, run_manager=None, **kwargs):
        if isinstance(messages[-1], HumanMessage):
            call_id = f"kall_{sum(isinstance(m, HumanMessage) for m in messages)}"
            return _result(AIMessage("", tool_calls=[{"name": "read_document", "args": {"path": SAMPLE}, "id": call_id}]))
        SEEN.append(list(messages))
        return _result(AIMessage("ferdig"))


# --- Minne --------------------------------------------------------------------


def test_conversation_remembers_earlier_questions():
    conversation = Conversation(CountingModel())
    first = conversation.ask("Hvem har ansvar for design?")
    assert first[-1].content == "Jeg ser 1 spørsmål: Hvem har ansvar for design?"
    second = conversation.ask("Og budsjettet for det?")
    assert second[-1].content == "Jeg ser 2 spørsmål: Hvem har ansvar for design? | Og budsjettet for det?"


def test_ask_returns_only_new_messages():
    conversation = Conversation(CountingModel())
    conversation.ask("En")
    new = conversation.ask("To")
    assert [type(m) for m in new] == [HumanMessage, AIMessage]
    assert len(conversation.history()) == 4


def test_new_thread_forgets():
    conversation = Conversation(CountingModel())
    conversation.ask("En")
    old_thread = conversation.thread_id
    conversation.new_thread()
    assert conversation.thread_id != old_thread
    assert conversation.ask("To")[-1].content == "Jeg ser 1 spørsmål: To"
    assert conversation.history()[0].content == "To"


def test_two_conversations_do_not_share_memory():
    a, b = Conversation(CountingModel()), Conversation(CountingModel())
    a.ask("Til A")
    assert b.ask("Til B")[-1].content == "Jeg ser 1 spørsmål: Til B"


# --- Kontekstrydding ------------------------------------------------------------


def test_old_tool_results_are_cleared_for_the_model_but_kept_in_memory(monkeypatch):
    monkeypatch.setattr(agent, "CONTEXT_TRIGGER_TOKENS", 100)  # rydd med en gang
    monkeypatch.setattr(agent, "KEEP_TOOL_RESULTS", 1)
    SEEN.clear()
    conversation = Conversation(ToolReadingModel())
    for question in ["En", "To", "Tre"]:
        conversation.ask(question)

    # Det modellen fikk se i siste kall: bare det nyeste verktøysvaret er helt.
    tool_messages = [m for m in SEEN[-1] if isinstance(m, ToolMessage)]
    assert [m.content == CLEARED_PLACEHOLDER for m in tool_messages] == [True, True, False]
    assert "Prosjektplan" in tool_messages[-1].content

    # Minnet (checkpointeren) har fortsatt alle svarene i sin helhet.
    stored = [m for m in conversation.history() if isinstance(m, ToolMessage)]
    assert len(stored) == 3 and all("Prosjektplan" in m.content for m in stored)


def test_small_conversations_are_not_cleared():
    SEEN.clear()
    conversation = Conversation(ToolReadingModel())  # standardgrensen (40 000 tokens)
    conversation.ask("En")
    conversation.ask("To")
    assert all(m.content != CLEARED_PLACEHOLDER for m in SEEN[-1] if isinstance(m, ToolMessage))


# --- Chat i terminalen ------------------------------------------------------------


def _typed(*lines):
    """Later som brukeren skriver disse linjene, og så Ctrl-D."""
    remaining = list(lines)

    def read(prompt):
        if not remaining:
            raise EOFError
        return remaining.pop(0)

    return read


def test_chat_keeps_memory_until_new(capsys):
    chat(Conversation(CountingModel()), read=_typed("Første", "Andre", "/ny", "Tredje"))
    out = capsys.readouterr().out
    assert "Jeg ser 2 spørsmål: Første | Andre" in out
    assert "🆕 Ny samtale" in out
    assert "Jeg ser 1 spørsmål: Tredje" in out
    assert "👤 Bruker:" not in out  # spørsmålet skrives ikke ut to ganger


def test_chat_stops_on_exit_command_and_skips_empty_lines(capsys):
    chat(Conversation(CountingModel()), read=_typed("", "Hei", "/avslutt", "Aldri"))
    out = capsys.readouterr().out
    assert "Jeg ser 1 spørsmål: Hei" in out
    assert "Aldri" not in out


def test_chat_survives_errors(capsys):
    class Broken:
        def ask(self, question):
            raise ConnectionError("ingen nett")

    chat(Broken(), read=_typed("Hei", "Igjen"))
    assert capsys.readouterr().out.count("⚠️ Feil: ConnectionError: ingen nett") == 2
