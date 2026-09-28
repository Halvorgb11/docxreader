"""Agenten: kobler Claude sammen med verktøyene våre.

`create_agent` bygger en løkke som ser slik ut:

    1. Send meldingene til Claude.
    2. Svarer Claude med et verktøykall (tool call)?
         ja  -> kjør verktøyet, legg resultatet til som en ToolMessage, gå til 1.
         nei -> Claude har et ferdig svar. Stopp.

Alt agenten gjør lagres som en liste med meldinger:
- HumanMessage : det brukeren skrev
- AIMessage    : det Claude svarte (enten tekst eller en forespørsel om verktøykall)
- ToolMessage  : resultatet av et verktøykall, sendt tilbake til Claude

Samtaleminne (Conversation):
- En CHECKPOINTER lagrer meldingslisten etter hvert steg. InMemorySaver
  lagrer i minnet (borte når programmet avslutter).
- THREAD_ID sier hvilken samtale meldingene hører til. Samme thread_id =
  agenten husker tidligere spørsmål; ny thread_id = ny samtale.
- Med minne vokser meldingslisten for hvert spørsmål – også med store
  verktøysvar (hele dokumenter, bilder). ContextEditingMiddleware rydder:
  når konteksten blir stor, byttes gamle verktøysvar ut med en kort tekst
  i det som SENDES til Claude. Minnet selv beholder alt.
"""

import uuid

from dotenv import load_dotenv
from langchain.agents import create_agent
from langchain.agents.middleware import ClearToolUsesEdit, ContextEditingMiddleware
from langchain_core.language_models import BaseChatModel
from langchain_core.messages import BaseMessage
from langgraph.checkpoint.memory import InMemorySaver

from docxreader.tools import ALL_TOOLS

# Leser .env og legger ANTHROPIC_API_KEY i miljøvariablene.
# langchain-anthropic finner nøkkelen der automatisk.
load_dotenv()

# "leverandør:modellnavn". LangChain ser "anthropic:" og bruker
# langchain-anthropic (ChatAnthropic) under panseret.
MODEL = "anthropic:claude-sonnet-5"

SYSTEM_PROMPT = """Du er en assistent som svarer på spørsmål om dokumenter.
Bruk verktøyene dine til å lese filene brukeren viser til. Svar på norsk,
og oppgi hvor informasjonen står (overskrift, lysbilde, ark og rad, eller
del av e-posten/vedlegg) når det er relevant.
Hvis du ikke finner svaret i dokumentet, si det i stedet for å gjette.
Regn ikke i hodet med tall fra regneark eller CSV (også vedlegg): bruk
query_table med calculate/aggregate, så blir summer og produkter eksakte."""

# Kontekstrydding: når meldingene til Claude er over CONTEXT_TRIGGER_TOKENS
# (anslått), byttes alle verktøysvar unntatt de KEEP_TOOL_RESULTS nyeste ut.
CONTEXT_TRIGGER_TOKENS = 40_000
KEEP_TOOL_RESULTS = 4
CLEARED_PLACEHOLDER = (
    "[Dette verktøysvaret er fjernet for å spare plass. Kall verktøyet på nytt "
    "hvis du trenger innholdet igjen.]"
)


def build_agent(model: str | BaseChatModel = MODEL, checkpointer=None):
    """Lag agenten. Verktøyene i `tools` er de eneste Claude kan bruke.

    `model` kan byttes ut (testene bruker en falsk modell). Med `checkpointer`
    husker agenten samtaler, se Conversation.
    """
    return create_agent(
        model=model,
        tools=ALL_TOOLS,
        system_prompt=SYSTEM_PROMPT,
        checkpointer=checkpointer,
        # Middleware er kode som kjøres rundt hvert modellkall i agentløkken.
        middleware=[
            ContextEditingMiddleware(
                edits=[
                    ClearToolUsesEdit(
                        trigger=CONTEXT_TRIGGER_TOKENS,
                        keep=KEEP_TOOL_RESULTS,
                        placeholder=CLEARED_PLACEHOLDER,
                    )
                ]
            )
        ],
    )


def ask(question: str) -> list[BaseMessage]:
    """Still agenten ett spørsmål, uten minne, og returner alle meldingene."""
    agent = build_agent()
    # Input til agenten er en "state" med en liste meldinger.
    result = agent.invoke({"messages": [{"role": "user", "content": question}]})
    # Output er den oppdaterte staten: alle meldingene, inkludert verktøykall.
    return result["messages"]


class Conversation:
    """En samtale der agenten husker tidligere spørsmål og svar.

        samtale = Conversation()
        samtale.ask("Hvem har ansvar for design i samples/prosjektplan.docx?")
        samtale.ask("Og hvor mye er budsjettert for det?")  # "det" = design
    """

    def __init__(self, model: str | BaseChatModel = MODEL):
        # Én checkpointer kan holde mange samtaler; thread_id skiller dem.
        self.agent = build_agent(model, checkpointer=InMemorySaver())
        self.new_thread()

    def new_thread(self) -> None:
        """Start en ny samtale (glem det som er sagt så langt)."""
        self.thread_id = str(uuid.uuid4())

    @property
    def config(self) -> dict:
        # "configurable" er LangGraphs sted for innstillinger per kall.
        return {"configurable": {"thread_id": self.thread_id}}

    def history(self) -> list[BaseMessage]:
        """Alle meldingene i samtalen, slik checkpointeren har lagret dem."""
        return self.agent.get_state(self.config).values.get("messages", [])

    def ask(self, question: str) -> list[BaseMessage]:
        """Still et spørsmål i samtalen. Returnerer bare de NYE meldingene."""
        before = len(self.history())
        # Vi sender bare det nye spørsmålet; checkpointeren legger det til
        # etter de lagrede meldingene før Claude kalles.
        result = self.agent.invoke({"messages": [{"role": "user", "content": question}]}, self.config)
        return result["messages"][before:]
