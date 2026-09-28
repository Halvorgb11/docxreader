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
"""

from dotenv import load_dotenv
from langchain.agents import create_agent
from langchain_core.messages import BaseMessage

from docxreader.tools import docx_outline, read_docx, read_docx_section, search_docx

# Leser .env og legger ANTHROPIC_API_KEY i miljøvariablene.
# langchain-anthropic finner nøkkelen der automatisk.
load_dotenv()

# "leverandør:modellnavn". LangChain ser "anthropic:" og bruker
# langchain-anthropic (ChatAnthropic) under panseret.
MODEL = "anthropic:claude-sonnet-5"

SYSTEM_PROMPT = """Du er en assistent som svarer på spørsmål om dokumenter.
Bruk verktøyene dine til å lese filene brukeren viser til. Svar på norsk,
og oppgi hvilken overskrift informasjonen står under når det er relevant.
Hvis du ikke finner svaret i dokumentet, si det i stedet for å gjette."""


def build_agent():
    """Lag agenten. Verktøyene i `tools` er de eneste Claude kan bruke."""
    return create_agent(
        model=MODEL,
        tools=[read_docx, docx_outline, read_docx_section, search_docx],
        system_prompt=SYSTEM_PROMPT,
    )


def ask(question: str) -> list[BaseMessage]:
    """Still agenten ett spørsmål og returner alle meldingene fra kjøringen."""
    agent = build_agent()
    # Input til agenten er en "state" med en liste meldinger.
    result = agent.invoke({"messages": [{"role": "user", "content": question}]})
    # Output er den oppdaterte staten: alle meldingene, inkludert verktøykall.
    return result["messages"]
