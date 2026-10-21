from crewai import Agent, LLM
from dotenv import load_dotenv
from app.agents.crewai.excel_tool import excel_cleaning_tool

load_dotenv()


def create_coordinator():
    llm = LLM(
        model="openai/gpt-4o-mini",
        temperature=0
    )

    agent = Agent(
        role="Excel Cleaning Coordinator",
        goal=(
            "Understand Excel data cleaning requirements "
            "and create a clear cleaning plan."
        ),
        backstory=(
            "You are an experienced Data Engineer "
            "specializing in Excel data quality."
        ),
        llm=llm,
        tools=[excel_cleaning_tool],
        verbose=True,
        allow_delegation=False
    )



    return agent