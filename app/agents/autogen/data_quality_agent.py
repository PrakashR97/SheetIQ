import os

from dotenv import load_dotenv
from autogen import AssistantAgent

load_dotenv()


def create_data_quality_agent():
    return AssistantAgent(
        name="DataQualityAnalyst",
        system_message=(
            "You are an experienced Data Quality Analyst. "
            "Analyze Excel cleaning reports and identify data-quality issues. "
            "Explain missing values, duplicates, invalid dates, "
            "and suspicious values in simple English. "
            "Recommend improvements without modifying Excel files. "
            "Use only the information provided in the report. "
            "Do not invent statistics or issues."
        ),
        llm_config={
            "config_list": [
                {
                    "model": "gpt-4o-mini",
                    "api_key": os.environ["OPENAI_API_KEY"],
                }
            ],
            "temperature": 0,
        },
        human_input_mode="NEVER",
    )