import json

from app.agents.autogen.data_quality_agent import (
    create_data_quality_agent,
)


def analyze_cleaning_report(report):
    """Analyze an Excel cleaning report using AutoGen."""

    agent = create_data_quality_agent()

    report_data = report.to_dict()

    prompt = (
        "Analyze the following Excel cleaning report:\n\n"
        f"{json.dumps(report_data, indent=2, default=str)}\n\n"
        "Provide a concise data quality analysis covering:\n"
        "1. Sheets and rows processed.\n"
        "2. Duplicate and blank rows removed.\n"
        "3. Missing values replaced.\n"
        "4. Suspicious or invalid values.\n"
        "5. Recommendations for improving data quality.\n\n"
        "Use simple English. Do not invent information."
    )

    response = agent.generate_reply(
        messages=[
            {
                "role": "user",
                "content": prompt,
            }
        ]
    )

    if not isinstance(response, str) or not response.strip():
        raise ValueError("AutoGen returned an empty or invalid response")

    return response