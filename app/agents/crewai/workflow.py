from crewai import Crew, Process

from app.agents.crewai.coordinator import create_coordinator
from app.agents.crewai.tasks import (
    create_cleaning_task,
    create_execution_task,
)

def create_execution_crew(file_path: str):
    coordinator = create_coordinator()

    task = create_execution_task(
        coordinator,
        file_path
    )

    crew = Crew(
        agents=[coordinator],
        tasks=[task],
        process=Process.sequential,
        verbose=True
    )

    return crew

def create_cleaning_crew():
    coordinator = create_coordinator()
    task = create_cleaning_task(coordinator)

    crew = Crew(
        agents=[coordinator],
        tasks=[task],
        process=Process.sequential,
        verbose=True
    )



    return crew