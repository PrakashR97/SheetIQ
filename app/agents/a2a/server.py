from a2a.server.apps import A2AStarletteApplication
from a2a.server.request_handlers import DefaultRequestHandler
from a2a.server.tasks import InMemoryTaskStore
from a2a.types import (
    AgentCard,
    AgentCapabilities,
    AgentSkill,
)

from app.agents.a2a.excel_executor import ExcelCleaningExecutor


# Step 1: Define the agent's capabilities
excel_skill = AgentSkill(
    id="excel_data_cleaning",
    name="Excel Data Cleaning",
    description=(
        "Clean Excel workbooks, remove duplicates, "
        "handle missing values, normalize dates, "
        "and generate a CSV data quality report."
    ),
    tags=["excel", "data-cleaning", "crewai", "autogen"],
    examples=["/path/to/sample.xlsx"],
)


# Step 2: Create the agent card
agent_card = AgentCard(
    name="Excel Cleaning Agent",
    description=(
        "An A2A agent that uses CrewAI and AutoGen "
        "to clean Excel workbooks and generate CSV reports."
    ),
    url="http://127.0.0.1:8000/",
    version="1.0.0",
    default_input_modes=["text"],
    default_output_modes=["text"],
    capabilities=AgentCapabilities(streaming=False),
    skills=[excel_skill],
)


# Step 3: Connect the executor
request_handler = DefaultRequestHandler(
    agent_executor=ExcelCleaningExecutor(),
    task_store=InMemoryTaskStore(),
)


# Step 4: Create the A2A application
server = A2AStarletteApplication(
    agent_card=agent_card,
    http_handler=request_handler,
)

app = server.build()