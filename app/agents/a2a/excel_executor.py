import asyncio
from pathlib import Path

from a2a.server.agent_execution import AgentExecutor, RequestContext
from a2a.server.events import EventQueue
from a2a.utils import new_agent_text_message

from app.agents.crewai.excel_tool import excel_cleaning_tool


class ExcelCleaningExecutor(AgentExecutor):
    async def execute(
        self,
        context: RequestContext,
        event_queue: EventQueue,
    ) -> None:

        # Step 1: Receive the Excel file path
        file_path = context.get_user_input().strip()

        # Step 2: Validate the file
        if not file_path:
            await event_queue.enqueue_event(
                new_agent_text_message(
                    "Please provide an Excel file path."
                )
            )
            return

        source = Path(file_path).expanduser()

        if not source.is_file():
            await event_queue.enqueue_event(
                new_agent_text_message(
                    f"File not found: {source}"
                )
            )
            return

        # Step 3: Run the existing cleaning workflow
        result = await asyncio.to_thread(
            excel_cleaning_tool.run,
            file_path=str(source),
        )

        # Step 4: Return the cleaning results
        await event_queue.enqueue_event(
            new_agent_text_message(str(result))
        )

    async def cancel(
        self,
        context: RequestContext,
        event_queue: EventQueue,
    ) -> None:
        raise NotImplementedError(
            "Cancellation is not supported yet."
        )