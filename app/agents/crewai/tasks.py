from crewai import Task


def create_cleaning_task(coordinator):
    return Task(
        description=(
            "Prepare a cleaning plan for an Excel workbook. "
            "Requirements: process all sheets, remove fully blank rows, "
            "remove exact duplicate rows, replace missing name-related "
            "values with Unknown, replace other missing text with N/A, "
            "replace missing numeric values with 0, "
            "normalize dates using the user's selected format, "
            "trim whitespace, and report suspicious values. "
            "Preserve the original workbook and its formatting. "
            "Do not execute cleaning or modify any files."
        ),
        expected_output=(
            "A clear, numbered Excel cleaning plan describing "
            "the cleaning rules, validation checks, and output files."
        ),
        agent=coordinator
    )


def create_execution_task(coordinator, file_path: str):
    return Task(
        description=(
            f"Clean the Excel workbook at this exact path: {file_path}\n"
            "Use the Excel Data Cleaner tool exactly once. "
            "Do not modify the workbook manually. "
            "Report the output file path and cleaning statistics."
        ),
        expected_output=(
            "The cleaned Excel file path and a summary of "
            "rows processed, duplicates removed, blank rows removed, "
            "null values replaced, and suspicious values."
        ),
        agent=coordinator,
        tools=coordinator.tools,
    )