from pathlib import Path

from crewai.tools import tool

from app.services.excel_cleaner import clean_workbook
from app.services.report import export_report_csv
from app.agents.autogen.report_analyzer import analyze_cleaning_report


@tool("Excel Data Cleaner")
def excel_cleaning_tool(file_path: str) -> str:
    """
    Clean an Excel workbook using predefined cleaning rules.

    Args:
        file_path: Path to an existing .xlsx or .xlsm file.

    Returns:
        Cleaned Excel path, CSV report path, and AI analysis.
    """
    try:
        source = Path(file_path).expanduser().resolve()

        if not source.is_file():
            return f"Error: File not found: {source}"

        # Step 1: Clean the Excel workbook
        report = clean_workbook(source)

        # Step 2: Generate the CSV cleaning report
        csv_path = export_report_csv(report)

        # Step 3: Analyze the cleaning report using AutoGen
        try:
            analysis = analyze_cleaning_report(report)
            report.ai_analysis = analysis
            csv_path = export_report_csv(report)
        except Exception as error:
            analysis = f"AI analysis unavailable: {error}"

        # Step 4: Return cleaning results and AI analysis
        return (
            f"Excel cleaning completed successfully.\n"
            f"Output file: {report.output_file}\n"
            f"Cleaning report CSV: {csv_path}\n"
            f"Sheets processed: {report.sheets_processed}\n"
            f"Rows processed: {report.rows_processed}\n"
            f"Blank rows removed: {report.blank_rows_removed}\n"
            f"Duplicates removed: {report.duplicate_rows_removed}\n"
            f"Null values replaced: {report.null_values_replaced}\n"
            f"Values trimmed: {report.values_trimmed}\n"
            f"Suspicious values: {len(report.invalid_values)}\n"
            f"\n===== AUTOGEN DATA QUALITY ANALYSIS =====\n"
            f"{analysis}"
        )

    except Exception as error:
        return f"Excel cleaning failed: {error}"