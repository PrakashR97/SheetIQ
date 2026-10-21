
from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class CleaningConfig:
    # Workbook processing
    process_all_sheets: bool = True
    remove_blank_rows: bool = True
    remove_exact_duplicates: bool = True
    preserve_original: bool = True

    # Cleaning options
    auto_detect_types: bool = True
    trim_whitespace: bool = True
    standardize_case: bool = True
    clean_numbers: bool = True
    validate_values: bool = True
    generate_report: bool = True

    # Null handling
    name_null_value: str = "Unknown"
    other_null_value: str = "N/A"

    # User-selectable output date format
    date_format: str = "%Y-%m-%d"

    # Output
    output_suffix: str = "_cleaned"


DEFAULT_CONFIG = CleaningConfig()

