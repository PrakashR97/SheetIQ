
import re
from typing import Any


NAME_COLUMN_PATTERN = re.compile(
    r"(^|_)(name|first_name|last_name|middle_name|full_name|"
    r"customer_name|member_name|employee_name|person_name|"
    r"auditor|client|patient|customer|member|employee|"
    r"doctor|provider|agent|contact)(_|$)",
    re.IGNORECASE,
)


def is_blank(value: Any) -> bool:
    """Return True for empty cells and whitespace-only strings."""
    return value is None or (
        isinstance(value, str) and not value.strip()
    )


def is_name_column(column_name: Any) -> bool:
    """Identify columns whose missing values should become Unknown."""
    if not isinstance(column_name, str):
        return False

    normalized = re.sub(
        r"[\s\-]+", "_", column_name.strip().lower()
    )

    return bool(NAME_COLUMN_PATTERN.search(normalized))


def normalize_text(value: Any, standardize_case: bool = False) -> Any:
    """Trim text and optionally standardize its case."""
    if not isinstance(value, str):
        return value

    value = value.strip()

    if standardize_case:
        value = value.casefold()

    return value


def null_replacement(column_name: Any, value: Any) -> Any:
    """Return the required replacement for a blank non-date cell."""
    if not is_blank(value):
        return value

    return "Unknown" if is_name_column(column_name) else "N/A"

