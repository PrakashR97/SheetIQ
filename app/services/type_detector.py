
from datetime import date, datetime
from collections import Counter
from typing import Any


def detect_value_type(value: Any) -> str:
    """Classify a single Python or Excel cell value."""
    if value is None:
        return "blank"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, datetime):
        return "datetime"
    if isinstance(value, date):
        return "date"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, str):
        return "text"
    return "unknown"


def detect_column_type(values: list[Any]) -> str:
    """Infer a column type from nonblank values."""
    counts = Counter(
        detect_value_type(value)
        for value in values
        if value is not None and not (
            isinstance(value, str) and not value.strip()
        )
    )

    if not counts:
        return "empty"

    # Treat date and datetime as one compatible family.
    if set(counts).issubset({"date", "datetime"}):
        return "datetime" if "datetime" in counts else "date"

    # Avoid treating mixed columns as one misleading type.
    detected_type, frequency = counts.most_common(1)[0]

    if frequency / sum(counts.values()) < 0.8:
        return "mixed"

    if len(counts) > 1:
        return "mixed"

    return detected_type

