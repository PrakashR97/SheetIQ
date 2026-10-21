
from typing import Any, Iterable

from app.services.cleaning_rules import is_blank


def is_fully_blank_row(values: Iterable[Any]) -> bool:
    """Check whether every cell in a row is blank."""
    return all(is_blank(value) for value in values)


def find_duplicate_rows(
    rows: Iterable[Iterable[Any]],
) -> list[int]:
    """
    Return zero-based data-row indexes to remove.
    Keep the first occurrence of each exact row.
    """
    seen: set[tuple[Any, ...]] = set()
    duplicates: list[int] = []

    for index, row in enumerate(rows):
        signature = tuple(row)

        if signature in seen:
            duplicates.append(index)
        else:
            seen.add(signature)

    return duplicates

