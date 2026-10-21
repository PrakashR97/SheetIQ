from pathlib import Path
from tempfile import TemporaryDirectory

import pytest
from openpyxl import Workbook

from app.services.excel_cleaner import clean_workbook


def test_missing_file():
    with TemporaryDirectory() as folder:
        missing = Path(folder) / "missing.xlsx"

        with pytest.raises(FileNotFoundError):
            clean_workbook(missing)


def test_unsupported_file():
    with TemporaryDirectory() as folder:
        csv_file = Path(folder) / "sample.csv"
        csv_file.write_text("Name,Age\nAlice,25\n")

        with pytest.raises(ValueError):
            clean_workbook(csv_file)


def test_original_file_protection():
    with TemporaryDirectory() as folder:
        source = Path(folder) / "original.xlsx"

        wb = Workbook()
        wb.active.append(["Name"])
        wb.active.append([" Alice "])
        wb.save(source)
        wb.close()

        original_bytes = source.read_bytes()

        with pytest.raises(ValueError):
            clean_workbook(source, output_path=source)

        assert source.read_bytes() == original_bytes