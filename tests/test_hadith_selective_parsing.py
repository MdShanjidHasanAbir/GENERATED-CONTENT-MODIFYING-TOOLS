import importlib.util
import json
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook


def load_hadith_module():
    module_path = Path(__file__).resolve().parents[1] / "SINGLE HADITH CONTENT" / "reconcile_hadith_outputs.py"
    spec = importlib.util.spec_from_file_location("reconcile_hadith_outputs_for_test", module_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


hadith = load_hadith_module()


def write_workbook(path: Path, header: list[str], rows: list[list[str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(header)
    for row in rows:
        sheet.append(row)
    workbook.save(path)


def read_rows(path: Path) -> list[tuple]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        return [tuple(row) for row in workbook.active.iter_rows(values_only=True)]
    finally:
        workbook.close()


class HadithSelectiveParsingTest(unittest.TestCase):
    def test_selected_reconciled_files_update_book_wise_final_without_removing_existing_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            books_path = root / "BOOKS.xlsx"
            reconciled_dir = root / "reconciled_output"
            final_dir = root / "BOOK WISE FINAL"
            write_workbook(
                books_path,
                ["id", "language_id", "book_name"],
                [["101", "en", "BUKHARI"], ["102", "en", "MUSLIM"]],
            )
            write_workbook(
                reconciled_dir / "EN_BUKHARI_MARGE.xlsx",
                ["id", "response"],
                [["1", "{\"heading\":\"bukhari\"}"]],
            )
            write_workbook(
                reconciled_dir / "EN_MUSLIM_MARGE.xlsx",
                ["id", "response"],
                [["2", "{\"heading\":\"muslim\"}"]],
            )
            write_workbook(
                final_dir / "EN" / "EN_MUSLIM.xlsx",
                ["book_id", "language_id", "hadith_id", "content"],
                [["existing", "en", "existing", "keep me"]],
            )

            counts = hadith.write_book_wise_final(
                reconciled_dir,
                books_path,
                final_dir,
                source_paths=[reconciled_dir / "EN_BUKHARI_MARGE.xlsx"],
                clear_existing=False,
            )

            self.assertEqual(counts, {"EN_BUKHARI": 1})
            self.assertEqual(
                read_rows(final_dir / "EN" / "EN_BUKHARI.xlsx"),
                [
                    ("book_id", "language_id", "hadith_id", "content"),
                    ("101", "en", "1", "{\"heading\":\"bukhari\"}"),
                ],
            )
            self.assertEqual(
                read_rows(final_dir / "EN" / "EN_MUSLIM.xlsx"),
                [
                    ("book_id", "language_id", "hadith_id", "content"),
                    ("existing", "en", "existing", "keep me"),
                ],
            )

    def test_selected_final_files_update_content_without_removing_existing_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            books_path = root / "BOOKS.xlsx"
            final_dir = root / "BOOK WISE FINAL"
            updated_dir = root / "BOOK WISE FINAL UPDATED CONTENT"
            write_workbook(
                books_path,
                ["id", "language_id", "book_name"],
                [["101", "en", "BUKHARI"], ["102", "en", "MUSLIM"]],
            )
            write_workbook(
                final_dir / "EN" / "EN_BUKHARI.xlsx",
                ["book_id", "language_id", "hadith_id", "content"],
                [["101", "en", "1", "plain content"]],
            )
            write_workbook(
                updated_dir / "EN" / "EN_MUSLIM.xlsx",
                ["book_id", "language_id", "hadith_id", "content"],
                [["existing", "en", "existing", "keep me"]],
            )

            counts = hadith.write_updated_content_workbooks(
                final_dir,
                updated_dir,
                books_path,
                source_paths=[final_dir / "EN" / "EN_BUKHARI.xlsx"],
                clear_existing=False,
            )

            self.assertEqual(counts, {"EN/EN_BUKHARI.xlsx": 1})
            rows = read_rows(updated_dir / "EN" / "EN_BUKHARI.xlsx")
            converted = json.loads(rows[1][3])
            self.assertEqual(converted["heading"]["description"], "plain content")
            self.assertEqual(rows[0], ("book_id", "language_id", "hadith_id", "content", "content_error_details"))
            self.assertEqual(
                read_rows(updated_dir / "EN" / "EN_MUSLIM.xlsx"),
                [
                    ("book_id", "language_id", "hadith_id", "content"),
                    ("existing", "en", "existing", "keep me"),
                ],
            )


if __name__ == "__main__":
    unittest.main()
