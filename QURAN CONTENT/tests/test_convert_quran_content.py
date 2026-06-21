import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook


MODULE_PATH = Path(__file__).resolve().parents[1] / "convert_quran_content.py"
spec = importlib.util.spec_from_file_location("convert_quran_content", MODULE_PATH)
quran = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = quran
spec.loader.exec_module(quran)


class ConvertQuranContentTest(unittest.TestCase):
    def test_infers_output_id_column_from_quran_content_filename(self):
        self.assertEqual(quran.quran_id_column(Path("EN_Quran Juz Content.xlsx")), "juz_id")
        self.assertEqual(quran.quran_id_column(Path("EN_Quran Page Content.xlsx")), "page_id")
        self.assertEqual(quran.quran_id_column(Path("EN_Quran Surah Content.xlsx")), "surah_id")

    def test_converts_workbook_to_requested_three_columns(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_path = root / "EN_Quran Surah Content.xlsx"
            output_path = root / "OUTPUT" / input_path.name
            self._write_source_workbook(input_path, rows=[["1", "chatgpt", "id: 1", "Yes", "time", "response one"]])

            summary = quran.convert_quran_content_workbook(input_path, output_path)

            self.assertEqual(summary.input_file, input_path.name)
            self.assertEqual(summary.output_file, output_path.name)
            self.assertEqual(summary.row_count, 1)
            self.assertEqual(
                self._read_rows(output_path),
                [
                    ("surah_id", "language_id", "contents"),
                    ("1", "en", "response one"),
                ],
            )

    def test_converts_every_sheet_and_skips_rows_without_id_or_response(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_path = root / "BN_Quran Page Content.xlsx"
            output_path = root / "OUTPUT" / input_path.name

            workbook = Workbook()
            first = workbook.active
            first.title = "Results"
            first.append(["ID", "Platform", "Keyword", "Success", "Timestamp", "Response"])
            first.append([2.0, "chatgpt", "id_page: 2", "Yes", "time", "page response"])
            first.append([3, "chatgpt", "id_page: 3", "Yes", "time", None])
            second = workbook.create_sheet(title="Extra")
            second.append(["ID", "Response"])
            second.append([4, "extra response"])
            workbook.save(input_path)

            summary = quran.convert_quran_content_workbook(input_path, output_path)

            self.assertEqual(summary.row_count, 2)
            self.assertEqual(
                self._read_rows(output_path, sheet_name="Results"),
                [
                    ("page_id", "language_id", "contents"),
                    ("2", "bn", "page response"),
                ],
            )
            self.assertEqual(
                self._read_rows(output_path, sheet_name="Extra"),
                [
                    ("page_id", "language_id", "contents"),
                    ("4", "bn", "extra response"),
                ],
            )

    def test_converts_folder_outputs_each_input_workbook(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_dir = root / "INPUT"
            output_dir = root / "OUTPUT"
            input_dir.mkdir()
            self._write_source_workbook(
                input_dir / "EN_Quran Juz Content.xlsx",
                rows=[["1", "chatgpt", "id: 1", "Yes", "time", "juz response"]],
            )
            self._write_source_workbook(
                input_dir / "EN_Quran Surah Content.xlsx",
                rows=[["1", "chatgpt", "id: 1", "Yes", "time", "surah response"]],
            )

            summaries = quran.convert_quran_content_folder(input_dir, output_dir)

            self.assertEqual([summary.output_file for summary in summaries], ["EN_Quran Juz Content.xlsx", "EN_Quran Surah Content.xlsx"])
            self.assertTrue((output_dir / "EN_Quran Juz Content.xlsx").exists())
            self.assertTrue((output_dir / "EN_Quran Surah Content.xlsx").exists())

    def _write_source_workbook(self, path, rows):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Results"
        sheet.append(["ID", "Platform", "Keyword", "Success", "Timestamp", "Response"])
        for row in rows:
            sheet.append(row)
        workbook.save(path)

    def _read_rows(self, path, sheet_name="Results"):
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            sheet = workbook[sheet_name]
            return [tuple(row) for row in sheet.iter_rows(values_only=True)]
        finally:
            workbook.close()


if __name__ == "__main__":
    unittest.main()
