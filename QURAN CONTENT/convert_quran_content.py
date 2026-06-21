from __future__ import annotations

import argparse
import csv
import re
import sys
import unicodedata
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook, load_workbook


@dataclass(frozen=True)
class QuranConversionSummary:
    input_file: str
    output_file: str
    sheet_count: int
    row_count: int


def normalize_text(value) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    return " ".join(text.replace("\u00a0", " ").split())


def normalize_id(value) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    try:
        number = Decimal(text)
    except (InvalidOperation, ValueError):
        return text
    if number == number.to_integral_value():
        return str(number.quantize(Decimal(1)))
    return format(number.normalize(), "f")


def quran_id_column(path: Path | str) -> str:
    name = normalize_text(Path(path).stem).casefold()
    if "juz" in name:
        return "juz_id"
    if "page" in name:
        return "page_id"
    if "surah" in name:
        return "surah_id"
    return "id"


def language_id(path: Path | str) -> str:
    name = normalize_text(Path(path).name)
    match = re.match(r"([A-Za-z]{2})(?:[_\-\s]|$)", name)
    if match:
        return match.group(1).lower()
    return ""


def discover_workbooks(directory: Path | str) -> list[Path]:
    return sorted(path for path in Path(directory).rglob("*.xlsx") if not path.name.startswith("~$"))


def convert_quran_content_folder(
    input_dir: Path | str = Path("INPUT"),
    output_dir: Path | str = Path("OUTPUT"),
) -> list[QuranConversionSummary]:
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    summaries = []
    for input_path in discover_workbooks(input_dir):
        output_path = output_dir / input_path.name
        summaries.append(convert_quran_content_workbook(input_path, output_path))

    write_summary_csv(output_dir / "conversion_summary.csv", summaries)
    return summaries


def convert_quran_content_workbook(input_path: Path | str, output_path: Path | str) -> QuranConversionSummary:
    input_path = Path(input_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    id_header = quran_id_column(input_path)
    lang = language_id(input_path)
    workbook = load_workbook(input_path, read_only=True, data_only=True)
    output_workbook = Workbook(write_only=True)
    row_count = 0
    sheet_count = 0
    try:
        for sheet in workbook.worksheets:
            output_sheet = output_workbook.create_sheet(title=sheet.title)
            output_sheet.append([id_header, "language_id", "contents"])
            sheet_count += 1

            rows = sheet.iter_rows(values_only=True)
            header = next(rows, None)
            columns = _header_map(header)
            id_index = columns.get("id")
            response_index = columns.get("response")
            if id_index is None or response_index is None:
                continue

            for row in rows:
                if not row or not any(cell is not None for cell in row):
                    continue
                row_id = normalize_id(_cell(row, id_index))
                contents = normalize_text(_cell(row, response_index))
                if not row_id or not contents:
                    continue
                output_sheet.append([row_id, lang, contents])
                row_count += 1

        output_workbook.save(output_path)
    finally:
        workbook.close()

    return QuranConversionSummary(
        input_file=input_path.name,
        output_file=output_path.name,
        sheet_count=sheet_count,
        row_count=row_count,
    )


def write_summary_csv(path: Path | str, summaries: Iterable[QuranConversionSummary]) -> None:
    path = Path(path)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["input_file", "output_file", "sheet_count", "row_count"])
        writer.writeheader()
        for summary in summaries:
            writer.writerow(
                {
                    "input_file": summary.input_file,
                    "output_file": summary.output_file,
                    "sheet_count": summary.sheet_count,
                    "row_count": summary.row_count,
                }
            )


def print_summary(summaries: Iterable[QuranConversionSummary]) -> None:
    for summary in summaries:
        print(
            f"{summary.input_file} -> {summary.output_file}: "
            f"sheets={summary.sheet_count}, rows={summary.row_count}"
        )


def _header_map(header) -> dict[str, int]:
    if not header:
        return {}
    return {normalize_text(value).casefold(): index for index, value in enumerate(header) if normalize_text(value)}


def _cell(row, index: int):
    return row[index] if index < len(row) else None


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Convert Quran content workbooks to id/language_id/contents columns.")
    parser.add_argument("--input-dir", type=Path, default=Path("INPUT"))
    parser.add_argument("--output-dir", type=Path, default=Path("OUTPUT"))
    args = parser.parse_args(argv)

    summaries = convert_quran_content_folder(args.input_dir, args.output_dir)
    print_summary(summaries)

    if not summaries:
        print("No Quran content workbooks found.", file=sys.stderr)
        return 1
    print(f"Converted Quran content workbooks written to: {args.output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
