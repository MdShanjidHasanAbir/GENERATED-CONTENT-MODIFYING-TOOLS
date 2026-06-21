# Dua Content Workflow Design

## Goal

Add a `Dua Content` workflow that mirrors the Quran content workflow pattern:

1. Convert workbook rows from `DUA CONTENT/INPUT` into clean output workbooks.
2. Update selected output workbook `contents` cells into the approved Dua JSON structure.

The first step writes to:

`DUA CONTENT/OUTPUT`

The second step writes updated copies to:

`DUA CONTENT/UPDATED CONTENT`

The updater does not modify files in `DUA CONTENT/OUTPUT`.

## Current Folder Structure

The workspace contains:

- `DUA CONTENT/INPUT`
- `DUA CONTENT/OUTPUT`

The observed input workbook is:

`DUA CONTENT/INPUT/EN _SINGLE DUA (21-6-26).xlsx`

Its `Results` sheet has the same source shape as Quran input workbooks:

- `ID`
- `Platform`
- `Keyword`
- `Success`
- `Timestamp`
- `Response`

The `Response` cells contain JSON-like data, often wrapped as a list with a `result` object.

## Conversion Step

The conversion step reads every `.xlsx` workbook under `DUA CONTENT/INPUT` and writes matching workbooks under `DUA CONTENT/OUTPUT`.

Output columns:

- `id`
- `language_id`
- `contents`

Field mapping:

- `id` comes from the source `ID` column, normalized like the Quran converter.
- `language_id` comes from the leading language code in the source filename, such as `EN`, normalized to lowercase.
- `contents` comes from the source `Response` column.

Rows with missing `ID` or missing `Response` are skipped.

Sheets without both `ID` and `Response` headers are copied as headers only with no converted rows.

## Target JSON Structure

The Dua updater uses one JSON structure:

```json
{
  "meta": {
    "title": "",
    "description": ""
  },
  "heading": {
    "title": "",
    "description": ""
  },
  "faqs": []
}
```

The user-provided sample had the heading description outside the `heading` object because of brace placement. The implementation will use the corrected structure above so the title and description are grouped consistently.

## Current Input JSON Shapes

Existing `contents` cells may contain wrapped JSON:

```json
[
  {
    "keyword": "Dua title",
    "result": {
      "meta": {},
      "heading": {},
      "description": "...",
      "faqs": []
    }
  }
]
```

They may also contain direct JSON objects.

The updater normalizes both forms. For wrapped rows, it uses the first object in the list and then uses its `result` object as the source content.

## Field Mapping

The updater writes only the approved target keys.

For all rows:

- `meta.title` comes from source `meta.title`
- `meta.description` comes from source `meta.description`
- `heading.title` comes from source `heading.title`, source string `heading`, source `title`, or source `keyword`
- `heading.description` comes from source `heading.description`, source top-level `description`, source `summary`, or an empty string
- `faqs` comes from source `faqs`; if missing, use an empty list

If a `contents` cell contains invalid JSON, the updater still writes a valid target JSON object and places the original text in `heading.description`.

The output JSON is compact and preserves Unicode characters.

## Script Design

Create a new script:

`DUA CONTENT/convert_dua_content.py`

The script contains reusable functions:

- `normalize_text(value)`
- `normalize_id(value)`
- `language_id(path)`
- `discover_workbooks(directory)`
- `convert_dua_content_folder(input_dir, output_dir)`
- `convert_dua_content_workbook(input_path, output_path)`
- `normalize_dua_content_json(value)`
- `available_dua_output_workbooks(output_dir)`
- `update_dua_output_workbook(input_path, output_path)`
- `update_dua_output_files(output_dir, updated_dir, filenames)`

The CLI supports two modes:

```powershell
python convert_dua_content.py convert
```

```powershell
python convert_dua_content.py update-output-json --files "EN _SINGLE DUA (21-6-26).xlsx"
```

For backward simplicity, running the script without a subcommand performs `convert`.

If no files are provided in update mode, the script updates all `.xlsx` files found in `DUA CONTENT/OUTPUT`.

## Dashboard Design

The dashboard gains a third sidebar workflow:

`Dua Content`

The Dua panel contains:

- primary button: `Convert Input to Output`
- file checkbox list from `DUA CONTENT/OUTPUT`
- secondary button: `Update selected JSON`
- status and logs for each action

There is no JSON structure selector for Dua because only one structure is supported.

The dashboard server gains:

- `dua` workflow for conversion
- `dua-update-json` workflow for selected output updates
- `/api/dua-output-files` endpoint that lists `.xlsx` files in `DUA CONTENT/OUTPUT`

## Data Flow

1. User selects `Dua Content` in the dashboard.
2. User clicks `Convert Input to Output`.
3. Server runs `DUA CONTENT/convert_dua_content.py convert`.
4. Script writes converted workbooks to `DUA CONTENT/OUTPUT`.
5. Dashboard refreshes the Dua output file list.
6. User selects one or more output workbook files.
7. User clicks `Update selected JSON`.
8. Server runs `DUA CONTENT/convert_dua_content.py update-output-json --files ...`.
9. Script writes updated copies to `DUA CONTENT/UPDATED CONTENT`.

## Error Handling

If there are no input workbooks, conversion exits with a nonzero status and prints a clear message.

If the user tries to update a selected file that does not exist under `OUTPUT`, the updater reports that filename and exits with a nonzero status.

If no `.xlsx` files are selected in the dashboard, the update button stays disabled.

If a workbook sheet lacks a `contents` column during update, that sheet is copied unchanged.

## Testing

Add tests for:

- converting Dua input workbooks into `id`, `language_id`, and `contents`
- converting multiple sheets and skipping rows without `ID` or `Response`
- normalizing wrapped `result` JSON into the Dua target shape
- normalizing direct JSON into the Dua target shape
- wrapping invalid JSON into a valid target shape
- writing selected output workbooks to `DUA CONTENT/UPDATED CONTENT`
- reporting missing selected update files
- dashboard workflow registry containing `dua` and `dua-update-json`
- dashboard command construction for selected Dua files
- `/api/dua-output-files` listing `.xlsx` files from `DUA CONTENT/OUTPUT`

## Out of Scope

- Updating `DUA CONTENT/OUTPUT` in place.
- Selecting individual sheet tabs inside a workbook.
- Adding multiple Dua JSON structures.
- Editing workbook cell values from the dashboard.
