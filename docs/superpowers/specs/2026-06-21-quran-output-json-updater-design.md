# Quran Output JSON Updater Design

## Goal

Add a Quran output update workflow that rewrites the `contents` column in selected Quran output workbooks into one of three approved JSON structures.

The source files are selected from:

`QURAN CONTENT/OUTPUT`

Updated copies are written to:

`QURAN CONTENT/UPDATED OUTPUT`

The existing files in `QURAN CONTENT/OUTPUT` are not modified.

## Target Structures

The workflow supports three output structure names:

- `juz.json`
- `surah.json`
- `page.json`

`juz.json` and `page.json` use this top-level shape:

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
  "searching_terms": []
}
```

`surah.json` uses this top-level shape:

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
  "lessons": [],
  "faqs": []
}
```

## Current Input Shapes

Existing `contents` cells may contain more than one JSON shape.

Some rows contain direct JSON with keys like `meta`, `heading`, `lessons`, and `faqs`.

Other rows contain a wrapped list shape:

```json
[
  {
    "keyword": "al quran juz 1",
    "result": {
      "meta": {},
      "heading": "Al Quran Juz 1",
      "summary": "...",
      "searching_terms": "term one, term two"
    }
  }
]
```

The updater normalizes both forms. For wrapped rows, it uses the first object in the list and then uses that object's `result` value as the content source.

## Field Mapping

The updater keeps useful data from the source JSON and reshapes it.

For all structures:

- `meta.title` comes from source `meta.title`
- `meta.description` comes from source `meta.description`
- `heading.title` comes from source `heading.title`, source string `heading`, or source `title`
- `heading.description` comes from source `heading.description`, source `summary`, source `description`, or an empty string

For `juz.json` and `page.json`:

- `searching_terms` comes from source `searching_terms`
- if `searching_terms` is a comma-separated string, split it into a list
- if `searching_terms` is missing, use an empty list

For `surah.json`:

- `lessons` comes from source `lessons`; if missing, use an empty list
- `faqs` comes from source `faqs`; if missing, use an empty list

The output JSON is written compactly with Unicode preserved.

## Script Changes

Update `QURAN CONTENT/convert_quran_content.py` with reusable functions:

- `available_quran_output_workbooks(output_dir)`
- `normalize_quran_content_json(value, structure_type)`
- `update_quran_output_workbook(input_path, output_path, structure_type)`
- `update_quran_output_files(output_dir, updated_dir, structure_type, filenames)`

The existing Quran conversion behavior remains unchanged. The new updater is additive.

The CLI gains an update mode:

```powershell
python convert_quran_content.py update-output-json --structure juz.json --files "EN_Quran Juz Content (21-6-26).xlsx"
```

If no files are provided in update mode, the script updates all `.xlsx` files found in `QURAN CONTENT/OUTPUT`.

## Dashboard Changes

The dashboard Quran section gains a second action area named `Update Output JSON`.

Controls:

- structure selector with `juz.json`, `surah.json`, and `page.json`
- checkbox list of files from `QURAN CONTENT/OUTPUT`
- action button to update the selected files

The dashboard sends selected filenames and structure type to the local server. The server runs the Quran script update mode as a subprocess and streams logs using the same job infrastructure as the existing run buttons.

## Data Flow

1. Dashboard requests available Quran output files.
2. Server scans `QURAN CONTENT/OUTPUT` for `.xlsx` files.
3. User selects structure type and one or more files.
4. Dashboard starts the `quran-update-json` workflow with selected options.
5. Server runs `convert_quran_content.py update-output-json`.
6. Script reads each selected workbook from `OUTPUT`.
7. Script writes transformed copies to `UPDATED OUTPUT`.
8. Dashboard shows logs and generated output paths.

## Error Handling

If a selected file does not exist under `OUTPUT`, the script reports that file and exits with a nonzero status.

If a `contents` cell contains invalid JSON, the updater still writes a valid target JSON object with empty optional fields and places the original text in `heading.description`.

If a row has no `contents` column, that sheet is copied with headers and rows unchanged.

If no `.xlsx` files are selected in the dashboard, the button stays disabled.

## Testing

Add unit tests for:

- direct `surah.json` data is kept in the requested `surah.json` shape
- wrapped `result` data is converted to `juz.json`
- comma-separated `searching_terms` becomes a list
- invalid JSON is wrapped into a valid target shape
- selected output workbooks are written to `UPDATED OUTPUT`
- missing selected files cause a clear error
- dashboard server command construction includes structure type and selected filenames

## Out of Scope

- Editing Quran input workbooks.
- Updating `QURAN CONTENT/OUTPUT` in place.
- Adding custom user-defined JSON schemas.
- Automatically guessing the structure type from filename when the user explicitly selects a structure.
