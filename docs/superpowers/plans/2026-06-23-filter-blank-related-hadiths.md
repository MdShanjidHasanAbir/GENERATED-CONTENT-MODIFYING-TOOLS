# Filter Blank Related Hadiths Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Remove rows with blank normalized `related_hadiths` from `BOOK WISE FINAL UPDATED CONTENT` only, and add their original mapped `INPUT` rows to `missing_input`.

**Architecture:** Extend the updated-content generation path so it can collect skipped final rows as missing hadith IDs. Reuse existing `MissingId`, workbook pairing, and `write_missing_input_workbooks` behavior so the generated `_missing.xlsx` files keep the existing `id`, `arabic`, `translation` shape.

**Tech Stack:** Python, `unittest`, `openpyxl`, existing workbook helpers in `SINGLE HADITH CONTENT/reconcile_hadith_outputs.py`.

---

### Task 1: Regression Test

**Files:**
- Modify: `tests/test_hadith_selective_parsing.py`

- [ ] **Step 1: Write failing test**

Add a test that builds a temporary `INPUT`, `BOOK WISE FINAL`, `BOOKS.xlsx`, and `missing_input`. The final workbook contains one row with a valid related hadith and one row with empty `related_hadiths`. The expected result is that updated content contains only the valid row and missing input contains the original mapped keyword row for the blank-related hadith ID.

- [ ] **Step 2: Run targeted test**

Run: `python -B -m unittest tests.test_hadith_selective_parsing.HadithSelectiveParsingTest.test_updated_content_skips_blank_related_hadiths_and_adds_missing_input`
Expected: FAIL because `write_updated_content_workbooks` currently writes blank-related rows and does not update `missing_input`.

### Task 2: Implementation

**Files:**
- Modify: `SINGLE HADITH CONTENT/reconcile_hadith_outputs.py`

- [ ] **Step 1: Add optional input and missing paths**

Add optional `input_dir` and `missing_input_dir` parameters to `write_updated_content_workbooks`. Existing callers continue working when these are omitted.

- [ ] **Step 2: Collect skipped IDs**

When `_write_updated_content_workbook` converts a row and normalized `related_hadiths` is empty, skip writing the row and return enough metadata to map it back to the input workbook: language, collection, and hadith ID.

- [ ] **Step 3: Append to missing input generation**

Use existing workbook pairing helpers to resolve skipped IDs back to the original `INPUT` workbook, then call `write_missing_input_workbooks` with those `MissingId` entries. Do this only when both `input_dir` and `missing_input_dir` are provided.

### Task 3: Verification

**Files:**
- Test: `tests/test_hadith_selective_parsing.py`
- Test: `tests/test_hadith_output_filtering.py`

- [ ] **Step 1: Run targeted test**

Run: `python -B -m unittest tests.test_hadith_selective_parsing.HadithSelectiveParsingTest.test_updated_content_skips_blank_related_hadiths_and_adds_missing_input`
Expected: PASS.

- [ ] **Step 2: Run full tests**

Run: `python -B -m unittest discover -s tests`
Expected: PASS.
