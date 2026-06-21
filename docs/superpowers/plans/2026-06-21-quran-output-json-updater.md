# Quran Output JSON Updater Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a Quran dashboard workflow that rewrites selected output workbook `contents` cells into `juz.json`, `surah.json`, or `page.json` structure and saves copies under `QURAN CONTENT/UPDATED OUTPUT`.

**Architecture:** Extend `QURAN CONTENT/convert_quran_content.py` with reusable JSON normalization and workbook update functions plus a CLI subcommand. Extend `dashboard_server.py` with Quran output file metadata and a `quran-update-json` workflow whose command accepts selected structure/files. Extend the existing dashboard frontend with Quran-only controls for selecting structure and multiple files.

**Tech Stack:** Python standard library (`json`, `argparse`, `unittest`, `pathlib`), `openpyxl`, browser HTML/CSS/JavaScript, existing dashboard subprocess runner.

---

## Task 1: Quran JSON Normalizers

**Files:**
- Modify: `QURAN CONTENT/tests/test_convert_quran_content.py`
- Modify: `QURAN CONTENT/convert_quran_content.py`

- [ ] **Step 1: Write failing tests**

Add tests that call `normalize_quran_content_json(value, structure_type)` for wrapped Juz data, direct Surah data, comma-separated search terms, and invalid JSON.

- [ ] **Step 2: Verify red**

Run: `python -m unittest tests.test_convert_quran_content -v` from `QURAN CONTENT`.

Expected: fails because `normalize_quran_content_json` does not exist.

- [ ] **Step 3: Implement normalizers**

Add `SUPPORTED_JSON_STRUCTURES`, JSON parsing helpers, `_source_content_data`, `_string_list`, and `normalize_quran_content_json`.

- [ ] **Step 4: Verify green**

Run: `python -m unittest tests.test_convert_quran_content -v`.

Expected: Quran tests pass.

## Task 2: Workbook Update Functions and CLI

**Files:**
- Modify: `QURAN CONTENT/tests/test_convert_quran_content.py`
- Modify: `QURAN CONTENT/convert_quran_content.py`

- [ ] **Step 1: Write failing workbook tests**

Add tests for `update_quran_output_files` writing selected workbooks into `UPDATED OUTPUT` and reporting missing selected files.

- [ ] **Step 2: Verify red**

Run: `python -m unittest tests.test_convert_quran_content -v`.

Expected: fails because update functions do not exist.

- [ ] **Step 3: Implement workbook update functions and CLI**

Add `available_quran_output_workbooks`, `update_quran_output_workbook`, `update_quran_output_files`, `print_update_summary`, and `update-output-json` CLI subcommand.

- [ ] **Step 4: Verify green**

Run: `python -m unittest tests.test_convert_quran_content -v`.

Expected: Quran tests pass.

## Task 3: Dashboard Backend Integration

**Files:**
- Modify: `tests/test_dashboard_server.py`
- Modify: `dashboard_server.py`

- [ ] **Step 1: Write failing dashboard tests**

Add tests proving the registry contains `quran-update-json`, command construction includes `update-output-json`, `--structure`, selected `--files`, and API metadata lists Quran output files.

- [ ] **Step 2: Verify red**

Run: `python -m unittest tests.test_dashboard_server -v`.

Expected: fails because the workflow/API additions are missing.

- [ ] **Step 3: Implement backend integration**

Add workflow fields for option-based commands, implement `quran-update-json`, add `api_quran_output_files`, and expose `/api/quran-output-files`.

- [ ] **Step 4: Verify green**

Run: `python -m unittest tests.test_dashboard_server -v`.

Expected: dashboard tests pass.

## Task 4: Dashboard Frontend Controls

**Files:**
- Modify: `dashboard/index.html`
- Modify: `dashboard/styles.css`
- Modify: `dashboard/app.js`

- [ ] **Step 1: Add Quran update controls**

Add structure selector, output file checkbox list, and update button to the dashboard.

- [ ] **Step 2: Wire frontend requests**

Fetch `/api/quran-output-files`, post selected data to `/api/start/quran-update-json`, and poll that workflow status.

- [ ] **Step 3: Verify backend tests still pass**

Run: `python -m unittest tests.test_dashboard_server -v`.

Expected: dashboard tests pass.

## Task 5: Full Verification and Dashboard Restart

**Files:**
- Modify only if verification reveals a defect.

- [ ] **Step 1: Run all relevant tests**

Run dashboard, Quran, and Hadith tests explicitly from their proper working directories.

- [ ] **Step 2: Restart dashboard server**

Stop any old `dashboard_server.py` process on port `8765`, start the updated server, and verify `/`, `/api/workflows`, and `/api/quran-output-files`.

- [ ] **Step 3: Report URL**

Tell the user the dashboard is available at `http://localhost:8765`.
