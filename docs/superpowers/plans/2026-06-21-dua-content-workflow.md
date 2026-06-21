# Dua Content Workflow Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a dashboard-controlled Dua workflow that converts `DUA CONTENT/INPUT` workbooks to `DUA CONTENT/OUTPUT`, then updates selected output files into the approved Dua JSON shape under `DUA CONTENT/UPDATED CONTENT`.

**Architecture:** Create a standalone `DUA CONTENT/convert_dua_content.py` modeled on the Quran converter, with tested conversion and JSON update functions. Extend the dashboard server with `dua` and `dua-update-json` workflows plus a file-list API. Add Dua-specific frontend controls while keeping the existing operator console layout.

**Tech Stack:** Python standard library, `openpyxl`, `unittest`, HTML/CSS/JavaScript, existing dashboard subprocess runner.

---

## Task 1: Dua Converter Script

**Files:**
- Create: `DUA CONTENT/tests/test_convert_dua_content.py`
- Create: `DUA CONTENT/tests/__init__.py`
- Create: `DUA CONTENT/convert_dua_content.py`

- [ ] Write failing tests for converting input workbooks, JSON normalization, updated output copies, and missing selected files.
- [ ] Run `python -m unittest tests.test_convert_dua_content -v` from `DUA CONTENT` and confirm failure.
- [ ] Implement `convert_dua_content.py`.
- [ ] Run the test again and confirm pass.

## Task 2: Dashboard Backend

**Files:**
- Modify: `tests/test_dashboard_server.py`
- Modify: `dashboard_server.py`

- [ ] Write failing tests for `dua`, `dua-update-json`, selected file command construction, and `/api/dua-output-files`.
- [ ] Run `python -m unittest tests.test_dashboard_server -v` and confirm failure.
- [ ] Implement the dashboard registry and API changes.
- [ ] Run the test again and confirm pass.

## Task 3: Dashboard Frontend

**Files:**
- Modify: `dashboard/index.html`
- Modify: `dashboard/styles.css`
- Modify: `dashboard/app.js`

- [ ] Add a Dua update panel with file checkboxes and `Update selected JSON`.
- [ ] Fetch `/api/dua-output-files` and post to `/api/start/dua-update-json`.
- [ ] Keep the panel visible only for the `Dua Content` workflow.
- [ ] Run dashboard tests again.

## Task 4: Verification and Restart

**Files:**
- Modify only if verification exposes a defect.

- [ ] Run dashboard tests.
- [ ] Run Quran tests.
- [ ] Run Hadith tests.
- [ ] Run Dua tests.
- [ ] Restart the dashboard on `http://localhost:8765`.
- [ ] Verify `/`, `/api/workflows`, and `/api/dua-output-files` return `200`.
