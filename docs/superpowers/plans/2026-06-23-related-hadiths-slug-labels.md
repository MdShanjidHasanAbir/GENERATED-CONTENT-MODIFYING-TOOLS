# Related Hadiths Slug Labels Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add `slug` to every normalized `related_hadiths` item and make recognized book labels language-aware.

**Architecture:** Keep the change inside the existing final content converter in `SINGLE HADITH CONTENT/reconcile_hadith_outputs.py`. Reuse collection detection and catalog-based book id lookup, then derive labels and slugs from the recognized collection.

**Tech Stack:** Python, `unittest`, `openpyxl`, JSON normalization helpers already in the repository.

---

### Task 1: Regression Tests

**Files:**
- Modify: `tests/test_hadith_selective_parsing.py`

- [ ] **Step 1: Write failing tests**

Add assertions that converted `related_hadiths` include `slug`, that Bengali rows emit Bengali labels for recognized collections, and that non-core catalog collections receive generated slugs.

- [ ] **Step 2: Run tests to verify failure**

Run: `python -m unittest tests.test_hadith_selective_parsing.HadithSelectiveParsingTest`
Expected: FAIL because current output has no `slug` and still uses English labels.

### Task 2: Converter Update

**Files:**
- Modify: `SINGLE HADITH CONTENT/reconcile_hadith_outputs.py`

- [ ] **Step 1: Add slug helpers**

Add explicit slugs for BUKHARI, MUSLIM, MAJAH, DAWUD, TIRMIDHI, and NASAI. Add fallback slug generation from the normalized collection key for all other books.

- [ ] **Step 2: Add localized label helpers**

Add per-language labels for the six recognized books and preserve the existing English fallback for collections without a localized label.

- [ ] **Step 3: Include fields during normalization**

Update both object and text normalization paths to return `{"id", "book_id", "slug", "label"}`.

### Task 3: Verification

**Files:**
- Test: `tests/test_hadith_selective_parsing.py`

- [ ] **Step 1: Run targeted tests**

Run: `python -m unittest tests.test_hadith_selective_parsing.HadithSelectiveParsingTest`
Expected: PASS.

- [ ] **Step 2: Run broader repository tests if available**

Run: `python -m unittest discover -s tests`
Expected: PASS.
