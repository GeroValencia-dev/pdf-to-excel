# PDF Class List → Excel Converter Plan

## Overview
Build a Python script that batch-processes all PDF class lists in a folder,
extracting header metadata and the student table from each, and writes a
structured `.xlsx` file per PDF.

Target format: Saint Francis of Assisi College class list (SFAC layout).

---

## Sub-Task 1 — Project Setup

**Intent:** Establish the Python environment and dependencies.

**Expected Outcomes:**
- `requirements.txt` present with `pdfplumber`, `openpyxl`, `pandas`
- Script file `pdf_to_excel.py` created (empty scaffold)

**Todo List:**
1. Create `requirements.txt` with: `pdfplumber`, `openpyxl`, `pandas`
2. Create `pdf_to_excel.py` with imports and `if __name__ == "__main__":` entry point

**Relevant Context:** Workspace root `c:\Users\gerot\Documents\automate`

**Status:** [ ] pending

---

## Sub-Task 2 — PDF Parser

**Intent:** Extract structured data from a single SFAC-format PDF.

**Expected Outcomes:**
- Function `parse_pdf(path) -> dict` returns:
  - `course_code`, `course_title`, `units`, `schedule`, `class_size`
  - `students`: list of dicts with keys `no`, `name`, `program`, `year_level`
  - `prepared_by`, `date`

**Todo List:**
1. Open PDF with `pdfplumber`, read all pages
2. Use regex to extract header fields from page 1 text
3. Parse student rows (lines 1–30, stopping at "Nothing Follows")
4. Split each row into: number, full name, program, year level
5. Return structured dict

**Relevant Context:**
- PDF text sample in `Culnut B-list.pdf` (lines 8–40 of extracted text)
- Student rows format: `{no}{NAME, FIRSTNAME MIDDLENAME}{PROGRAM}{YEAR}`
- Row numbers are 1–30, year level is concatenated at end (e.g. `BSHMII`)

**Status:** [ ] pending

---

## Sub-Task 3 — Excel Writer

**Intent:** Write the parsed data to a formatted `.xlsx` file.

**Expected Outcomes:**
- One `.xlsx` per PDF with:
  - Row 1–6: header block (school name, course info, schedule)
  - Row 8+: student table with columns: No. | Name | Program | Year Level
  - Bold headers, auto-width columns

**Todo List:**
1. Create function `write_excel(data: dict, out_path: str)`
2. Use `openpyxl` to build workbook
3. Write header metadata rows (merged cells for labels)
4. Write column headers (bold)
5. Write student rows
6. Auto-fit column widths
7. Save file

**Relevant Context:** Output file name = same as PDF file name, `.xlsx` extension

**Status:** [ ] pending

---

## Sub-Task 4 — Batch Runner

**Intent:** Process all PDFs in a given folder.

**Expected Outcomes:**
- Script accepts a folder path as CLI argument (defaults to current folder)
- Loops over all `*.pdf` files, calls parse + write per file
- Prints status per file (success / skipped / error)

**Todo List:**
1. Use `argparse` for input folder and optional output folder args
2. Glob all `*.pdf` in input folder
3. For each PDF: call `parse_pdf()` → `write_excel()` → print result
4. Handle errors per file gracefully (log and continue)

**Relevant Context:** `pdf_to_excel.py` main entry point

**Status:** [ ] pending

---

## Sub-Task 5 — Smoke Test

**Intent:** Verify the script works on the sample PDF.

**Expected Outcomes:**
- Running `python pdf_to_excel.py .` produces `Culnut B-list.xlsx`
- Excel has correct header info and all 30 students

**Todo List:**
1. Run script against workspace folder
2. Open output Excel and verify header fields and student count
3. Fix any parsing edge cases (e.g. `CAñADA`, `N/A` in name field)

**Relevant Context:** `Culnut B-list.pdf` — 30 students, page 2 is "Nothing Follows" page

**Status:** [ ] pending
