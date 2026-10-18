"""
pdf_to_excel.py
---------------
Batch-converts SFAC-format class-list PDFs to structured .xlsx files.

Usage:
    python pdf_to_excel.py [input_folder] [--out output_folder]
"""

import re
import sys
import argparse
from pathlib import Path

import pdfplumber
import openpyxl
from openpyxl.styles import Font, Alignment


# ---------------------------------------------------------------------------
# Sub-Task 2 — PDF Parser
# ---------------------------------------------------------------------------

# Regex patterns for header fields (all on one fused line in pdfplumber output)
_RE_COURSE_CODE  = re.compile(r"Course Code:\s*(.+?)Course Title:", re.IGNORECASE)
_RE_COURSE_TITLE = re.compile(r"Course Title:\s*(.+?)No\. of Units:", re.IGNORECASE)
_RE_UNITS        = re.compile(r"No\. of Units:\s*(\S+)", re.IGNORECASE)
_RE_SCHEDULE     = re.compile(r"Schedule \(Time/Day/Room\):\s*(.+?)Class Size:", re.IGNORECASE)
_RE_CLASS_SIZE   = re.compile(r"Class Size:\s*(\d+)", re.IGNORECASE)
_RE_SEMESTER     = re.compile(r"(First|Second|Third|Summer)\s+Semester\s+\d{4}-\d{4}", re.IGNORECASE)
_RE_PREPARED_BY  = re.compile(r"^(.+?)\s*(\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4})$")

# Student row: leading number, then everything up to the program code + year
# Program codes end in a Roman-numeral-style suffix (I, II, III, IV, V) or just a number
_RE_STUDENT_ROW  = re.compile(
    r"^(\d{1,2})\s*(.+?)\s+(BSH?[A-Z]+(?:\s?[IVX]+|\d+|N/A)?)$"
)

# Alternate: split trailing program+year from name
# e.g. "1ALAVERA, ISHIE SHANAIA DOLLETEBSHMII"
#      number | NAME | PROGRAM+YEAR_LEVEL
_RE_STUDENT_FUSED = re.compile(
    r"^(\d{1,2})([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚa-záéíóúñ,\.\s/\-']+?)\s*((?:BS|AB|BT|BPA|BS)[A-Z]+(?:[IVX]+|\d+)?)$"
)


def _extract_header(text: str) -> dict:
    """Pull course metadata from the fused header text of page 1."""
    data = {}
    m = _RE_COURSE_CODE.search(text)
    data["course_code"] = m.group(1).strip() if m else ""
    m = _RE_COURSE_TITLE.search(text)
    data["course_title"] = m.group(1).strip() if m else ""
    m = _RE_UNITS.search(text)
    data["units"] = m.group(1).strip() if m else ""
    m = _RE_SCHEDULE.search(text)
    data["schedule"] = m.group(1).strip() if m else ""
    m = _RE_CLASS_SIZE.search(text)
    data["class_size"] = m.group(1).strip() if m else ""
    m = _RE_SEMESTER.search(text)
    data["semester"] = m.group(0).strip() if m else ""
    return data



def _split_prog_year(prog_year: str):
    """
    Split a program+year string into (program, year_level).
    Handles both fused ('BSHMII') and space-separated ('BSHM II') forms.
    """
    prog_year = prog_year.strip()
    # Space-separated: 'BSHM II', 'BSHM I', 'BS HRM II'
    m = re.match(r"^([A-Z][A-Z\s]+?)\s+((?:[IVX]+|\d+))$", prog_year)
    if m:
        return m.group(1).strip(), m.group(2)
    # Fused: 'BSHMII', 'BSHMI'
    m = re.match(r"^([A-Z]+?)((?:[IVX]+|\d+))$", prog_year)
    if m:
        return m.group(1), m.group(2)
    return prog_year, "N/A"



def _parse_student_line(line: str):
    """
    Parse a single student line into (no, name, program, year_level).
    Returns None if the line does not look like a student row.
    """
    line = line.strip()
    if not line:
        return None
    # Lines that are obviously not students
    if line.startswith("*") or "Nothing Follows" in line:
        return None
    if re.match(r"^(Name|Prepared|Professor|Verified|Checked|Received|Record|\(Sig|Dr\.|Campus|College)", line, re.I):
        return None

    # Try the fused regex (no spaces between fields, common in pdfplumber output)
    m = _RE_STUDENT_FUSED.match(line)
    if m:
        no        = m.group(1).strip()
        full_name = m.group(2).strip().rstrip(",").strip()
        prog_year = m.group(3).strip()
        program, year_level = _split_prog_year(prog_year)
        return {"no": no, "name": full_name, "program": program, "year_level": year_level}

    # Fallback: spaced regex
    m = _RE_STUDENT_ROW.match(line)
    if m:
        no        = m.group(1).strip()
        full_name = m.group(2).strip()
        prog_year = m.group(3).strip()
        program, year_level = _split_prog_year(prog_year)
        return {"no": no, "name": full_name, "program": program, "year_level": year_level}

    return None


def parse_pdf(path: str) -> dict:
    """
    Extract structured data from a single SFAC-format class-list PDF.
    Returns a dict with course metadata and a list of student dicts.
    """
    students = []
    header_data = {}
    prepared_by = ""
    prep_date   = ""

    with pdfplumber.open(path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            lines = text.splitlines()

            if page_idx == 0:
                # Header is on the first page — join all lines for regex search
                full_text = " ".join(lines)
                header_data = _extract_header(full_text)

            for line in lines:
                student = _parse_student_line(line)
                if student:
                    # Avoid duplicates (page 2 repeats header only, not students)
                    if not any(s["no"] == student["no"] for s in students):
                        students.append(student)

                # Prepared by line — name and date appear on same line as the label
                if not prepared_by:
                    m = _RE_PREPARED_BY.match(line.strip())
                    if m:
                        raw = m.group(1).strip()
                        # Strip leading "Prepared by:" label if present
                        raw = re.sub(r"^Prepared\s+by:\s*", "", raw, flags=re.I)
                        prepared_by = raw.strip()
                        prep_date   = m.group(2).strip()

    return {
        **header_data,
        "students":    students,
        "prepared_by": prepared_by,
        "date":        prep_date,
    }


# ---------------------------------------------------------------------------
# Sub-Task 3 — Excel Writer
# ---------------------------------------------------------------------------

def write_excel(data: dict, out_path: str):
    """Write parsed class-list data to a formatted .xlsx file."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Class List"

    bold = Font(bold=True)

    def cell(row, col, value, bold_=False, align="left"):
        c = ws.cell(row=row, column=col, value=value)
        if bold_:
            c.font = bold
        c.alignment = Alignment(horizontal=align, wrap_text=True)
        return c

    # --- Header block (rows 1-7) ---
    cell(1, 1, "Saint Francis of Assisi College", bold_=True, align="center")
    ws.merge_cells("A1:D1")
    cell(2, 1, "96 Bayanan, City of Bacoor, Cavite", align="center")
    ws.merge_cells("A2:D2")
    cell(3, 1, "COLLEGE DEPARTMENT", align="center")
    ws.merge_cells("A3:D3")
    cell(4, 1, "CLASS LIST", bold_=True, align="center")
    ws.merge_cells("A4:D4")
    cell(5, 1, data.get("semester", ""), align="center")
    ws.merge_cells("A5:D5")

    cell(6, 1, f"Course Code: {data.get('course_code', '')}")
    cell(6, 2, f"Course Title: {data.get('course_title', '')}")
    cell(6, 3, f"No. of Units: {data.get('units', '')}")

    cell(7, 1, f"Schedule: {data.get('schedule', '')}")
    cell(7, 2, f"Class Size: {data.get('class_size', '')}")

    # --- Column headers (row 9) ---
    headers = ["No.", "Name", "Program", "Year Level"]
    for col, h in enumerate(headers, start=1):
        cell(9, col, h, bold_=True, align="center")

    # --- Student rows ---
    for i, s in enumerate(data["students"], start=10):
        cell(i, 1, s["no"],         align="center")
        cell(i, 2, s["name"])
        cell(i, 3, s["program"],    align="center")
        cell(i, 4, s["year_level"], align="center")

    # --- Footer ---
    footer_row = 10 + len(data["students"]) + 1
    cell(footer_row,     1, "Prepared by:")
    cell(footer_row + 1, 1, data.get("prepared_by", ""))
    cell(footer_row + 2, 1, data.get("date", ""))

    # --- Auto-fit column widths ---
    col_widths = [6, 45, 12, 14]
    for col_idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

    wb.save(out_path)


# ---------------------------------------------------------------------------
# Sub-Task 4 — Batch Runner
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Convert SFAC PDF class lists to Excel.")
    parser.add_argument(
        "input_folder",
        nargs="?",
        default=".",
        help="Folder containing PDF files (default: current directory)",
    )
    parser.add_argument(
        "--out",
        default=None,
        help="Output folder for .xlsx files (default: same as input folder)",
    )
    args = parser.parse_args()

    input_dir  = Path(args.input_folder).resolve()
    output_dir = Path(args.out).resolve() if args.out else input_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    pdf_files = sorted(input_dir.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found in: {input_dir}")
        sys.exit(0)

    for pdf_path in pdf_files:
        out_path = output_dir / (pdf_path.stem + ".xlsx")
        try:
            data = parse_pdf(str(pdf_path))
            write_excel(data, str(out_path))
            student_count = len(data["students"])
            print(f"[OK]      {pdf_path.name}  ->  {out_path.name}  ({student_count} students)")
        except Exception as exc:
            print(f"[ERROR]   {pdf_path.name}  ->  {exc}")


if __name__ == "__main__":
    main()
