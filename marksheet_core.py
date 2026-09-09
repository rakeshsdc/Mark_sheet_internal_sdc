"""
Core logic for the Mark Sheet Generator:
- parse_marks_docx: reads the uploaded .docx (Sl.No, Name, Marks table) into a DataFrame
- compute_marks: scales raw marks to /25 and derives SSA marks from the /25 score
- generate_marksheet: builds the final formatted mark sheet .docx matching the college template
"""

import io
import os
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
LOGO_PATH = os.path.join(ASSETS_DIR, "logo.jpg")


class MarkSheetError(Exception):
    """Raised for problems with the uploaded student list."""
    pass


def parse_marks_docx(file) -> pd.DataFrame:
    """
    Reads an uploaded .docx file containing a 3-column table: Sl.No, Name, Marks.
    Column names are matched case-insensitively and don't need to match exactly
    (e.g. 'Sl.no', 'SL NO', 'Marks Obtained' are all accepted).
    Returns a DataFrame with columns: Sl.No, Name, Marks (float).
    """
    try:
        doc = Document(file)
    except Exception as exc:
        raise MarkSheetError(
            "Could not open the uploaded file. Please make sure it is a valid .docx file."
        ) from exc

    if not doc.tables:
        raise MarkSheetError(
            "No table was found in the uploaded document. "
            "Please upload a .docx file with a table containing Sl.No, Name and Marks columns."
        )

    table = doc.tables[0]
    rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]

    if len(rows) < 2:
        raise MarkSheetError("The table in the uploaded document has no data rows.")

    header = [h.lower() for h in rows[0]]

    def find_col(keywords):
        for i, h in enumerate(header):
            if any(k in h for k in keywords):
                return i
        return None

    idx_sl = find_col(["sl"])
    idx_name = find_col(["name"])
    idx_marks = find_col(["mark"])

    missing = []
    if idx_sl is None:
        missing.append("Sl.No")
    if idx_name is None:
        missing.append("Name")
    if idx_marks is None:
        missing.append("Marks")
    if missing:
        raise MarkSheetError(
            "Could not find the following required column(s) in the uploaded table: "
            + ", ".join(missing)
            + ". Please check the column headers in the uploaded file."
        )

    records = []
    for r in rows[1:]:
        if not any(cell.strip() for cell in r):
            continue  # skip fully blank rows
        name = r[idx_name].strip()
        if not name:
            continue
        raw_marks = r[idx_marks].strip()
        try:
            marks_val = float(raw_marks)
        except ValueError:
            raise MarkSheetError(
                f"Could not read marks for '{name}' (value: '{raw_marks}'). "
                "Marks must be numeric."
            )
        records.append(
            {
                "Sl.No": r[idx_sl].strip(),
                "Name": name,
                "Marks": marks_val,
            }
        )

    if not records:
        raise MarkSheetError("No valid student rows were found in the uploaded table.")

    return pd.DataFrame(records)


def compute_marks(df: pd.DataFrame, max_test_marks: float, ssa_component: float) -> pd.DataFrame:
    """
    Adds two columns to the DataFrame:
    - Marks_25: raw Marks scaled to out of 25, rounded to 1 decimal
    - SSA: (Marks_25 / 25) * ssa_component, rounded to 1 decimal
    """
    if max_test_marks <= 0:
        raise MarkSheetError("Maximum Test Marks must be greater than 0.")

    out = df.copy()
    out["Marks_25"] = (out["Marks"] / max_test_marks * 25).round(1)
    out["SSA"] = (out["Marks_25"] / 25 * ssa_component).round(1)
    return out


def _set_no_borders(table):
    """Remove all borders from a python-docx table (used for the details table)."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        borders.append(el)
    tblPr.append(borders)


def generate_marksheet(
    department: str,
    course_code: str,
    course_name: str,
    max_test_marks: float,
    ssa_component: float,
    df: pd.DataFrame,
    college_name: str = "Sanatana Dharma College, Alappuzha",
    exam_title: str = "Internal Examination, September 2026",
) -> io.BytesIO:
    """
    Builds the formatted mark sheet .docx (matching the college template layout)
    and returns it as an in-memory BytesIO buffer, ready for download.
    """
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)

    # --- Header: logo + college name + exam title ---
    if os.path.exists(LOGO_PATH):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(LOGO_PATH, width=Inches(0.9))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(college_name)
    run.bold = True
    run.font.size = Pt(14)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(exam_title)
    run.font.size = Pt(12)

    doc.add_paragraph()  # spacer

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Mark Sheet")
    run.bold = True
    run.underline = True
    run.font.size = Pt(13)

    doc.add_paragraph()  # spacer

    # --- Details: Department / Course Code / Course Name ---
    details_table = doc.add_table(rows=3, cols=2)
    details_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    _set_no_borders(details_table)
    details_table.columns[0].width = Inches(1.8)
    details_table.columns[1].width = Inches(4.5)

    fields = [
        ("Department", department),
        ("Course Code", course_code),
        ("Course Name", course_name),
    ]
    for i, (label, value) in enumerate(fields):
        left_cell = details_table.cell(i, 0)
        left_cell.text = ""
        left_p = left_cell.paragraphs[0]
        left_run = left_p.add_run(f"{label}\t:")
        left_run.bold = True

        right_cell = details_table.cell(i, 1)
        right_cell.text = str(value) if value else ""

    doc.add_paragraph()  # spacer

    # --- Marks table: Sl.No | Name | Marks | Marks (out of 25) | SSA ---
    marks_table = doc.add_table(rows=1, cols=5)
    marks_table.style = "Table Grid"
    marks_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    widths = [Inches(0.7), Inches(2.3), Inches(1.1), Inches(1.4), Inches(1.4)]
    headers = [
        "Sl.No",
        "Name",
        "Marks",
        "Marks\n(out of 25)",
        f"SSA\n(out of {ssa_component:g})",
    ]
    hdr_cells = marks_table.rows[0].cells
    for i, htext in enumerate(headers):
        hdr_cells[i].width = widths[i]
        hdr_cells[i].text = ""
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(htext)
        run.bold = True

    for _, row in df.iterrows():
        cells = marks_table.add_row().cells
        values = [
            str(row["Sl.No"]),
            str(row["Name"]),
            f"{row['Marks']:.1f}",
            f"{row['Marks_25']:.1f}",
            f"{row['SSA']:.1f}",
        ]
        for i, val in enumerate(values):
            cells[i].width = widths[i]
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i != 1 else WD_ALIGN_PARAGRAPH.LEFT
            p.add_run(val)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf
