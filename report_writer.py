"""
Taxpayer Profile Inquiry — Summary File Writer
==============================================
Maintains Profile_Summary_<date>.xlsx in the PDF folder.

Sheet "Profiles"  — one row per registration number:
  Length | Registration Number | file name | Status | Name | Reference No |
  Registered for Sales Tax | Category | PP/REG/INC No. | Email | Cell |
  Address | Registered On | Tax Office | Registration Status |
  Income Tax Status | Sales Tax Status | Branches | PDF Status

Sheet "Branches"  — one row per business / branch line of the profile card:
  Registration Number | Name | Sr | Business/Branch Name |
  Business/Branch Address | Principal Activity

The file is created when a run starts (every row "Not-Downloaded") and
re-saved after every record, so it is always up to date.
"""

import pathlib
import re
from typing import List

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from constants import DONE_RESULTS, RES_PENDING, RES_PROCESSING, RES_FAILED

# ── Styles ────────────────────────────────────────────────────────────────────
_HEADER_FILL = PatternFill("solid", fgColor="1A5530")
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_HEADER_ALIGN = Alignment(horizontal="center", vertical="center")

_HEADERS = [
    "Length", "Registration Number", "file name", "Status", "Name",
    "Reference No", "Registered for Sales Tax", "Category", "PP/REG/INC No.",
    "Email", "Cell", "Address", "Registered On", "Tax Office",
    "Registration Status", "Income Tax Status", "Sales Tax Status",
    "Branches", "PDF Status",
]
_COL_WIDTHS = [9, 22, 30, 26, 34, 16, 24, 40, 16, 24, 18, 44, 15, 20,
               34, 18, 18, 10, 20]

_BRANCH_HEADERS = ["Registration Number", "Name", "Sr", "Business/Branch Name",
                   "Business/Branch Address", "Principal Activity"]
_BRANCH_WIDTHS = [22, 34, 6, 34, 44, 48]

STATUS_DOWNLOADED     = "Downloaded"
STATUS_NOT_DOWNLOADED = "Not-Downloaded"
PDF_PRESENT = "Saved in Profile_PDFs"
PDF_ABSENT  = "Not in Profile_PDFs"


def report_path(output_dir: pathlib.Path, date_for_file: str) -> pathlib.Path:
    return output_dir / f"Profile_Summary_{date_for_file}.xlsx"


def _status(result: str) -> str:
    if result in DONE_RESULTS:
        return STATUS_DOWNLOADED
    if result in (RES_PENDING, RES_PROCESSING):
        return STATUS_NOT_DOWNLOADED
    if result.startswith(RES_FAILED):
        return "Error" + result[len(RES_FAILED):]        # "Error - <reason>"
    return result                                        # "Invalid number"


def split_reg_status(reg_status: str):
    """'Income Tax: Active , Sales Tax: OPERATIVE' -> ('Active', 'OPERATIVE')."""
    income = sales = ""
    m = re.search(r"income\s*tax\s*:\s*([^,;]*)", reg_status, re.I)
    if m:
        income = m.group(1).strip()
    m = re.search(r"sales\s*tax\s*:\s*([^,;]*)", reg_status, re.I)
    if m:
        sales = m.group(1).strip()
    return income, sales


def _write_header(ws, headers, widths):
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font      = _HEADER_FONT
        cell.fill      = _HEADER_FILL
        cell.alignment = _HEADER_ALIGN
    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def write_report(records: List, output_dir: pathlib.Path,
                 date_for_file: str) -> str:
    """Write the summary workbook and return its path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Profiles"
    _write_header(ws, _HEADERS, _COL_WIDTHS)

    for row, rec in enumerate(records, start=2):
        reg = rec.reg_no or rec.reg_raw
        file_name = f"{rec.reg_no}_{date_for_file}.pdf" if rec.reg_no else ""
        on_disk = bool(file_name) and (output_dir / file_name).exists()
        income, sales = split_reg_status(rec.reg_status)

        ws.cell(row=row, column=1, value=f"=LEN(B{row})")
        # Registration numbers stay text: keeps leading zeros (0123456),
        # letters (E123456) and stops CNICs turning into 3.52E+12.
        reg_cell = ws.cell(row=row, column=2, value=str(reg))
        reg_cell.number_format = "@"
        ws.cell(row=row, column=3, value=file_name)
        ws.cell(row=row, column=4, value=_status(rec.result))
        ws.cell(row=row, column=5, value=rec.name or None)
        ws.cell(row=row, column=6, value=rec.reference_no or None)
        ws.cell(row=row, column=7, value=rec.sales_tax or None)
        ws.cell(row=row, column=8, value=rec.category or None)
        ws.cell(row=row, column=9, value=rec.pp_reg_inc or None)
        ws.cell(row=row, column=10, value=rec.email or None)
        cell_cell = ws.cell(row=row, column=11, value=rec.cell or None)
        cell_cell.number_format = "@"
        ws.cell(row=row, column=12, value=rec.address or None)
        ws.cell(row=row, column=13, value=rec.registered_on or None)
        ws.cell(row=row, column=14, value=rec.tax_office or None)
        ws.cell(row=row, column=15, value=rec.reg_status or None)
        ws.cell(row=row, column=16, value=income or None)
        ws.cell(row=row, column=17, value=sales or None)
        ws.cell(row=row, column=18, value=len(rec.branches) or None)
        ws.cell(row=row, column=19, value=PDF_PRESENT if on_disk else PDF_ABSENT)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(_HEADERS))}{max(len(records) + 1, 2)}"

    # ── Branches sheet ────────────────────────────────────────────────────
    bs = wb.create_sheet("Branches")
    _write_header(bs, _BRANCH_HEADERS, _BRANCH_WIDTHS)
    row = 2
    for rec in records:
        for br in rec.branches:
            reg_cell = bs.cell(row=row, column=1, value=str(rec.reg_no or rec.reg_raw))
            reg_cell.number_format = "@"
            bs.cell(row=row, column=2, value=rec.name or None)
            bs.cell(row=row, column=3, value=br.get("sr") or None)
            bs.cell(row=row, column=4, value=br.get("name") or None)
            bs.cell(row=row, column=5, value=br.get("address") or None)
            bs.cell(row=row, column=6, value=br.get("activity") or None)
            row += 1
    bs.freeze_panes = "A2"
    bs.auto_filter.ref = f"A1:F{max(row - 1, 2)}"

    path = report_path(output_dir, date_for_file)
    wb.save(str(path))
    return str(path)
