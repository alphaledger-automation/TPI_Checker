"""
Taxpayer Profile Inquiry — Excel Input Handler
==============================================
Reads registration numbers (NTN / CNIC / Passport No. / Reg-Inc No.) from
an .xlsx file. Auto-detects the registration column or lets the user pick
one manually. The input workbook is only ever read — results go to a
separate report.
"""

import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

import openpyxl

from constants import (
    REG_COLUMN_ALIASES, RES_PENDING, RES_INVALID,
    PARAM_NTN, PARAM_CNIC, PARAM_PASSPORT, PARAM_REGINC, PARAM_AUTO,
)


@dataclass
class Record:
    """One registration number and the profile IRIS returns for it."""
    idx: int                       # 1-based position in the run
    reg_raw: str                   # value exactly as found in Excel
    reg_no: str = ""               # cleaned value typed into IRIS
    param_type: str = ""           # NTN / CNIC / Passport No. / Reg/Inc. No.

    # ── Profile fields ────────────────────────────────────────────────────
    reg_no_shown: str = ""
    reference_no: str = ""
    sales_tax: str = ""            # "Registered for Sales Tax"
    name: str = ""
    category: str = ""
    pp_reg_inc: str = ""
    email: str = ""
    cell: str = ""
    address: str = ""
    registered_on: str = ""
    tax_office: str = ""
    reg_status: str = ""           # "Income Tax: Active , Sales Tax: OPERATIVE"
    branches: List[dict] = field(default_factory=list)

    result: str = RES_PENDING
    pdf_file: str = ""
    timestamp: str = ""


def _open_sheet(filepath: str):
    ext = filepath.rsplit(".", 1)[-1].lower()
    if ext not in ("xlsx", "xlsm"):
        raise ValueError(
            f"Unsupported file format: .{ext}\n"
            "Please save the file as .xlsx and try again."
        )
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    return wb, wb.active


def _header_row(ws) -> List[str]:
    for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        return ["" if c is None else str(c).strip() for c in row]
    return []


def get_column_names(filepath: str) -> List[str]:
    """Return every non-blank column header from the active sheet."""
    wb, ws = _open_sheet(filepath)
    try:
        return [h for h in _header_row(ws) if h]
    finally:
        wb.close()


def detect_reg_column(columns: List[str]) -> Optional[str]:
    """Match headers against known aliases. Returns the header, or None."""
    for col in columns:
        if col.strip().lower() in REG_COLUMN_ALIASES:
            return col
    return None


def classify(value, forced: str = PARAM_AUTO) -> Tuple[str, str]:
    """
    Turn an Excel cell value into (reg_no, param_type).

      7 characters        -> NTN   (digits, or with a letter: E123456)
      8 digits            -> NTN   (trailing check digit dropped)
      5-6 digits          -> NTN   (Excel dropped the leading zeros)
      13 digits           -> CNIC
      9-12 or 14+ digits  -> Reg/Inc. No.  (STRN, incorporation numbers)
      letters + digits    -> Passport No.  (2+ letters) else Reg/Inc. No.

    `forced` overrides the type: the value is then only cleaned, not judged.
    Dashes and spaces are ignored. ("", "") means the cell is unusable.
    """
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    s = str(value).strip()
    s = re.sub(r"\.0+$", "", s)              # "7243972.0" read as text
    clean = re.sub(r"[^A-Za-z0-9]", "", s).upper()
    if not clean:
        return "", ""

    if forced and forced != PARAM_AUTO:
        if forced == PARAM_NTN and clean.isdigit():
            if len(clean) == 8:
                clean = clean[:7]
            elif 5 <= len(clean) <= 7:
                clean = clean.zfill(7)
        return clean, forced

    if clean.isdigit():
        if len(clean) == 13:
            return clean, PARAM_CNIC
        if len(clean) == 8:
            return clean[:7], PARAM_NTN
        if 5 <= len(clean) <= 7:
            return clean.zfill(7), PARAM_NTN
        if len(clean) >= 9:
            return clean, PARAM_REGINC
        return "", ""

    letters = sum(ch.isalpha() for ch in clean)
    if len(clean) in (7, 8) and letters <= 2:
        return clean[:7], PARAM_NTN
    if letters >= 2 and 6 <= len(clean) <= 12:
        return clean, PARAM_PASSPORT
    if 4 <= len(clean) <= 20:
        return clean, PARAM_REGINC
    return "", ""


def read_records(filepath: str, column_name: str,
                 forced: str = PARAM_AUTO) -> List[Record]:
    """
    Read the chosen column and return one Record per UNIQUE registration
    number, in the order they first appear. Blank cells are skipped.
    """
    wb, ws = _open_sheet(filepath)
    try:
        headers = _header_row(ws)
        if column_name not in headers:
            raise ValueError(
                f"Column '{column_name}' not found.\n"
                f"Available: {', '.join(h for h in headers if h)}"
            )
        col = headers.index(column_name)

        records: List[Record] = []
        seen = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if col >= len(row) or row[col] is None:
                continue
            raw = str(row[col]).strip()
            if not raw:
                continue
            reg_no, ptype = classify(row[col], forced)
            key = reg_no or raw
            if key in seen:
                continue
            seen.add(key)
            rec = Record(idx=len(records) + 1, reg_raw=raw,
                         reg_no=reg_no, param_type=ptype)
            if not reg_no:
                rec.result = RES_INVALID
            records.append(rec)
        return records
    finally:
        wb.close()
