# Taxpayer Profile Inquiry

Bulk taxpayer profile lookup for Pakistan. Feed a list of NTN / CNIC / Passport / Reg-Inc numbers, get each taxpayer's full IRIS profile as a PDF plus a structured Excel summary.

Reads registration numbers from an Excel file, looks each one up on FBR IRIS under **Online Verifications → Taxpayer Profile Inquiry**, saves the profile card as a PDF, and writes a summary workbook with name, category, tax office, registration status, sales tax registration, address, contact details and every business or branch line.

Browser automation runs headless in the background — it does not take over your mouse or keyboard.

---

## Requirements

- Windows 10/11
- Python 3.10+

Everything else installs on first run. `run.bat` installs Python itself via `winget` if it is missing.

## Install and run

```bash
git clone https://github.com/alphaledger-automation/TPI_Checker.git
cd TPI_Checker
pip install -r requirements.txt
python main.py
```

Or double-click **`run.bat`** on Windows. On first launch it installs what is missing:

- Python packages (`customtkinter`, `playwright`, `openpyxl`, `pillow`, `pytesseract`, `pdfplumber`)
- Playwright's Chromium build (~150 MB, once)
- Tesseract-OCR (Windows may prompt to allow the installer)

Prepare dependencies without opening the window:

```bash
python main.py --setup-only
```

## Input file

- First row must be the column headers.
- The registration column is auto-detected (`Registration Number`, `Registration No`, `NTN`, `CNIC`, …); pick it manually if detection misses.
- Duplicate numbers are looked up once.
- Your input file is never modified — results go to a separate summary workbook.

**Parameter type** is `Auto` by default, classifying each number as:

| Input | Type |
|---|---|
| 7 characters | NTN (digits, or with a letter prefix) |
| 13 digits | CNIC (dashes ignored) |
| 9+ digits | Reg/Inc. No. |
| letters + digits | Passport No. |

If a file holds only one kind of number, set Parameter Type explicitly rather than leaving it on Auto — it is faster and avoids a misclassification being rejected by the form.

## Output

```
<PDF folder>/<RegNo>_<DD-MM-YYYY>.pdf              one profile card per number
<PDF folder>/Profile_Summary_<DD-MM-YYYY>.xlsx     summary workbook
```

**Sheet `Profiles`** — one row per number:

`Length | Registration Number | file name | Status | Name | Reference No | Registered for Sales Tax | Category | PP/REG/INC No. | Email | Cell | Address | Registered On | Tax Office | Registration Status | Income Tax Status | Sales Tax Status | Branches | PDF Status`

**Sheet `Branches`** — one row per business or branch:

`Registration Number | Name | Sr | Business/Branch Name | Business/Branch Address | Principal Activity`

The workbook is created when the run starts and rewritten after every number, so it is always current. Keep it closed in Excel during a run.

Re-running the same day into the same folder resumes: numbers that already have a PDF are reported `Already downloaded` and skipped, unless you tick **Re-download existing PDFs**. The date in the file name is the day the profile was fetched, so a run on another day downloads fresh copies.

## Features

- Dashboard page: file picker, registration-column selector, parameter-type override, PDF output folder
- Current Status page: Done / Total counter and a live row per number with name, tax office, registration status and result
- Pause / Resume and Cancel, applied after the number in flight finishes
- Light / Dark theme, remembered between launches
- Optional visible browser window for diagnosing a failing run
- Captcha retried up to 5 times per number; a failed number stays pending so a second Start retries only what did not succeed
- Incremental Excel writing — a crash or cancel never loses completed rows

## Project layout

| File | Purpose |
|---|---|
| `run.bat` | Windows launcher (finds or installs Python) |
| `main.py` | Entry point; `--setup-only` flag |
| `bootstrap.py` | Dependency check and installer |
| `gui.py` | Dashboard / Current Status window |
| `browser_engine.py` | IRIS automation |
| `excel_handler.py` | Reads the input workbook |
| `report_writer.py` | Writes the Excel summary |
| `constants.py` | URLs, selectors, timings, palette |
| `Assets/` | Application icon |

`USER_GUIDE.txt` holds the end-user walkthrough and troubleshooting notes.

## Notes

- No credentials are stored or required — Taxpayer Profile Inquiry is a public IRIS verification.
- No taxpayer data, sample workbooks or downloaded profiles are included in this repository; `.gitignore` keeps `Profile_PDFs/` and input workbooks out of commits.
- Setup log: `%LocalAppData%\TPI_Checker\setup.log`

## License

MIT — see [LICENSE](LICENSE). Copyright (c) 2026 Alpha Edge Advisory.
