"""
Taxpayer Profile Inquiry — Constants & Configuration
====================================================
All URLs, selectors, timing values and column-detection aliases live here
so the rest of the code stays clean.

The differences from a plain ATL lookup are the IRIS tab that
is opened ("Taxpayer Profile Inquiry" instead of "Active Taxpayer List
(Income Tax)"), the extra parameter types (Passport No. / Reg-Inc No.)
and the fact that this form has no date field.
"""

import pathlib

# ── Target URL ────────────────────────────────────────────────────────────────
IRIS_URL = "https://iris.fbr.gov.pk/iris/home.jsf"

# ── DOM Selectors (IRIS 2.0 "Taxpayer Profile Inquiry" form) ─────────────────
SEL_TPI_LINK       = "a:has-text('Taxpayer Profile Inquiry')"
SEL_PARAM_BOX      = "#parameterType"
SEL_PARAM_CURRENT  = "#parameterType .select-box__current"
SEL_PARAM_OPTION   = "#parameterType label.select-box__option"
SEL_REG_INPUT      = "#regNo"
SEL_CAPTCHA_CANVAS = "canvas.captcha-canvas"
SEL_CAPTCHA_INPUT  = "input[placeholder='Enter '], input[placeholder='Enter']"
SEL_CAPTCHA_REFRESH = "button:has(i.fa-refresh)"
SEL_VERIFY_BUTTON  = "button.active-btn-green:has-text('Verify'), button:has-text('Verify')"

# Parameter types offered by this form
PARAM_NTN      = "NTN"
PARAM_CNIC     = "CNIC"
PARAM_PASSPORT = "Passport No."
PARAM_REGINC   = "Reg/Inc. No."
PARAM_TYPES    = (PARAM_NTN, PARAM_CNIC, PARAM_PASSPORT, PARAM_REGINC)
PARAM_AUTO     = "Auto"

# Text that only appears once a profile result is on screen
RESULT_MARKER    = "Tax Office"
RESULT_MARKER_2  = "Reference No"
NO_RECORD_MARKER = "no record exists"

# ── Profile fields read from the result card ─────────────────────────────────
# (key, label as shown on the IRIS card)
PROFILE_FIELDS = [
    ("reg_no_shown",   "Registration No"),
    ("reference_no",   "Reference No"),
    ("sales_tax",      "Registered for Sales Tax"),
    ("name",           "Name"),
    ("category",       "Category"),
    ("pp_reg_inc",     "PP/REG/INC No."),
    ("email",          "Email"),
    ("cell",           "Cell"),
    ("address",        "Address"),
    ("registered_on",  "Registered On"),
    ("tax_office",     "Tax Office"),
    ("reg_status",     "Registration Status"),
]

# ── Timing (milliseconds for Playwright) ─────────────────────────────────────
PAGE_LOAD_TIMEOUT_MS = 60_000
ELEMENT_TIMEOUT_MS   = 20_000

# ── Timing (seconds for Python sleeps) ───────────────────────────────────────
VERIFY_WAIT_SECS     = 10       # how long to wait for a result after Verify
RETRY_DELAY_SECS     = 3
INTER_RECORD_DELAY_SECS = 1     # brief pause between records
MAX_RETRIES          = 2        # whole-record retries (page reload in between)
MAX_CAPTCHA_RETRIES  = 5        # captcha attempts per record try

# ── Registration-number column auto-detection ────────────────────────────────
REG_COLUMN_ALIASES = [
    "registration number", "registration no", "registration no.",
    "registration #", "registration", "reg no", "reg no.", "reg number",
    "reg_no", "regno", "ntn", "cnic", "ntn/cnic", "cnic/ntn",
    "ntn / cnic", "cnic / ntn", "ntn no", "ntn no.", "ntn number",
    "cnic no", "cnic no.", "cnic number",
    "reg/inc no", "reg/inc no.", "reg/inc. no.", "inc no", "inc no.",
    "incorporation no", "incorporation no.", "passport", "passport no",
    "passport no.", "strn", "sales tax registration no",
]

# ── Output ───────────────────────────────────────────────────────────────────
PDF_SUBFOLDER = "Profile_PDFs"


def default_output_dir(excel_path: str = "") -> pathlib.Path:
    """Profile_PDFs next to the chosen Excel file (or in Downloads if none yet)."""
    if excel_path:
        return pathlib.Path(excel_path).resolve().parent / PDF_SUBFOLDER
    return pathlib.Path.home() / "Downloads" / PDF_SUBFOLDER


# ── Record result values ─────────────────────────────────────────────────────
RES_PENDING    = "Pending"
RES_PROCESSING = "Processing"
RES_DOWNLOADED = "Downloaded"
RES_EXISTING   = "Already downloaded"
RES_INVALID    = "Invalid number"
RES_FAILED     = "Failed"            # always followed by " - <reason>"

DONE_RESULTS = (RES_DOWNLOADED, RES_EXISTING)

# ── GUI Appearance ────────────────────────────────────────────────────────────
APP_NAME    = "Taxpayer Profile Inquiry"
APP_VERSION = "1.0"

# Colour palette — red / black / beige. Every entry is (light, dark).
# Light mode sits on warm beige paper with near-black text; dark mode is a
# warm charcoal so the cards lift off the window. The accent is a muted
# brick red in both modes.
C = {
    # surfaces
    "bg800": ("#F5EFE2", "#211A15"),      # window
    "side":  ("#EBE2D0", "#15100C"),      # sidebar
    "bg700": ("#FFFDF7", "#2A211B"),      # cards
    "bg900": ("#FBF6EB", "#181310"),      # inputs, log, table
    "bg600": ("#E3D8C2", "#352A22"),      # table heading
    "bg500": ("#D6C8AD", "#40332A"),
    "div":   ("#D8CBB4", "#3D2F26"),      # card borders
    # brick-red accent
    "a600": ("#EFDAD6", "#6E2B28"),       # selection background
    "a500": ("#9E3029", "#9E3029"),
    "a400": ("#A8392F", "#C4564E"),       # progress bar, checkbox
    "a300": ("#8A2520", "#D9534F"),       # accent text / primary button
    "a200": ("#6E1D19", "#E57373"),
    "acc":  ("#8A2520", "#D9534F"),
    "on_acc": ("#FFF8EF", "#1A1512"),     # text on an accent-filled button
    "sel_fg": ("#1A1512", "#FFF8EF"),     # text on a selected table row
    "seg_sel": ("#E3B7B1", "#9E3029"),    # selected Light/Dark switch segment
    "seg_hover": ("#D7A099", "#A8403A"),
    # warm neutrals
    "sl700": ("#E3DACA", "#2E251E"),      # hover background
    "sl600": ("#C9BCA3", "#453629"),      # input borders, scrollbars
    "sl500": ("#A89878", "#5A4939"),
    "sl400": ("#7E7059", "#7A6554"),      # disabled text
    "sl300": ("#5E5445", "#B09A82"),
    "sl200": ("#4A4238", "#D6C6B0"),      # secondary text
    "sl100": ("#2E2922", "#EDE2D2"),
    # text + status
    "white": ("#1A1512", "#F7EEE2"),      # primary text
    "ok":    ("#2A7030", "#7ACB78"),
    "warn":  ("#9C4A07", "#E6AC4A"),
    "err":   ("#A32019", "#F08A87"),
    "warn_hover": ("#F7EBD6", "#3A2D1C"),
    "err_hover":  ("#F7E0DD", "#3F2320"),
    "log_info": ("#4A4238", "#E3D5BF"), "log_ok": ("#2A7030", "#7ACB78"),
    "log_warn": ("#9C4A07", "#E6AC4A"), "log_err": ("#A32019", "#F08A87"),
}

# The header bar stays near-black in both modes, with a red counter.
HDR = {"bg": "#140F0B", "title": "#F7EEE2", "sub": "#D6C6B0", "count": "#D9534F"}

FONT_UI   = "Segoe UI"
FONT_MONO = "Consolas"
