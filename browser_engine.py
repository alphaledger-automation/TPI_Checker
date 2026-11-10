"""
Taxpayer Profile Inquiry — Playwright Browser Engine
====================================================
Handles all interaction with the FBR IRIS "Taxpayer Profile Inquiry"
verification form:
  1. Launch Chromium (window kept off-screen) → open IRIS → click
     "Taxpayer Profile Inquiry".
  2. For each registration number: select the parameter type (NTN / CNIC /
     Passport No. / Reg-Inc No.), fill the number, read the captcha, click
     Verify, read the profile card and save it as a one-page PDF.
  3. Retry failed records up to MAX_RETRIES times.
  4. Communicate progress/logs back to the GUI via a queue.

Designed to run entirely inside a background daemon thread.

Note: IRIS refuses headless browsers ("Web Page Blocked"), so a normal
Chromium window is used and simply positioned off-screen. It never needs
the mouse, keyboard or focus, so the PC stays usable during a run.

Unlike the Active Taxpayer List form, this form has NO date field — the
profile is whatever IRIS holds today. The run date is still used in the
PDF / summary file names so each day's download stands on its own.
"""

import base64
import datetime
import io
import json
import os
import pathlib
import queue
import re
import threading
import time
from typing import List, Optional

import pdfplumber
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
from playwright.sync_api import sync_playwright, Page, BrowserContext, Browser

from bootstrap import find_tesseract
from constants import (
    IRIS_URL,
    SEL_TPI_LINK, SEL_PARAM_BOX, SEL_PARAM_CURRENT, SEL_PARAM_OPTION,
    SEL_REG_INPUT, SEL_CAPTCHA_CANVAS, SEL_CAPTCHA_INPUT,
    SEL_CAPTCHA_REFRESH, SEL_VERIFY_BUTTON,
    RESULT_MARKER, RESULT_MARKER_2, NO_RECORD_MARKER, PROFILE_FIELDS,
    PAGE_LOAD_TIMEOUT_MS, ELEMENT_TIMEOUT_MS,
    VERIFY_WAIT_SECS, RETRY_DELAY_SECS, INTER_RECORD_DELAY_SECS,
    MAX_RETRIES, MAX_CAPTCHA_RETRIES,
    RES_PENDING, RES_PROCESSING, RES_DOWNLOADED, RES_EXISTING,
    RES_INVALID, RES_FAILED,
)
from report_writer import write_report

_tess = find_tesseract()
if _tess:
    pytesseract.pytesseract.tesseract_cmd = _tess


# ── Helper utilities ─────────────────────────────────────────────────────────

def _now() -> str:
    return datetime.datetime.now().strftime("%H:%M:%S")


def _now_iso() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _empty_profile() -> dict:
    out = {key: "" for key, _ in PROFILE_FIELDS}
    out["branches"] = []
    out["no_record"] = False
    return out


# Labels as they appear on the card, longest first so "Registration Status"
# is never swallowed by "Registration No".
_LABELS = sorted((label for _, label in PROFILE_FIELDS), key=len, reverse=True)
_LABEL_TO_KEY = {label: key for key, label in PROFILE_FIELDS}
def _label_pattern(label: str) -> str:
    """Regex for one label, tolerant of the line breaks IRIS prints."""
    pat = re.escape(label).replace(r"\ ", r"\s+")
    if label == "Registered for Sales Tax":
        # The card wraps this one around its own value:
        #   "Registered for Sales" / "No" / "Tax"
        pat = r"Registered\s+for\s+Sales(?:\s+Tax)?"
    return pat


# Labels are matched anywhere on a line, not only at its start: a PDF of
# the card re-flows several fields onto one line. Longest first, so
# "Registration Status" is never swallowed by "Registration No".
_LABELS = sorted((label for _, label in PROFILE_FIELDS), key=len, reverse=True)
_LABEL_TO_KEY = {label: key for key, label in PROFILE_FIELDS}
_LABEL_ALT = "|".join(_label_pattern(l) for l in _LABELS)
_FIELD_RE = re.compile(r"(?<![A-Za-z])(" + _LABEL_ALT + r")\s*:?\s*", re.I)

# Everything from here on belongs to the business / branch table.
_BRANCH_CUT_RE = re.compile(
    r"(?:(?m:^)\s*Sr\.?\s*$|Business\s*/?\s*Branch|Principal\s+Activity)", re.I)


def parse_result_text(text: str) -> dict:
    """
    Extract the profile fields from the text of the IRIS result card (or of
    a saved PDF). Used as a fallback when the DOM reader finds nothing.
    """
    out = _empty_profile()
    if not text:
        return out
    if NO_RECORD_MARKER in text.lower():
        out["no_record"] = True
        out["reg_status"] = "No record exists"
        return out

    # A PDF breaks long values across lines, sometimes after a hyphen
    # ("23-MAY-\n05"). Join those back up before anything else.
    text = re.sub(r"-[ \t]*\n[ \t]*", "-", text)
    head = _BRANCH_CUT_RE.split(text, maxsplit=1)[0]

    matches = list(_FIELD_RE.finditer(head))
    for i, m in enumerate(matches):
        label = re.sub(r"\s+", " ", m.group(1)).strip()
        key = _key_for_label(label)
        if not key:
            continue
        end = matches[i + 1].start() if i + 1 < len(matches) else len(head)
        value = re.sub(r"\s+", " ", head[m.end():end]).strip()
        value = value.replace("\\/", "/")        # IRIS escapes slashes
        if key == "sales_tax":
            # "Registered for Sales" + "No" + "Tax" -> "No"
            value = re.sub(r"\s*\bTax\b\s*$", "", value).strip()
        if value and not out.get(key):
            out[key] = value
    return out


def _key_for_label(label: str) -> str:
    """Map a label as printed back to its field key ('' if unknown)."""
    want = re.sub(r"[\s.:]+", " ", label).strip().lower()
    for known, key in _LABEL_TO_KEY.items():
        if re.sub(r"[\s.:]+", " ", known).strip().lower() == want:
            return key
    if want.startswith("registered for sales"):
        return "sales_tax"
    return ""


def parse_pdf(pdf_path: pathlib.Path) -> dict:
    """Read a saved profile PDF and extract the result fields."""
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            text = "\n".join((p.extract_text() or "") for p in pdf.pages)
        return parse_result_text(text)
    except Exception:
        return _empty_profile()


# Reads the profile straight out of the DOM: label/value pairs from the
# card's rows, plus the business / branch table.
_EXTRACT_JS = r"""
(labels) => {
    const norm = s => (s || '').replace(/\s+/g, ' ').trim();
    const key  = s => norm(s).replace(/[:\s.]+$/, '').toLowerCase();
    const want = new Set(labels.map(l => key(l)));

    const fields = {};
    const put = (k, v) => {
        if (!k || !v || fields[k]) return;
        if (want.has(key(v))) return;          // the next label, not a value
        if (v.length > 400) return;            // a whole block, not one value
        fields[k] = v;
    };

    // 1. Two-cell rows: "Label | Value".
    for (const tr of document.querySelectorAll('tr')) {
        const cells = tr.children;
        if (cells.length < 2) continue;
        const k = key(cells[0].innerText);
        if (want.has(k)) put(k, norm(cells[1].innerText));
    }

    // 2. Anything else: an element whose whole text IS a label. The value is
    //    the next element that carries text — as a sibling of the label, of
    //    its parent, or as the rest of the parent's own text.
    for (const el of document.querySelectorAll(
            'td,th,div,span,label,p,b,strong,dt,dd,li')) {
        const k = key(el.innerText);
        if (!want.has(k) || fields[k]) continue;

        let v = '';
        for (let sib = el.nextElementSibling; sib && !v; sib = sib.nextElementSibling)
            v = norm(sib.innerText);

        const parent = el.parentElement;
        if (!v && parent) {
            const ptxt = norm(parent.innerText), lbl = norm(el.innerText);
            if (ptxt.toLowerCase().indexOf(lbl.toLowerCase()) === 0)
                v = ptxt.slice(lbl.length).replace(/^[:\s]+/, '').trim();
            for (let sib = parent.nextElementSibling; sib && !v;
                 sib = sib.nextElementSibling)
                v = norm(sib.innerText);
        }
        put(k, v);
    }

    // Business / Branch table: the one whose header mentions Principal Activity
    let branches = [];
    for (const table of document.querySelectorAll('table')) {
        const head = norm(table.innerText).toLowerCase();
        if (head.indexOf('principal activity') === -1) continue;
        const rows = table.querySelectorAll('tr');
        const out = [];
        for (const tr of rows) {
            const cells = Array.from(tr.children).map(c => norm(c.innerText));
            if (cells.length < 3) continue;
            const first = cells[0].toLowerCase();
            if (first === 'sr' || first === 'sr.' ||
                cells.join(' ').toLowerCase().indexOf('principal activity') !== -1) continue;
            if (!cells.slice(1).join('')) continue;
            out.push({sr: cells[0], name: cells[1] || '',
                      address: cells[2] || '', activity: cells[3] || ''});
        }
        if (out.length) { branches = out; break; }
    }
    return JSON.stringify({fields: fields, branches: branches});
}
"""

# Returns the HTML of the smallest element that holds the whole result, so
# the surrounding portal (login box, menus) is left out of the PDF.
_RESULT_CARD_JS = r"""
(needles) => {
    const all = document.getElementsByTagName('*');
    let best = null;
    for (const el of all) {
        const txt = (el.innerText || '').toLowerCase();
        if (needles.every(n => txt.indexOf(n) !== -1)) {
            if (best === null || el.innerText.length < best.innerText.length) best = el;
        }
    }
    if (!best) return 'NOT_FOUND';
    // Walk out far enough to keep the FBR header strip above the fields.
    let card = best;
    for (let up = 0; up < 3; up++) {
        if (card.parentElement &&
            (card.parentElement.innerText || '').length < (card.innerText || '').length * 1.6) {
            card = card.parentElement;
        } else break;
    }
    return card.outerHTML;
}
"""

_CARD_DOC_HEAD = """<!doctype html><html><head><meta charset="utf-8">
<base href="https://iris.fbr.gov.pk/">
<style>
  @page { size: A4; margin: 14mm; }
  html,body { margin:0; padding:0; background:#fff; color:#000;
              font-family: Arial, Helvetica, sans-serif; font-size: 12pt; }
  img { max-width: 100%; }
  table { border-collapse: collapse; }
  .no-print, button, input, select { display: none !important; }
</style></head><body>"""

_CAPTCHA_ERROR_JS = r"""
() => {
    const sels = '.alert-danger,.text-danger,.toast-error,.toast-message,' +
                 '.swal2-popup,.mat-mdc-snack-bar-container,.invalid-feedback';
    for (const e of document.querySelectorAll(sels)) {
        const t = (e.innerText || '').trim();
        if (t && e.offsetParent !== null && /captcha|invalid|incorrect/i.test(t))
            return t.slice(0, 100);
    }
    return '';
}
"""


# ── Worker ───────────────────────────────────────────────────────────────────

class ProfileWorker:
    """
    Playwright-based automation that downloads Taxpayer Profiles from IRIS.

    Messages put on `out_q` (all tuples):
        ("log", level, text)      level = info | ok | err | warn
        ("record", Record)        a record changed
        ("report", path)          the Excel report was (re)written
        ("done", None)            the run is over
    """

    def __init__(
        self,
        records: List,
        run_date: datetime.date,
        output_dir: pathlib.Path,
        out_q: queue.Queue,
        stop_event: threading.Event,
        pause_event: threading.Event,
        show_browser: bool = False,
        overwrite: bool = False,
    ):
        self.records      = records
        self.output_dir   = output_dir
        self.out_q        = out_q
        self.stop_event   = stop_event
        self.pause_event  = pause_event      # SET = paused
        self.show_browser = show_browser
        self.overwrite    = overwrite

        self.date_file = run_date.strftime("%d-%m-%Y")   # for filenames

        self._form_is_fresh = False
        self._report_warned = False

        # Playwright objects — created in run()
        self._pw: Optional[object] = None
        self._browser: Optional[Browser] = None
        self._context: Optional[BrowserContext] = None
        self._page: Optional[Page] = None

    # ── Logging / progress helpers ────────────────────────────────────────

    def _log(self, msg: str, level: str = "info"):
        self.out_q.put(("log", level, f"[{_now()}] {msg}"))

    def _emit(self, rec):
        self.out_q.put(("record", rec))

    def _save_report(self):
        """(Re)write the summary file. Called after every record."""
        try:
            path = write_report(self.records, self.output_dir, self.date_file)
            self.out_q.put(("report", path))
            self._report_warned = False
        except Exception as exc:
            if not self._report_warned:          # e.g. the file is open in Excel
                self._report_warned = True
                self._log("Could not update the summary file (is it open in "
                          f"Excel? close it and it will update): {exc}", "warn")

    def _target(self, rec) -> pathlib.Path:
        return self.output_dir / f"{rec.reg_no}_{self.date_file}.pdf"

    # ── Public entry point ────────────────────────────────────────────────

    def run(self):
        """Call this in a daemon thread."""
        try:
            self._run_inner()
        except Exception as exc:
            self._log(f"Fatal error: {exc}", "err")
        finally:
            self._cleanup()
            for rec in self.records:            # run stopped mid-record
                if rec.result == RES_PROCESSING:
                    rec.result = RES_PENDING
                    self._emit(rec)
            self._save_report()
            self.out_q.put(("done", None))

    # ── Core loop ─────────────────────────────────────────────────────────

    def _run_inner(self):
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # 1. Anything already downloaded today is not fetched again
        pending = []
        for rec in self.records:
            if rec.result == RES_INVALID:
                continue
            target = self._target(rec)
            if (not self.overwrite and target.exists()
                    and target.stat().st_size > 800):
                self._apply(rec, parse_pdf(target))
                rec.pdf_file = target.name
                rec.result = RES_EXISTING
                rec.timestamp = rec.timestamp or _now_iso()
            else:
                rec.result = RES_PENDING
                pending.append(rec)
            self._emit(rec)

        self._save_report()                 # summary file exists from the start

        skipped = len(self.records) - len(pending)
        if skipped:
            self._log(f"{skipped} record(s) skipped (already downloaded or invalid).")
        if not pending:
            self._log("Nothing to download — all records are already done.", "ok")
            return

        # 2. Launch browser + open the Taxpayer Profile Inquiry form
        self._launch()
        self._open_form()

        # 3. Process each record
        total = len(pending)
        for n, rec in enumerate(pending, start=1):
            while self.pause_event.is_set() and not self.stop_event.is_set():
                time.sleep(0.5)
            if self.stop_event.is_set():
                self._log("Run cancelled by user.", "warn")
                break

            rec.result = RES_PROCESSING
            self._emit(rec)
            self._log(f"[{n}/{total}] {rec.param_type} {rec.reg_no} ...")

            self._process_with_retry(rec)
            rec.timestamp = _now_iso()
            self._emit(rec)

            if rec.result == RES_DOWNLOADED:
                self._log(
                    f"[{n}/{total}] {rec.reg_no} — {rec.name or 'Unknown'} — "
                    f"{rec.reg_status or 'no status'}", "ok")
            else:
                self._log(f"[{n}/{total}] {rec.reg_no} — {rec.result}", "err")

            self._save_report()
            if n < total:
                time.sleep(INTER_RECORD_DELAY_SECS)

        self._log("Finished.", "ok")

    # ── Browser ───────────────────────────────────────────────────────────

    def _launch(self):
        self._log("Launching Chromium browser (background)...")
        args = [
            "--window-size=1366,900",
            # keep the page fully alive while the window is off-screen
            "--disable-backgrounding-occluded-windows",
            "--disable-renderer-backgrounding",
            "--disable-background-timer-throttling",
            "--disable-features=CalculateNativeWinOcclusion",
        ]
        if not self.show_browser:
            args.append("--window-position=-32000,-32000")
        self._pw = sync_playwright().start()
        # Not headless: IRIS blocks headless browsers outright.
        self._browser = self._pw.chromium.launch(headless=False, args=args)
        self._context = self._browser.new_context(
            viewport={"width": 1366, "height": 900},
        )
        self._page = self._context.new_page()
        self._page.set_default_timeout(ELEMENT_TIMEOUT_MS)

    def _open_form(self):
        """Load IRIS and open the Taxpayer Profile Inquiry form."""
        page = self._page
        self._log("Opening FBR IRIS...")
        page.goto(IRIS_URL, wait_until="domcontentloaded",
                  timeout=PAGE_LOAD_TIMEOUT_MS)

        if "blocked" in (page.title() or "").lower():
            raise RuntimeError(
                "FBR IRIS blocked this connection (Web Page Blocked). "
                "Try again later or from another internet connection.")

        link = page.locator(SEL_TPI_LINK).first
        link.wait_for(state="visible", timeout=PAGE_LOAD_TIMEOUT_MS)
        link.click()
        page.locator(SEL_PARAM_BOX).first.wait_for(
            state="visible", timeout=ELEMENT_TIMEOUT_MS)
        page.locator(SEL_REG_INPUT).first.wait_for(
            state="visible", timeout=ELEMENT_TIMEOUT_MS)
        time.sleep(1)
        self._form_is_fresh = True
        self._log("Taxpayer Profile Inquiry form ready.", "ok")

    def _result_on_screen(self) -> bool:
        body = self._page.locator("body").inner_text(timeout=5_000)
        return self._is_result(body)

    @staticmethod
    def _is_result(body: str) -> bool:
        return (RESULT_MARKER in body and RESULT_MARKER_2 in body) \
            or NO_RECORD_MARKER in body.lower()

    def _reset_form(self):
        """Clear the previous result by re-clicking the sidebar link."""
        if self._form_is_fresh:
            self._form_is_fresh = False
            return
        page = self._page
        try:
            page.locator(SEL_TPI_LINK).first.click(timeout=8_000)
            page.locator(SEL_REG_INPUT).first.wait_for(
                state="visible", timeout=8_000)
            time.sleep(1)
            if not self._result_on_screen():
                return
        except Exception:
            pass
        # Sidebar click didn't give a clean form → full reload
        self._open_form()
        self._form_is_fresh = False

    # ── Retry wrapper ─────────────────────────────────────────────────────

    def _process_with_retry(self, rec):
        last_error = ""
        for attempt in range(1, MAX_RETRIES + 1):
            if self.stop_event.is_set():
                rec.result = RES_PENDING
                return
            try:
                self._process(rec)
                return
            except Exception as exc:
                last_error = str(exc).splitlines()[0][:120]
                self._log(f"  Attempt {attempt}/{MAX_RETRIES} failed: {last_error}",
                          "warn")
                if attempt < MAX_RETRIES:
                    time.sleep(RETRY_DELAY_SECS)
                    try:
                        self._open_form()
                    except Exception as exc2:
                        self._log(f"  Page reload failed: {exc2}", "warn")
        rec.result = f"{RES_FAILED} - {last_error}"

    # ── Single record ─────────────────────────────────────────────────────

    def _process(self, rec):
        self._reset_form()
        self._fill_form(rec)

        body = self._captcha_and_verify(rec)
        if body is None:
            if not self.stop_event.is_set():
                rec.result = f"{RES_FAILED} - captcha not accepted"
            else:
                rec.result = RES_PENDING
            return

        parsed = parse_result_text(body)
        if not parsed["no_record"]:
            dom = self._read_profile_dom()
            for key, value in dom.items():                 # DOM wins
                if value:
                    parsed[key] = value

        target = self._target(rec)
        saved = self._save_pdf(target, parsed["no_record"])

        if saved:
            from_pdf = parse_pdf(target)        # fill any gaps from the PDF
            for key, value in from_pdf.items():
                if not parsed.get(key):
                    parsed[key] = value
        self._apply(rec, parsed)

        if saved:
            rec.pdf_file = target.name
            rec.result = RES_DOWNLOADED
        else:
            rec.result = f"{RES_FAILED} - PDF not saved"

    def _read_profile_dom(self) -> dict:
        """Label/value pairs + branch rows straight from the rendered card."""
        out = {}
        try:
            raw = self._page.evaluate(_EXTRACT_JS, _LABELS)
            data = json.loads(raw) if raw else {}
        except Exception as exc:
            self._log(f"  Could not read the card from the page: "
                      f"{str(exc).splitlines()[0][:100]}", "warn")
            return out

        fields = data.get("fields") or {}
        for label, key in _LABEL_TO_KEY.items():
            norm = re.sub(r"[:\s.]+$", "", re.sub(r"\s+", " ", label)).lower()
            value = fields.get(norm, "")
            if value:
                out[key] = value
        branches = data.get("branches") or []
        if branches:
            out["branches"] = branches
        return out

    @staticmethod
    def _apply(rec, parsed: dict):
        for key, _ in PROFILE_FIELDS:
            value = parsed.get(key, "")
            if value:
                setattr(rec, key, value)
        if parsed.get("branches"):
            rec.branches = parsed["branches"]

    # ── Form filling ──────────────────────────────────────────────────────

    def _selected_param(self) -> str:
        return self._page.evaluate(
            "() => { const r = document.querySelector("
            "'#parameterType input.select-box__input:checked');"
            " return r ? r.value : ''; }") or ""

    @staticmethod
    def _same_param(a: str, b: str) -> bool:
        clean = lambda s: re.sub(r"[^a-z0-9]", "", (s or "").lower())
        return clean(a) == clean(b)

    def _select_param(self, param_type: str):
        """Pick the parameter type in the custom (non-<select>) box."""
        page = self._page
        for _ in range(4):
            if self._same_param(self._selected_param(), param_type):
                return
            try:
                page.locator(SEL_PARAM_CURRENT).first.click(timeout=5_000)
                page.locator(SEL_PARAM_OPTION).get_by_text(
                    param_type, exact=True).first.click(timeout=5_000)
            except Exception:
                # Same thing through the DOM, for when the list won't open
                page.evaluate(
                    """(target) => {
                        const clean = s => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
                        const cur = document.querySelector('#parameterType .select-box__current');
                        if (cur) cur.click();
                        for (const l of document.querySelectorAll('#parameterType label.select-box__option')) {
                            if (clean(l.innerText || l.textContent) === clean(target)) { l.click(); return; }
                        }
                    }""", param_type)
            time.sleep(0.5)
        if not self._same_param(self._selected_param(), param_type):
            raise RuntimeError(f"Could not select Parameter Type '{param_type}'")

    def _fill_form(self, rec):
        page = self._page
        self._select_param(rec.param_type)

        reg = page.locator(SEL_REG_INPUT).first
        reg.fill("")
        reg.fill(rec.reg_no)
        time.sleep(0.3)

    def _form_still_filled(self, rec) -> bool:
        page = self._page
        try:
            return (self._same_param(self._selected_param(), rec.param_type)
                    and page.locator(SEL_REG_INPUT).first.input_value() == rec.reg_no)
        except Exception:
            return False

    # ── Captcha + Verify ──────────────────────────────────────────────────

    def _read_captcha(self) -> str:
        b64 = self._page.evaluate(
            """(sel) => {
                const c = document.querySelector(sel) || document.querySelector('canvas');
                return (c && c.width > 0 && c.height > 0)
                    ? c.toDataURL('image/png').split(',')[1] : null;
            }""", SEL_CAPTCHA_CANVAS)
        if not b64:
            raise RuntimeError("Captcha image not found")
        img = Image.open(io.BytesIO(base64.b64decode(b64))).convert("RGBA")

        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[3])
        img = bg.resize((bg.width * 5, bg.height * 5), Image.LANCZOS).convert("L")
        img = ImageEnhance.Contrast(img).enhance(4.0)
        img = ImageEnhance.Sharpness(img).enhance(3.0)
        img = img.point(lambda p: 255 if p > 140 else 0)
        img = img.filter(ImageFilter.MedianFilter(3))

        best = ""
        for psm in (8, 7, 6, 13):
            cfg = (f"--psm {psm} -c tessedit_char_whitelist="
                   "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789")
            clean = re.sub(r"[^A-Za-z0-9]", "",
                           pytesseract.image_to_string(img, config=cfg))
            if len(clean) > len(best):
                best = clean
        return best

    def _refresh_captcha(self):
        try:
            self._page.locator(SEL_CAPTCHA_REFRESH).first.click(timeout=5_000)
            time.sleep(1.2)
        except Exception:
            pass

    def _dismiss_popups(self):
        page = self._page
        try:
            btn = page.locator(".swal2-confirm, .swal2-close").first
            if btn.is_visible():
                btn.click(timeout=2_000)
        except Exception:
            pass

    def _wait_outcome(self) -> Optional[str]:
        """
        After Verify: return the page text once a profile (or "No record
        exists") shows up; None if the captcha was rejected / nothing came.
        """
        page = self._page
        deadline = time.time() + VERIFY_WAIT_SECS
        while time.time() < deadline:
            time.sleep(0.4)
            try:
                body = page.locator("body").inner_text(timeout=5_000)
            except Exception:
                continue
            if RESULT_MARKER in body and RESULT_MARKER_2 in body:
                time.sleep(0.6)                     # let every field render
                return page.locator("body").inner_text(timeout=5_000)
            if NO_RECORD_MARKER in body.lower():
                return body
            try:
                if page.evaluate(_CAPTCHA_ERROR_JS):
                    return None
            except Exception:
                pass
        return None

    def _captcha_and_verify(self, rec) -> Optional[str]:
        page = self._page
        for attempt in range(1, MAX_CAPTCHA_RETRIES + 1):
            if self.stop_event.is_set():
                return None
            if attempt > 1:
                self._dismiss_popups()
                self._refresh_captcha()
                if not self._form_still_filled(rec):
                    self._fill_form(rec)

            try:
                text = self._read_captcha()
            except Exception as exc:
                self._log(f"  Captcha read error: {exc}", "warn")
                continue
            if len(text) < 4:
                continue

            cap = page.locator(SEL_CAPTCHA_INPUT).first
            cap.fill("")
            cap.fill(text)
            page.locator(SEL_VERIFY_BUTTON).first.click()

            body = self._wait_outcome()
            if body is not None:
                return body
            self._log(f"  Captcha attempt {attempt}/{MAX_CAPTCHA_RETRIES} not accepted",
                      "warn")
        return None

    # ── PDF ───────────────────────────────────────────────────────────────

    def _save_pdf(self, target: pathlib.Path, no_record: bool) -> bool:
        """Render only the profile card to a one-page PDF."""
        page = self._page
        if no_record:
            needles = [NO_RECORD_MARKER]
        else:
            needles = [RESULT_MARKER.lower(), RESULT_MARKER_2.lower()]
            # Keep the business / branch table in the PDF when there is one:
            # without this the smallest matching element stops above it.
            try:
                body = page.locator("body").inner_text(timeout=5_000).lower()
                if "principal activity" in body:
                    needles.append("principal activity")
            except Exception:
                pass
        tmp = target.with_suffix(".pdf.part")
        pdf_page = None
        try:
            card_html = page.evaluate(_RESULT_CARD_JS, needles)
            if not card_html or card_html == "NOT_FOUND":
                self._log("  Profile card not found — saving the whole page.", "warn")
                page.pdf(path=str(tmp), format="A4", print_background=True)
            else:
                pdf_page = self._context.new_page()
                try:
                    pdf_page.set_content(_CARD_DOC_HEAD + card_html + "</body></html>",
                                         wait_until="load", timeout=20_000)
                except Exception:
                    pass                            # a slow image is not fatal
                pdf_page.wait_for_timeout(400)
                pdf_page.pdf(path=str(tmp), format="A4", print_background=True)

            if not tmp.exists() or tmp.stat().st_size < 800:
                return False
            os.replace(tmp, target)
            return True
        except Exception as exc:
            self._log(f"  PDF error: {str(exc).splitlines()[0][:120]}", "warn")
            return False
        finally:
            try:
                if pdf_page:
                    pdf_page.close()
            except Exception:
                pass
            try:
                if tmp.exists():
                    tmp.unlink()
            except Exception:
                pass

    # ── Cleanup ───────────────────────────────────────────────────────────

    def _cleanup(self):
        """Shut down Playwright gracefully."""
        for obj in (self._context, self._browser):
            try:
                if obj:
                    obj.close()
            except Exception:
                pass
        try:
            if self._pw:
                self._pw.stop()
        except Exception:
            pass
