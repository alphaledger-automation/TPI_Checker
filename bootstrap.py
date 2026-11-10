"""
Taxpayer Profile Inquiry — Self-Setup (Bootstrap)
====================================
Makes the app download everything it needs by itself:

  1. Missing Python packages (customtkinter, playwright, openpyxl, pillow,
     pytesseract, pdfplumber) are installed with pip.
  2. Playwright's Chromium browser is downloaded if it isn't present.
  3. Tesseract-OCR (used to read the IRIS captcha) is installed with
     winget if it isn't present.

If anything needs downloading, a small progress window is shown (plain
tkinter, so it works even before customtkinter is installed).  If
everything is already in place, this module returns almost instantly and
shows nothing.

Every step is written to  <user-data-dir>/setup.log  for troubleshooting.
"""

from __future__ import annotations

import importlib
import importlib.util
import os
import pathlib
import queue
import re
import shutil
import site
import subprocess
import sys
import threading
import traceback
from typing import Callable, List, Optional

APP_TITLE    = "Taxpayer Profile Inquiry"
APP_DIR_NAME = "TPI_Checker"

# import-name -> pip-name.  Used to decide what is missing.
REQUIRED_MODULES = {
    "customtkinter": "customtkinter",
    "playwright": "playwright",
    "openpyxl": "openpyxl",
    "PIL": "pillow",
    "pytesseract": "pytesseract",
    "pdfplumber": "pdfplumber",
}

TESSERACT_WINGET_ID = "UB-Mannheim.TesseractOCR"
TESSERACT_MANUAL_URL = "https://github.com/UB-Mannheim/tesseract/wiki"

_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)   # 0 on non-Windows

LogFn = Callable[[str], None]


# ── Paths ─────────────────────────────────────────────────────────────────────

def _data_dir() -> pathlib.Path:
    """Per-user writable folder for the setup log."""
    if sys.platform == "win32":
        base = pathlib.Path(os.environ.get("LOCALAPPDATA", pathlib.Path.home()))
    else:
        base = pathlib.Path.home() / ".local" / "share"
    d = base / APP_DIR_NAME
    d.mkdir(parents=True, exist_ok=True)
    return d


def _app_dir() -> pathlib.Path:
    """Folder containing requirements.txt (next to main.py / this file)."""
    return pathlib.Path(__file__).resolve().parent


def log_file() -> pathlib.Path:
    return _data_dir() / "setup.log"


# ── Detection ─────────────────────────────────────────────────────────────────

def missing_python_packages() -> List[str]:
    """Return pip names of required packages that can't be imported."""
    return [
        pip_name for mod, pip_name in REQUIRED_MODULES.items()
        if importlib.util.find_spec(mod) is None
    ]


def chromium_installed() -> bool:
    """True if Playwright's Chromium binary exists on disk."""
    if importlib.util.find_spec("playwright") is None:
        return False
    try:
        from playwright.sync_api import sync_playwright
        pw = sync_playwright().start()
        try:
            return os.path.exists(pw.chromium.executable_path)
        finally:
            pw.stop()
    except Exception:
        return False


def find_tesseract() -> Optional[str]:
    """Return the path of tesseract.exe, or None if it isn't installed."""
    candidates = [_app_dir() / "tesseract" / "tesseract.exe"]
    for var in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
        root = os.environ.get(var)
        if root:
            candidates.append(pathlib.Path(root) / "Tesseract-OCR" / "tesseract.exe")
    local = os.environ.get("LOCALAPPDATA")
    if local:
        candidates.append(pathlib.Path(local) / "Programs" / "Tesseract-OCR" / "tesseract.exe")
        candidates.append(pathlib.Path(local) / "Tesseract-OCR" / "tesseract.exe")
    for c in candidates:
        if c.exists():
            return str(c)
    return shutil.which("tesseract")


# ── Installation steps ────────────────────────────────────────────────────────

def _run_streaming(cmd: List[str], log: LogFn, env: Optional[dict] = None) -> int:
    """Run a command, forwarding each output line to `log`. Returns exit code."""
    log("$ " + " ".join(cmd))
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        stdin=subprocess.DEVNULL,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
        creationflags=_NO_WINDOW,
    )
    assert proc.stdout is not None
    for line in proc.stdout:            # text mode: '\r' progress ticks split too
        line = line.strip()
        if line:
            log(line)
    return proc.wait()


def _add_user_site_to_path() -> None:
    """After a `pip --user` install the user site may not be on sys.path yet."""
    try:
        us = site.getusersitepackages()
        if us and os.path.isdir(us) and us not in sys.path:
            sys.path.append(us)
    except Exception:
        pass
    importlib.invalidate_caches()


def install_python_packages(log: LogFn) -> None:
    """pip-install requirements.txt (falls back to --user / ensurepip)."""
    req = _app_dir() / "requirements.txt"
    if not req.exists():
        # requirements.txt missing → install the known package list directly
        targets = list(REQUIRED_MODULES.values())
    else:
        targets = ["-r", str(req)]

    base = [sys.executable, "-m", "pip", "install",
            "--disable-pip-version-check"] + targets

    # Make sure pip itself exists
    if subprocess.call([sys.executable, "-m", "pip", "--version"],
                       stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL, creationflags=_NO_WINDOW) != 0:
        log("pip not found — bootstrapping it with ensurepip...")
        _run_streaming([sys.executable, "-m", "ensurepip", "--upgrade"], log)

    rc = _run_streaming(base, log)
    if rc != 0:
        log("Normal install failed — retrying with --user ...")
        rc = _run_streaming(base + ["--user"], log)
    if rc != 0:
        raise RuntimeError(
            "Could not install the required Python packages.\n"
            "Check your internet connection and try again."
        )
    _add_user_site_to_path()


def install_chromium(log: LogFn, on_percent: Optional[Callable[[int], None]] = None) -> None:
    """Download Playwright's Chromium using Playwright's own bundled driver."""
    from playwright._impl._driver import compute_driver_executable, get_driver_env

    driver = compute_driver_executable()
    driver_cmd = list(driver) if isinstance(driver, (tuple, list)) else [str(driver)]
    env = get_driver_env()
    env.pop("PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD", None)

    pct_re = re.compile(r"(\d{1,3})\s*%")

    def _log(line: str) -> None:
        log(line)
        m = pct_re.search(line)
        if m and on_percent:
            on_percent(min(int(m.group(1)), 100))

    rc = _run_streaming(driver_cmd + ["install", "chromium"], _log, env=env)
    if rc != 0:
        raise RuntimeError(
            "Could not download the Chromium browser.\n"
            "Check your internet connection (or firewall/proxy) and try again."
        )


def install_tesseract(log: LogFn) -> None:
    """Install Tesseract-OCR with winget (Windows will ask for permission)."""
    manual = (
        "Tesseract-OCR could not be installed automatically.\n"
        f"Please install it from:\n{TESSERACT_MANUAL_URL}\n"
        "(keep the default install folder), then start the app again."
    )
    if sys.platform != "win32" or shutil.which("winget") is None:
        raise RuntimeError(manual)

    _run_streaming(
        ["winget", "install", "-e", "--id", TESSERACT_WINGET_ID, "--silent",
         "--accept-package-agreements", "--accept-source-agreements",
         "--disable-interactivity"],
        log,
    )
    # winget's exit code is non-zero for "already installed" too, so trust
    # only what is actually on disk.
    if not find_tesseract():
        raise RuntimeError(manual)


# ── Orchestration ─────────────────────────────────────────────────────────────

def _setup_steps(log: LogFn, status: LogFn,
                 on_percent: Optional[Callable[[int], None]] = None) -> None:
    """Do whatever is needed. Runs in a worker thread when a GUI is shown."""
    missing = missing_python_packages()
    if missing:
        status("Installing required Python packages...")
        log("Missing packages: " + ", ".join(missing))
        install_python_packages(log)
        still = missing_python_packages()
        if still:
            raise RuntimeError("Packages still missing after install: " + ", ".join(still))
        log("Python packages installed.")

    if not chromium_installed():
        status("Downloading Chromium browser (one-time, ~150 MB)...")
        install_chromium(log, on_percent)
        if not chromium_installed():
            raise RuntimeError(
                "Chromium was downloaded but could not be found.\n"
                f"See the log for details:\n{log_file()}"
            )
        log("Chromium installed.")

    if not find_tesseract():
        status("Installing Tesseract-OCR (click Yes if Windows asks)...")
        install_tesseract(log)
        log("Tesseract-OCR installed.")


def _file_logger() -> LogFn:
    path = log_file()

    def _log(msg: str) -> None:
        try:
            with open(path, "a", encoding="utf-8") as fh:
                fh.write(msg + "\n")
        except Exception:
            pass
    return _log


def _needs_setup() -> bool:
    return (bool(missing_python_packages())
            or not chromium_installed()
            or not find_tesseract())


def ensure_ready(show_gui: bool = True) -> None:
    """
    Make sure every dependency is present, downloading what's missing.
    Exits the process with a message if setup cannot be completed.
    """
    flog = _file_logger()

    if not _needs_setup():
        return

    if show_gui:
        ok, err = _run_with_splash(flog)
    else:
        ok, err = True, ""
        try:
            _setup_steps(lambda m: (print(m), flog(m)),
                         lambda m: print(f"== {m}"))
        except Exception as exc:                       # noqa: BLE001
            flog(traceback.format_exc())
            ok, err = False, str(exc)

    if not ok:
        _fatal(f"{err}\n\nSetup log: {log_file()}")


def _fatal(message: str) -> None:
    """Show an error (GUI if possible, else console) and exit."""
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(f"{APP_TITLE} — Setup failed", message)
        root.destroy()
    except Exception:
        print("SETUP FAILED:\n" + message)
    sys.exit(1)


# ── Splash / progress window (plain tkinter) ──────────────────────────────────

def _run_with_splash(flog: LogFn):
    """Show a progress window while setup runs. Returns (ok, error_text)."""
    try:
        import tkinter as tk
        from tkinter import ttk
    except Exception:
        # No tkinter → fall back to console mode
        try:
            _setup_steps(lambda m: (print(m), flog(m)), lambda m: print(f"== {m}"))
            return True, ""
        except Exception as exc:                        # noqa: BLE001
            flog(traceback.format_exc())
            return False, str(exc)

    q: "queue.Queue[tuple]" = queue.Queue()
    result = {"ok": True, "err": ""}

    def log(msg: str) -> None:
        flog(msg)
        q.put(("log", msg))

    def status(msg: str) -> None:
        flog("== " + msg)
        q.put(("status", msg))

    def percent(p: int) -> None:
        q.put(("pct", p))

    def worker() -> None:
        try:
            _setup_steps(log, status, percent)
        except Exception as exc:                        # noqa: BLE001
            flog(traceback.format_exc())
            result["ok"], result["err"] = False, str(exc)
        finally:
            q.put(("done", None))

    root = tk.Tk()
    root.title(f"{APP_TITLE} — First-time setup")
    scale = max(1.0, root.winfo_fpixels("1i") / 96.0)    # 1.25 on a 125% display
    root.geometry(f"{int(560 * scale)}x{int(300 * scale)}")
    root.resizable(False, False)

    lbl_title = tk.Label(root, text="Preparing the application…",
                         font=("Segoe UI", 12, "bold"))
    lbl_title.pack(pady=(14, 2))
    lbl_status = tk.Label(root, text="Checking required components…",
                          font=("Segoe UI", 10))
    lbl_status.pack()
    bar = ttk.Progressbar(root, mode="indeterminate", length=int(500 * scale))
    bar.pack(pady=10)
    bar.start(12)
    tk.Label(root, text="This only happens once and needs an internet connection.",
             font=("Segoe UI", 9), fg="#666").pack()

    txt = tk.Text(root, height=8, width=68, font=("Consolas", 8),
                  state="disabled", bg="#f4f4f4")
    txt.pack(padx=12, pady=8)

    def pump() -> None:
        try:
            while True:
                kind, val = q.get_nowait()
                if kind == "log":
                    txt.configure(state="normal")
                    txt.insert("end", val + "\n")
                    txt.see("end")
                    txt.configure(state="disabled")
                elif kind == "status":
                    lbl_status.configure(text=val)
                    if str(bar.cget("mode")) != "indeterminate":
                        bar.configure(mode="indeterminate")
                        bar.start(12)
                elif kind == "pct":
                    if str(bar.cget("mode")) != "determinate":
                        bar.stop()
                        bar.configure(mode="determinate", maximum=100)
                    bar["value"] = val
                elif kind == "done":
                    root.destroy()
                    return
        except queue.Empty:
            pass
        root.after(100, pump)

    root.protocol("WM_DELETE_WINDOW", lambda: None)      # can't close mid-install
    threading.Thread(target=worker, daemon=True).start()
    root.after(100, pump)
    root.mainloop()
    return result["ok"], result["err"]
