"""
Taxpayer Profile Inquiry — Entry Point
=========================
Run this file:  python main.py     (or just double-click run.bat)

On start-up the app checks for everything it needs and downloads whatever
is missing by itself (Python packages, Playwright's Chromium browser and
Tesseract-OCR). Nothing needs to be installed manually except Python.

Options:
    python main.py --setup-only    download/verify dependencies, then exit
"""

import sys


def _enable_hidpi():
    """
    Tell Windows this app draws at the screen's real resolution. Without it
    Windows stretches the window on 125% / 150% displays and all text looks
    blurry. Must run before the first window (incl. the setup splash) opens.
    """
    if sys.platform != "win32":
        return
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)      # per-monitor
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def main():
    _enable_hidpi()

    # bootstrap uses only the standard library, so it can run before any
    # third-party package has been installed.
    from bootstrap import ensure_ready

    setup_only = "--setup-only" in sys.argv
    ensure_ready(show_gui=not setup_only)

    if setup_only:
        print("Setup complete. All dependencies are installed.")
        return

    from gui import TPICheckerApp
    app = TPICheckerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
