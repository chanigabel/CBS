"""Launcher entry point for the Excel Normalization desktop application.

This script is the PyInstaller entry point.  It:
1. Starts the Uvicorn server on a free local port.
2. Opens the app in Google Chrome if available, otherwise the default browser.
3. Keeps running until the user closes the console window or sends Ctrl+C.

The server binds to 127.0.0.1 only — no network exposure.
"""

import logging
import os
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from pathlib import Path

import uvicorn

# ---------------------------------------------------------------------------
# Logging — write to both console and a log file in the data directory
# ---------------------------------------------------------------------------

def _setup_logging() -> None:
    if getattr(sys, "frozen", False):
        log_dir = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "ExcelNormalization"
    else:
        log_dir = Path.cwd()

    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "app.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(str(log_file), encoding="utf-8"),
        ],
    )


# ---------------------------------------------------------------------------
# Port selection
# ---------------------------------------------------------------------------

def _find_free_port(preferred: int = 8765) -> int:
    """Return *preferred* if it is free, otherwise find any free port."""
    for port in range(preferred, preferred + 100):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


# ---------------------------------------------------------------------------
# Chrome detection
# ---------------------------------------------------------------------------

# Standard Chrome installation paths on Windows (64-bit and 32-bit)
_CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]


def _find_chrome() -> str | None:
    """Return the path to chrome.exe if Chrome is installed, else None.

    Checks:
    1. Well-known installation paths.
    2. LOCALAPPDATA (per-user Chrome install).
    3. The registry (HKCU and HKLM) via the 'where' command as a fallback.
    """
    # 1. Standard paths
    for path in _CHROME_PATHS:
        if Path(path).is_file():
            return path

    # 2. Per-user install under %LOCALAPPDATA%
    local_app = os.environ.get("LOCALAPPDATA", "")
    if local_app:
        candidate = Path(local_app) / "Google" / "Chrome" / "Application" / "chrome.exe"
        if candidate.is_file():
            return str(candidate)

    # 3. Registry lookup via 'where' (best-effort, silent on failure)
    try:
        result = subprocess.run(
            ["where", "chrome"],
            capture_output=True,
            text=True,
            timeout=3,
        )
        if result.returncode == 0:
            first_line = result.stdout.strip().splitlines()[0].strip()
            if first_line and Path(first_line).is_file():
                return first_line
    except Exception:
        pass

    return None


# ---------------------------------------------------------------------------
# Browser opener (runs in a background thread so it fires after server starts)
# ---------------------------------------------------------------------------

def _open_browser(url: str, delay: float = 1.5) -> None:
    """Open *url* in Chrome if available, otherwise fall back to the default browser."""
    time.sleep(delay)
    logger = logging.getLogger(__name__)

    chrome = _find_chrome()
    if chrome:
        try:
            # --new-window: always open a fresh window
            # --app=URL:    open in app mode (no address bar, cleaner look)
            subprocess.Popen(
                [chrome, f"--app={url}", "--new-window"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            logger.info(f"Opened Chrome: {chrome}")
            return
        except Exception as exc:
            logger.warning(f"Chrome launch failed ({exc}), falling back to default browser.")

    # Fallback: system default browser
    try:
        webbrowser.open(url)
        logger.info("Opened default browser.")
    except Exception as exc:
        logger.warning(f"Could not open browser: {exc}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    _setup_logging()
    logger = logging.getLogger(__name__)

    port = _find_free_port(8765)
    url = f"http://127.0.0.1:{port}"

    chrome = _find_chrome()
    browser_label = f"Chrome ({chrome})" if chrome else "default browser"

    logger.info(f"Starting Excel Normalization at {url}")
    print(f"\n  Excel Normalization is running at {url}")
    print(f"  Opening in {browser_label}...")
    print("  Press Ctrl+C to stop.\n")

    # Open browser in background after a short delay
    t = threading.Thread(target=_open_browser, args=(url,), daemon=True)
    t.start()

    # Start Uvicorn — import app here so _MEIPASS is already set
    from webapp.app import app as fastapi_app

    uvicorn.run(
        fastapi_app,
        host="127.0.0.1",
        port=port,
        log_level="info",
        reload=False,
    )


if __name__ == "__main__":
    main()
