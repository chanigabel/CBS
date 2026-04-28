# Building the Windows Installer

This document covers how to produce the offline Windows installer for
**Excel standardization** from the source repository.

---

## Prerequisites (build machine — needs internet once)

| Tool | Version | Install |
|---|---|---|
| Python | 3.11 or 3.12 (64-bit) | python.org |
| PyInstaller | ≥ 6.0 | `pip install pyinstaller` |
| Inno Setup | 6.x | https://jrsoftware.org/isinfo.php |

Install the project and its dependencies:

```bat
pip install -r requirements.txt
pip install pyinstaller
```

---

## Build the EXE bundle

```bat
build_exe.bat
```

Output: `dist\Excelstandardization\` — a self-contained folder with
`Excelstandardization.exe` and all required DLLs/assets.

The script automatically:
- Cleans previous `build\` and `dist\` artefacts
- Removes stale `__pycache__` bytecode
- Runs PyInstaller with `Excelstandardization.spec`
- Verifies that the exe and key assets are present in the output

You can test the bundle immediately without installing:

```bat
dist\Excelstandardization\Excelstandardization.exe
```

The app starts a local server and opens automatically in **Google Chrome**
if Chrome is installed, or in the default browser otherwise.

---

## Build the installer

```bat
build_installer.bat
```

This runs `build_exe.bat` first, then compiles the Inno Setup script.
`iscc.exe` is located automatically (PATH → Program Files (x86) → Program Files).

Output: `installer\Output\Excelstandardization_Setup_1.0.0.exe`

---

## What the installer does

- Installs to `%ProgramFiles%\Excel standardization\`
- Pre-creates writable runtime directories under `%LOCALAPPDATA%\Excelstandardization\`
- Creates a Start Menu shortcut (optional desktop shortcut)
- Registers a standard Windows uninstaller
- Offers to launch the app immediately after install

---

## Browser behaviour

When launched, the app tries to open in **Google Chrome** using `--app` mode
(no address bar, cleaner look).  Chrome is located by checking:

1. `C:\Program Files\Google\Chrome\Application\chrome.exe`
2. `C:\Program Files (x86)\Google\Chrome\Application\chrome.exe`
3. `%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe` (per-user install)
4. `where chrome` (PATH lookup)

If Chrome is not found, the system default browser is used instead.

---

## Runtime data location

All user data (uploads, working copies, exports, log file) is stored in:

```
%LOCALAPPDATA%\Excelstandardization\
  uploads\
  work\
  output\
  app.log
```

This folder is created by the installer and on first run.
It is **not** removed by the uninstaller (user data is preserved).

---

## Testing on a clean offline Windows machine

1. Copy `installer\Output\Excelstandardization_Setup_1.0.0.exe` to the target machine.
2. Run the installer — no internet required, no Python required.
3. Launch from the Start Menu or desktop shortcut.
4. Verify Chrome (or the default browser) opens at `http://127.0.0.1:8765`.
5. Upload a `.xlsx` or `.xlsm` file and confirm the sheet grid appears.
6. Click **Run standardization** and confirm corrected columns appear.
7. Click **Export / Download** and confirm a `.xlsx` file downloads.
8. Check `%LOCALAPPDATA%\Excelstandardization\app.log` for any errors.
9. Uninstall via **Add or Remove Programs** and confirm the entry is gone.
10. Confirm `%LOCALAPPDATA%\Excelstandardization\` still exists (user data preserved).

---

## Troubleshooting

**Browser does not open automatically**
Open `http://127.0.0.1:8765` manually in any browser.

**Port 8765 is in use**
The launcher automatically tries the next 100 ports. Check `app.log` for
the actual URL used.

**"Failed to execute script" on launch**
Run `Excelstandardization.exe` from a Command Prompt to see the full
traceback, then check `app.log`.

**PyInstaller misses a hidden import**
Add the missing module to `hidden_imports` in `Excelstandardization.spec`
and rebuild with `build_exe.bat`.

**Inno Setup not found**
Install Inno Setup 6 from https://jrsoftware.org/isinfo.php.
`build_installer.bat` checks PATH and both standard install locations automatically.
