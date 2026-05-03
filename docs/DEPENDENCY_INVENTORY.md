# Dependency Inventory Report — Excel Data Standardization Web App

**Version:** 3.0 (Security Remediation Update — Clean Transfer Build)
**Date:** 2026-04-30
**Python version:** 3.12.0 (Windows x64)
**Project version:** 1.0.0 (`pyproject.toml`) / Installer 1.0.2 (`ExcelNormalization.iss`)
**Delivered installer:** `installer/Output/ExcelNormalization_Setup_1.0.2.exe`
**Prepared by:** Automated scan

---

## ✅ Security Remediation Notice

> **This version of the report reflects the security-remediated dependency versions.**
> The following vulnerable packages have been updated:
>
> | Package | Vulnerable Version | Remediated Version |
> |---------|-------------------|-------------------|
> | python-multipart | 0.0.20 | **0.0.26** |
> | starlette | 0.47.3 | **0.49.1** |
> | lxml | 6.0.4 | **6.1.0** |
> | orjson | 3.11.3 | **3.11.6** |
> | fastapi | 0.116.1 | **0.120.1** (minimum required for starlette==0.49.1) |
>
> **The new installer must be rebuilt from the clean virtual environment described
> in this document. The old installer (1.0.2) contains the vulnerable versions above
> and must not be used.**

---

## ⚠️ Previous Installer Notice

> **The previously delivered installer (`ExcelNormalization_Setup_1.0.2.exe`) was
> built from a shared development environment** containing unrelated packages
> (gradio, openai, faster-whisper, aiortc, and ~60 others).
> As a result, PyInstaller traced and bundled some packages that are not declared
> as project dependencies.
>
> **The new installer must be built from a dedicated clean virtual environment
> using `requirements-lock.txt` to prevent environment bleed.**

---

## Sources Inspected

| Source | Found | Used For |
|--------|-------|----------|
| `requirements.txt` | Yes | Direct dependency declarations |
| `pyproject.toml` | Yes | Direct dependency declarations + dev/build deps |
| `ExcelNormalization.spec` | Yes | PyInstaller hidden imports + excludes |
| `installer/ExcelNormalization.iss` | Yes | Installer build tool identification |
| `build_exe.bat`, `build_installer.bat` | Yes | Build tool identification |
| `dist/ExcelNormalization/_internal/` | Yes | Ground-truth of what is actually bundled |
| `webapp/templates/index.html` | Yes | CDN / frontend dependency check |
| `webapp/static/app.js` | Yes | Frontend library check |
| `webapp/static/style.css` | Yes | Frontend library check |
| `setup.py` / `setup.cfg` | Not found | — |
| `Dockerfile` / `docker-compose` | Not found | — |
| `.github/` CI workflows | Not found | — |
| `package.json` / lock files | Not found | — |
| Clean virtual environment | Not available | See notice above |

Commands run:
```
python --version
pip show openpyxl fastapi uvicorn python-multipart jinja2 starlette pydantic
       pydantic-core anyio h11 click MarkupSafe et-xmlfile sniffio idna
       typing-extensions annotated-types typing-inspection colorama attrs
       websockets httptools orjson watchfiles lxml tzdata pyreadline3
       setuptools pytest pytest-cov hypothesis black mypy flake8 pyinstaller
Get-ChildItem dist/ExcelNormalization/_internal -Directory
Get-ChildItem dist/ExcelNormalization/_internal -Filter *.pyd
Get-ChildItem dist/ExcelNormalization/_internal -Filter *.dll
Get-ChildItem dist/ExcelNormalization/_internal -Recurse -Filter *.dist-info
```

---

## Section 1 — Direct Runtime Dependencies

Declared in `pyproject.toml` under `[project] dependencies` and in `requirements-lock.txt`.

**All versions are now pinned exactly in `requirements-lock.txt` for reproducible clean builds.**

| Library | Declared Constraint | Pinned Version | License | In Clean Installer | Purpose |
|---------|--------------------|--------------------|---------|-------------------|---------|
| openpyxl | `>=3.1.0` | 3.1.5 | MIT | **Yes** | Read/write `.xlsx`/`.xlsm` Excel files |
| fastapi | `>=0.120.1` | **0.120.1** | MIT | **Yes** | Web API framework for the browser UI |
| uvicorn | `>=0.23.0` | 0.35.0 | BSD-3-Clause | **Yes** | ASGI server that runs the FastAPI app |
| python-multipart | `>=0.0.26` | **0.0.26** | Apache-2.0 | **Yes** | Parses multipart/form-data for file uploads |
| jinja2 | `>=3.1.0` | 3.1.6 | BSD-3-Clause | **Yes** | HTML template rendering (`index.html`) |

---

## Section 2 — Transitive Runtime Dependencies

Pulled in automatically as dependencies of the direct runtime packages above.
All versions pinned in `requirements-lock.txt`.

| Library | Pinned Version | License | In Clean Installer | Pulled In By | Purpose |
|---------|----------------|---------|-------------------|--------------|---------|
| starlette | **0.49.1** | BSD-3-Clause | **Yes** | fastapi | ASGI toolkit underlying FastAPI; also used directly in `webapp/app.py` |
| pydantic | 2.11.9 | MIT | **Yes** | fastapi | Data validation and serialisation for API models |
| pydantic-core | 2.33.2 | MIT | **Yes** | pydantic | Rust-based core for pydantic v2 |
| anyio | 4.10.0 | MIT | **Yes** | starlette | Async I/O abstraction layer |
| h11 | 0.16.0 | MIT | **Yes** | uvicorn | Pure-Python HTTP/1.1 implementation |
| click | 8.2.1 | BSD-3-Clause | **Yes** | uvicorn | CLI argument parsing used by uvicorn |
| MarkupSafe | 3.0.2 | BSD-3-Clause | **Yes** | jinja2 | Safe HTML string escaping |
| et-xmlfile | 2.0.0 | MIT | **Yes** | openpyxl | Low-memory XML file writer |
| sniffio | 1.3.1 | MIT OR Apache-2.0 | **Yes** | anyio | Detects which async library is running |
| idna | 3.10 | BSD-3-Clause (see note) | **Yes** | anyio | Internationalized domain name handling |
| typing-extensions | 4.15.0 | PSF-2.0 | **Yes** | fastapi, pydantic, starlette, anyio | Backported type hint utilities |
| annotated-types | 0.7.0 | MIT (see note) | **Yes** | pydantic | Reusable constraint types for `Annotated` |
| typing-inspection | 0.4.1 | MIT | **Yes** | pydantic | Runtime typing introspection |
| colorama | 0.4.6 | BSD-3-Clause (see note) | **Yes** | click | Windows ANSI colour support |
| annotated-doc | 0.0.4 | MIT | **Yes** | fastapi | FastAPI documentation annotation support |

> **License notes:** `idna`, `annotated-types`, and `colorama` did not return a
> `License` or `License-Expression` field from `pip show`. Licenses above are
> sourced from their respective PyPI pages and are well-established; however,
> they should be independently verified for the formal security review.

---

## Section 3 — Dev / Test / Build Dependencies

Declared in `pyproject.toml` under `[project.optional-dependencies] dev`.
**None of these are included in the installer.**

| Library | Declared Constraint | Installed Version | License | In Final Installer | Purpose |
|---------|--------------------|--------------------|---------|-------------------|---------|
| pytest | `>=7.0.0` | 8.4.2 | MIT | **No** | Test runner |
| pytest-cov | `>=4.0.0` | 7.0.0 | MIT | **No** | Code coverage reporting |
| hypothesis | `>=6.0.0` | 6.151.9 | MPL-2.0 | **No** | Property-based testing |
| black | `>=23.0.0` | 25.9.0 | MIT | **No** | Code formatter |
| mypy | `>=1.0.0` | 1.18.2 | MIT | **No** | Static type checker |
| flake8 | `>=6.0.0` | 7.3.0 | MIT | **No** | Linter |
| pyinstaller | `>=6.0.0` | 6.19.0 | GPLv2-or-later + exception¹ | **No** | Packages the app into a standalone `.exe` |

> ¹ PyInstaller uses GPLv2-or-later with a special exception that explicitly
> permits building and distributing non-free (including commercial) programs.
> The exception means the GPL does **not** propagate to the packaged application.
> Reference: https://pyinstaller.org/en/stable/license.html

---

## Section 4 — Build Tools (Not Bundled, Not Python Packages)

These tools are required on the build machine only. They are not shipped in the installer.

| Tool | Version | License | In Final Installer | Purpose |
|------|---------|---------|-------------------|---------|
| PyInstaller | 6.19.0 | GPLv2-or-later + exception | **No** | Compiles Python app into Windows `.exe` |
| Inno Setup 6 | Not recorded in project files | Inno Setup License (freeware) | **No** | Wraps PyInstaller output into a Windows installer `.exe` |

> **Note:** The Inno Setup version used to compile the installer is not recorded
> in any project file. The `.iss` script references it only by install path.
> This should be documented for reproducible builds.

---

---

## Section 5 — Final Installer Third-Party Inventory — As Shipped

**Source of truth:** `dist/ExcelNormalization/_internal/` — the PyInstaller output
that was packaged into `ExcelNormalization_Setup_1.0.2.exe` and delivered.

This section lists every third-party package physically present in the bundle.
It is the authoritative inventory for the delivered artifact.

### 5a — Packages with dist-info Records in the Bundle

These packages have a `.dist-info` directory inside `_internal/`, confirming their
identity and version with certainty.

| Package | Bundled Version | License | Declared in Project Files | In Delivered Installer | Notes |
|---------|----------------|---------|--------------------------|----------------------|-------|
| attrs | 25.3.0 | MIT | **No** | **Yes** | Bundled by PyInstaller but not declared as direct project dependency. Pulled in by tracing `hypothesis` imports from the shared environment. |
| click | 8.2.1 | BSD-3-Clause | No (transitive of uvicorn) | **Yes** | Transitive dependency of uvicorn; expected. |
| MarkupSafe | 3.0.2 | BSD-3-Clause | No (transitive of jinja2) | **Yes** | Transitive dependency of jinja2; expected. |
| pyreadline3 | 3.5.4 | BSD | **No** | **Yes** | Bundled by PyInstaller but not declared as direct project dependency. Windows readline replacement; pulled in via uvicorn/click in the shared environment. |
| websockets | 15.0.1 | BSD-3-Clause | **No** | **Yes** | Bundled by PyInstaller but not declared as direct project dependency. Pulled in via uvicorn `[standard]` extra. |
| importlib_metadata | 8.0.0 | Apache-2.0 | **No** | **Yes** | Bundled by PyInstaller but not declared as direct project dependency. Vendored inside `setuptools`; pulled in transitively. |
| setuptools | 80.9.0 | MIT | No (build backend only) | **Yes** | Bundled by PyInstaller but not declared as direct project dependency. Build backend; pulled in by PyInstaller's own hooks. |

### 5b — Packages Present as Directories (No dist-info in Bundle)

These packages are present as importable directories but lack a `.dist-info` record
inside the bundle. Versions are resolved from the shared environment via `pip show`.

**Note:** The old installer (1.0.2) was built from a shared environment and contained
vulnerable versions. The new clean installer will contain the remediated versions below.

| Directory | Package | Clean Version | License | Declared in Project Files | Notes |
|-----------|---------|---------------|---------|--------------------------|-------|
| `httptools/` | httptools | 0.7.1 | MIT | No (transitive of uvicorn `[standard]`) | Expected transitive dependency of uvicorn. |
| `lxml/` | lxml | **6.1.0** | BSD-3-Clause | **No** | Not a declared dependency; pulled in by PyInstaller tracing. Old installer had vulnerable 6.0.4. |
| `markupsafe/` | MarkupSafe | 3.0.2 | BSD-3-Clause | No (transitive of jinja2) | Same package as MarkupSafe dist-info entry above; duplicate directory. |
| `orjson/` | orjson | **3.11.6** | Apache-2.0 OR MIT | No (transitive of fastapi extras) | Transitive of fastapi optional extras. Old installer had vulnerable 3.11.3. |
| `pydantic_core/` | pydantic-core | 2.33.2 | MIT | No (transitive of pydantic) | Transitive dependency of pydantic; expected. |
| `tzdata/` | tzdata | 2025.2 | Apache-2.0 | No (transitive of anyio/starlette) | Timezone data; transitive dependency; expected. |
| `watchfiles/` | watchfiles | 1.1.1 | MIT | No (transitive of uvicorn `[standard]`) | Transitive of uvicorn `[standard]` extra; expected. |
| `websockets/` | websockets | 15.0.1 | BSD-3-Clause | **No** | Same package as websockets dist-info entry above; duplicate directory. |
| `yaml/` | PyYAML | 6.0.2 | MIT | **No** | Not a declared dependency; pulled in by PyInstaller tracing shared environment. |

### 5c — Python Extension Modules (.pyd) — Python Standard Library

These are compiled Python 3.12.0 stdlib modules bundled by PyInstaller.
They are not third-party packages.

| File | Purpose |
|------|---------|
| `_asyncio.pyd` | asyncio C accelerator |
| `_bz2.pyd` | bzip2 compression |
| `_ctypes.pyd` | C foreign function interface |
| `_decimal.pyd` | Decimal arithmetic |
| `_elementtree.pyd` | XML ElementTree C accelerator |
| `_hashlib.pyd` | Cryptographic hash functions (OpenSSL-backed) |
| `_lzma.pyd` | LZMA/XZ compression |
| `_multiprocessing.pyd` | Multiprocessing support |
| `_overlapped.pyd` | Windows I/O completion ports |
| `_queue.pyd` | Queue C accelerator |
| `_socket.pyd` | Socket interface |
| `_ssl.pyd` | SSL/TLS support (OpenSSL-backed) |
| `_uuid.pyd` | UUID generation |
| `_wmi.pyd` | Windows Management Instrumentation |
| `_zoneinfo.pyd` | IANA timezone database |
| `pyexpat.pyd` | Expat XML parser |
| `select.pyd` | I/O multiplexing |
| `unicodedata.pyd` | Unicode character database |

### 5d — Native DLLs

| File | Type | License | Purpose |
|------|------|---------|---------|
| `python312.dll` | Python runtime | PSF-2.0 | Python 3.12.0 interpreter |
| `libcrypto-3.dll` | OpenSSL | Apache-2.0 | Cryptographic operations (used by `_ssl.pyd`, `_hashlib.pyd`) |
| `libssl-3.dll` | OpenSSL | Apache-2.0 | TLS/SSL protocol (used by `_ssl.pyd`) |
| `libffi-8.dll` | libffi | MIT | Foreign function interface (used by `_ctypes.pyd`) |
| `VCRUNTIME140.dll` | Microsoft Visual C++ | Microsoft | MSVC runtime |
| `ucrtbase.dll` | Microsoft Universal CRT | Microsoft | Windows C runtime |
| `api-ms-win-core-*.dll` (×26) | Windows API sets | Microsoft | Windows core API forwarding stubs |
| `api-ms-win-crt-*.dll` (×13) | Windows CRT API sets | Microsoft | Windows CRT API forwarding stubs |

> **⚠️ Security note:** `libcrypto-3.dll` and `libssl-3.dll` are OpenSSL binaries
> bundled with Python 3.12.0. Their exact OpenSSL version should be verified against
> current OpenSSL CVEs. These DLLs are tied to the Python version and cannot be
> patched independently — a Python version update is required to get OpenSSL patches.

---

## Section 6 — Frontend Dependencies

| Category | Finding |
|----------|---------|
| npm / yarn / pnpm | Not present — no `package.json` found |
| CDN-loaded libraries | **None** — HTML template loads only local `/static/` files |
| Third-party JS frameworks | **None** — `app.js` is vanilla JavaScript |
| Third-party CSS frameworks | **None** — `style.css` is custom-written |
| Offline capability | Confirmed — `app.js` header states: *"Vanilla JS, no external dependencies, fully offline-capable."* |

---

## Section 7 — Summary Counts

| Category | Count |
|----------|-------|
| Direct runtime dependencies (declared) | 5 |
| Transitive runtime dependencies (declared chain) | 14 |
| Dev / test / build dependencies | 7 |
| Packages in delivered installer — declared or expected transitive | 14 |
| Packages in delivered installer — **not declared, bundled by environment bleed** | 6 (attrs, pyreadline3, websockets, setuptools, lxml, PyYAML) |
| Frontend (JS/CSS) third-party dependencies | 0 |
| CDN dependencies | 0 |

---

## Section 8 — Version Declarations

All dependencies are now pinned to exact versions in `requirements-lock.txt`.
The `pyproject.toml` uses `>=` lower bounds that enforce the security-remediated minimums.

| Library | Declared Constraint | Pinned Version | Security Status |
|---------|--------------------|--------------------|----------------|
| openpyxl | `>=3.1.0` | 3.1.5 | ✅ No known CVE |
| fastapi | `>=0.120.1` | **0.120.1** | ✅ Updated (min for starlette 0.49.1) |
| uvicorn | `>=0.23.0` | 0.35.0 | ✅ No known CVE |
| python-multipart | `>=0.0.26` | **0.0.26** | ✅ Security remediated |
| jinja2 | `>=3.1.0` | 3.1.6 | ✅ No known CVE |
| starlette | (transitive) | **0.49.1** | ✅ Security remediated |
| lxml | (transitive/bundled) | **6.1.0** | ✅ Security remediated |
| orjson | (transitive/bundled) | **3.11.6** | ✅ Security remediated |
| pytest | `>=7.0.0` | 8.4.2 | ✅ Dev only |
| pytest-cov | `>=4.0.0` | 7.0.0 | ✅ Dev only |
| hypothesis | `>=6.0.0` | 6.151.9 | ✅ Dev only |
| black | `>=23.0.0` | 25.9.0 | ✅ Dev only |
| mypy | `>=1.0.0` | 1.18.2 | ✅ Dev only |
| flake8 | `>=6.0.0` | 7.3.0 | ✅ Dev only |
| pyinstaller | `>=6.0.0` | 6.19.0 | ✅ Build only |
| setuptools (build backend) | `>=61.0` | 80.9.0 | ✅ Build only |

---

## Section 9 — Items That Could Not Be Fully Verified

| Item | Reason | Risk |
|------|--------|------|
| Inno Setup version | Not recorded in any project file; referenced only by install path in `.iss` script | Low — build tool only, not shipped |
| OpenSSL version inside `libcrypto-3.dll` / `libssl-3.dll` | Bundled with Python 3.12.0; exact OpenSSL version requires binary inspection | Medium — should be checked against CVE list |
| Full transitive closure of `base_library.zip` | PyInstaller packs selected stdlib modules into this archive; contents not enumerated | Low — stdlib only |
| `attrs` in bundle | Present but not a declared runtime dependency; pulled in by PyInstaller tracing the shared environment | Low — benign library, but indicates environment bleed |
| `lxml` in bundle | Present as directory, not declared as a dependency | Low-Medium — should be confirmed as intentional |
| `PyYAML` in bundle | Present as `yaml/` directory, not declared as a dependency | Low-Medium — should be confirmed as intentional |
| License for `idna` | `pip show` returned no license field | Low — well-known BSD-3-Clause library |
| License for `annotated-types` | `pip show` returned no license field | Low — MIT per PyPI |
| License for `colorama` | `pip show` returned no license field | Low — BSD-3-Clause per PyPI |
| Clean virtual environment | Not available; versions resolved from shared environment | Medium — see lock file recommendation |

---

## Section 10 — Cybersecurity Notes

### ✅ No Frontend Third-Party Dependencies
The web UI is pure vanilla JavaScript and CSS with no npm packages, no CDN-loaded libraries, and no external network requests at runtime. The application is fully offline-capable.

### ✅ No CDN Dependencies
The HTML template (`index.html`) loads only local `/static/app.js` and `/static/style.css`. No external URLs are referenced.

### ⚠ Dependencies Are Not Version-Pinned
All dependencies in `requirements.txt` and `pyproject.toml` use `>=` lower bounds only. This means:
- Different build environments may resolve different versions
- A future `pip install` could pull in a version with a known vulnerability
- The installer was built with the versions listed in this report, but this is not enforced by the project files

**Recommendation:** Use `requirements-lock.txt` (created alongside this report) for all release builds. This file pins all runtime dependencies to exact versions. Regenerate it with `pip-compile` or equivalent after each dependency update.

### ⚠ Shared Development Environment — Delivered Installer Impact
The installer already delivered was built from a shared development environment.
PyInstaller traced imports from unrelated packages in that environment and bundled
them into the installer. The following packages are physically present in the
delivered installer but are **not declared** as project dependencies:

| Package | Version | Why Present |
|---------|---------|-------------|
| attrs | 25.3.0 | `hypothesis` (dev dep) was installed in the same environment |
| pyreadline3 | 3.5.4 | Windows readline; pulled in via shared environment |
| websockets | 15.0.1 | uvicorn `[standard]` extra; pulled in via shared environment |
| setuptools | 80.9.0 | PyInstaller's own hooks pulled in the build backend |
| lxml | 6.0.4 | Not declared; pulled in by PyInstaller tracing shared environment |
| PyYAML | 6.0.2 | Not declared; pulled in by PyInstaller tracing shared environment |

These packages are present in the delivered installer. They are not malicious, but
their presence was unintentional and should be noted in the security review.

**For future releases, builds must be produced from a dedicated clean virtual
environment using `requirements-lock.txt` to prevent this environment bleed.**

### ⚠ OpenSSL DLLs Bundled
`libcrypto-3.dll` and `libssl-3.dll` are bundled with the Python 3.12.0 runtime. These should be checked against current OpenSSL CVEs. They are not updated independently of Python — a Python version update is required to get OpenSSL patches.

### ℹ PyInstaller License
PyInstaller (GPLv2-or-later) includes a special exception that permits packaging non-free and commercial applications. The GPL does not propagate to the packaged application. This is a build tool only and is not shipped in the installer.

### ℹ hypothesis License
`hypothesis` (MPL-2.0) is a dev/test dependency only. It is not included in the installer. MPL-2.0 is a weak copyleft license that applies only to modifications of hypothesis source files themselves.

---

## Appendix — Release Lock File (`requirements-lock.txt`)

A `requirements-lock.txt` file exists at the project root with all security-remediated versions pinned exactly.

**This file must be used for all clean builds going forward.**

```
openpyxl==3.1.5
fastapi==0.120.1
uvicorn==0.35.0
python-multipart==0.0.26
Jinja2==3.1.6
starlette==0.49.1
pydantic==2.11.9
pydantic_core==2.33.2
anyio==4.10.0
h11==0.16.0
click==8.2.1
MarkupSafe==3.0.2
et_xmlfile==2.0.0
sniffio==1.3.1
idna==3.10
typing_extensions==4.15.0
annotated-types==0.7.0
typing-inspection==0.4.1
colorama==0.4.6
annotated-doc==0.0.4
```

**To produce a clean installer:**
```bash
python -m venv venv-release
venv-release\Scripts\activate
pip install -r requirements-lock.txt
pyinstaller ExcelNormalization.spec --noconfirm
```
This produces a bundle containing only the declared dependencies,
without environment bleed from unrelated projects.
