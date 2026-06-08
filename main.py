import io
import asyncio
import os
import sys
import traceback
import uuid
import zipfile
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, Response

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from Altomizer.alt_management import (
    build_alt_excel,
    build_alt_inventory,
    build_alt_preview_entry,
    build_alt_preview_images,
    generate_missing_alt_rows_with_claude,
    generate_missing_alt_rows_with_copilot,
    generate_missing_alt_rows_with_gemini,
    generate_missing_alt_rows_with_groq,
    generate_missing_alt_rows_with_openrouter,
    normalize_alt_text,
    prepare_docx_preview_context,
    summarize_alt_rows,
)
from Altomizer.docx_tools import (
    extract_excel_images,
    inject_alt_texts_into_docx,
    make_grids,
    parse_alt_injection_workbook,
    safe_download_stem,
)
from Altomizer.color_correction import (
    build_color_correction_output_name,
    normalize_hex_color,
    process_docx_bytes,
)
from Altomizer.list_correction import (
    ListProcessor,
    build_list_correction_output_name,
)
from Altomizer.excel_merger import (
    build_excel_merge_output_name,
    merge_excel_workbooks,
    supported_excel_filename,
)
from Altomizer.pdf_alt_tools import (
    build_pdf_alt_inventory,
    build_pdf_preview_images,
    inject_pdf_alt_texts,
    validate_pdf_file,
)

app = FastAPI(title="HBS Alto", version="1.0.0")

BASE_DIR = Path(__file__).resolve().parent
TEMP_DIR = BASE_DIR / "temp"
TEMP_DIR.mkdir(parents=True, exist_ok=True)
ALT_SESSIONS: dict[str, dict] = {}
PDF_ALT_SESSIONS: dict[str, dict] = {}
LIST_CORRECTION_SESSIONS: dict[str, dict] = {}
COLOR_CORRECTION_SESSIONS: dict[str, dict] = {}
EXCEL_MERGER_SESSIONS: dict[str, dict] = {}


def load_local_env() -> None:
    env_path = BASE_DIR / ".env"
    if not env_path.exists():
        return

    try:
        for raw_line in env_path.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            if not key:
                continue
            cleaned = value.strip().strip('"').strip("'")
            os.environ.setdefault(key, cleaned)
    except OSError:
        return


def env_flag(name: str, default: bool) -> bool:
    raw_value = str(os.getenv(name, "1" if default else "0")).strip().lower()
    return raw_value not in {"0", "false", "no", "off", ""}


load_local_env()
WEB_UI_ENABLED = env_flag("ALTOMIZER_WEB_UI_ENABLED", True)

DESKTOP_ONLY_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>HBS Alto Desktop Only</title>
  <style>
    :root {
      color-scheme: light;
      font-family: "Segoe UI", "Aptos", sans-serif;
    }
    * {
      box-sizing: border-box;
    }
    body {
      margin: 0;
      min-height: 100vh;
      display: grid;
      place-items: center;
      padding: 28px;
      background:
        radial-gradient(circle at top right, rgba(29, 107, 87, 0.12), transparent 28%),
        radial-gradient(circle at top left, rgba(196, 167, 120, 0.1), transparent 32%),
        linear-gradient(180deg, #f8f3ea 0%, #f3ebdd 100%);
      background-attachment: fixed;
      background-repeat: no-repeat;
      background-size: cover;
      color: #1f2a1f;
    }
    .notice {
      width: min(560px, 100%);
      padding: 28px 30px;
      border-radius: 26px;
      border: 1px solid rgba(201, 180, 148, 0.58);
      background: rgba(255, 251, 245, 0.96);
      box-shadow: 0 22px 54px rgba(77, 55, 24, 0.12);
    }
    h1 {
      margin: 0 0 10px;
      font-size: clamp(2rem, 3vw, 2.7rem);
      line-height: 1;
      color: #2f2418;
    }
    p {
      margin: 0;
      color: #5e665c;
      line-height: 1.6;
      font-size: 1rem;
    }
    code {
      padding: 0.12rem 0.36rem;
      border-radius: 999px;
      background: rgba(29, 107, 87, 0.08);
      color: #154f40;
      font-family: Consolas, "Courier New", monospace;
      font-size: 0.96rem;
    }
  </style>
</head>
<body>
  <main class="notice">
    <h1>HBS Alto Desktop Only</h1>
    <p>This deployment has the browser interface disabled so the working UI is not exposed in DevTools. Run the native app instead with <code>python Altomizer/run.py</code> or use the packaged desktop executable.</p>
  </main>
</body>
</html>
"""


def desktop_only_response() -> HTMLResponse:
    return HTMLResponse(DESKTOP_ONLY_HTML)


def browser_ui_response(html: str) -> HTMLResponse:
    if not WEB_UI_ENABLED:
        return desktop_only_response()
    return HTMLResponse(html)


@app.middleware("http")
async def add_response_security_headers(request: Request, call_next):
    response = await call_next(request)
    response.headers["Cache-Control"] = "no-store, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "no-referrer"
    response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
    if response.headers.get("content-type", "").startswith("text/html"):
        response.headers["Content-Security-Policy"] = "default-src 'self' 'unsafe-inline' data: blob:; img-src 'self' data: blob:; frame-ancestors 'none'; base-uri 'self'; form-action 'self'"
    return response


def write_upload(session_dir: Path, upload: UploadFile) -> Path:
    filename = Path(upload.filename or "upload.bin").name
    target_path = session_dir / filename
    upload.file.seek(0)
    with target_path.open("wb") as handle:
        handle.write(upload.file.read())
    return target_path


DOCX_REQUIRED_PARTS = frozenset({"[Content_Types].xml", "_rels/.rels", "word/document.xml"})


def read_file_head(path: Path, length: int = 512) -> bytes:
    try:
        with path.open("rb") as handle:
            return handle.read(length)
    except OSError:
        return b""


def invalid_docx_signature_message(path: Path, head: bytes) -> str:
    filename = path.name
    if not head:
        return f"{filename} could not be read or is empty. Upload a saved Word DOCX file."

    stripped = head.lstrip()
    if head.startswith(b"PK"):
        return (
            f"{filename} starts like a DOCX package, but it is incomplete or corrupt. "
            "DOCX files are ZIP containers internally, and this one is missing the ZIP directory Word needs. "
            "Re-download the file, or open it in Word and use File > Save As > Word Document (*.docx), then upload that saved copy."
        )
    if head.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return (
            f"{filename} looks like a legacy .doc or an encrypted/protected Office file, not a normal DOCX package. "
            "Open it in Word and save a fresh Word Document (*.docx), then upload that copy."
        )
    if head.startswith(b"%PDF"):
        return f"{filename} is a PDF file. HBS Alto needs the source Word DOCX for ALT inventory."
    if stripped.startswith((b"<!DOCTYPE", b"<html", b"<?xml", b"<HTML")):
        return (
            f"{filename} looks like a web/XML file instead of a Word DOCX package. "
            "Download the actual DOCX file and upload that copy."
        )
    return (
        f"{filename} is not a valid Word DOCX package. "
        "Open it in Word and save a fresh Word Document (*.docx), then upload that copy."
    )


def docx_package_validation_error(path: Path) -> str | None:
    head = read_file_head(path)
    if not zipfile.is_zipfile(path):
        return invalid_docx_signature_message(path, head)

    try:
        with zipfile.ZipFile(path, "r") as archive:
            names = set(archive.namelist())
    except zipfile.BadZipFile:
        return invalid_docx_signature_message(path, head)
    except RuntimeError as exc:
        return f"{path.name} could not be opened as a DOCX package: {exc}"

    missing_parts = sorted(DOCX_REQUIRED_PARTS.difference(names))
    if missing_parts:
        inner_docx = sorted(name for name in names if name.lower().endswith(".docx"))
        if inner_docx:
            return (
                f"{path.name} is a ZIP archive containing a DOCX, not the DOCX package itself. "
                f"Extract {Path(inner_docx[0]).name} and upload that file."
            )
        return (
            f"{path.name} is a ZIP archive, but it is missing required Word DOCX parts: "
            f"{', '.join(missing_parts)}."
        )

    return None


def validate_uploaded_docx(source_path: Path) -> None:
    error_message = docx_package_validation_error(source_path)
    if error_message:
        raise HTTPException(status_code=422, detail=error_message)


def preview_response(preview_entry: dict) -> Response:
    ext = str(preview_entry.get("ext", "png")).lower()
    media_type = "image/jpeg" if ext in {"jpg", "jpeg"} else f"image/{ext}"
    return Response(content=preview_entry.get("bytes", b""), media_type=media_type)


def build_equation_previews_for_session(session_id: str) -> None:
    session = ALT_SESSIONS.get(session_id)
    if not session:
        return

    session["equation_previews_status"] = "processing"
    try:
        pdf_path = None
        session["pdf_path"] = None
        preview_images = session.setdefault("preview_images", {})
        for row in session.get("rows", []):
            row_id = row.get("id")
            if str(row.get("role", "")).lower() != "equation" or not isinstance(row_id, int):
                continue
            if row_id in preview_images:
                continue
            try:
                preview_entry = build_alt_preview_entry(row, session.get("source_path"), pdf_path)
            except Exception as exc:
                session.setdefault("equation_preview_errors", {})[row_id] = str(exc)
                continue
            if preview_entry is not None:
                preview_images[row_id] = preview_entry
        session["equation_previews_status"] = "ready"
    except Exception as exc:
        session["equation_previews_status"] = "failed"
        session["equation_previews_error"] = str(exc)


@app.get("/list-correction", response_class=HTMLResponse)
async def list_correction_page():
    return browser_ui_response(
        """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>HBS Alto - List Correction</title>
  <style>
    :root {
      --bg: #f4efe7;
      --panel: rgba(255, 251, 245, 0.94);
      --line: #d9cbb9;
      --ink: #1f2a1f;
      --muted: #6b7166;
      --accent: #1d6b57;
      --accent-soft: #d8efe7;
      --warn: #9a3412;
      --warn-soft: #ffedd5;
      --shadow: 0 20px 45px rgba(58, 43, 24, 0.08);
      font-family: "Segoe UI", "Aptos", sans-serif;
    }

    * {
      box-sizing: border-box;
    }

    .is-hidden {
      display: none !important;
    }

    body {
      margin: 0;
      background:
        radial-gradient(circle at top right, rgba(29, 107, 87, 0.12), transparent 28%),
        radial-gradient(circle at top left, rgba(196, 167, 120, 0.1), transparent 32%),
        linear-gradient(180deg, #f8f3ea 0%, #f3ebdd 100%);
      background-attachment: fixed;
      background-repeat: no-repeat;
      background-size: cover;
      color: var(--ink);
    }

    body.guard-locked {
      overflow: hidden;
    }

    body.guard-locked .shell {
      pointer-events: none;
      user-select: none;
      filter: blur(10px) saturate(0.88);
    }

    body.guard-locked {
      overflow: hidden;
    }

    body.guard-locked .shell {
      pointer-events: none;
      user-select: none;
      filter: blur(10px) saturate(0.88);
    }

    body.guard-locked {
      overflow: hidden;
    }

    body.guard-locked .shell {
      pointer-events: none;
      user-select: none;
      filter: blur(10px) saturate(0.88);
    }

    .shell {
      max-width: 1240px;
      margin: 0 auto;
      padding: 28px 20px 40px;
    }

    .panel {
      background: var(--panel);
      border: 1px solid rgba(217, 203, 185, 0.8);
      border-radius: 24px;
      box-shadow: var(--shadow);
    }

    .appbar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 20px;
      padding: 24px 32px;
      margin-bottom: 18px;
      background: linear-gradient(180deg, rgba(255, 250, 244, 0.96) 0%, rgba(249, 241, 229, 0.98) 100%);
      border: 1px solid rgba(201, 180, 148, 0.58);
      border-radius: 24px;
      box-shadow: 0 22px 54px rgba(77, 55, 24, 0.12);
      color: #35261a;
    }

    .appbar-left,
    .appbar-right {
      display: flex;
      align-items: center;
      gap: 16px;
      min-width: 0;
    }

    .brandmark {
      font-size: 2.48rem;
      line-height: 1;
      font-weight: 800;
      letter-spacing: -0.06em;
      color: #2f2418;
      white-space: nowrap;
    }

    .header-action {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 13px 18px;
      border-radius: 14px;
      color: rgba(61, 42, 24, 0.78);
      text-decoration: none;
      transition: background 140ms ease, color 140ms ease;
    }

    .header-action:hover,
    .header-action.is-current {
      background: rgba(29, 107, 87, 0.1);
      color: #1f2a1f;
    }

    .menu-node {
      position: relative;
    }

    .tool-menu-trigger {
      appearance: none;
      border: 0;
      background: transparent;
      cursor: pointer;
      font: inherit;
      min-width: 50px;
      min-height: 50px;
    }

    .devtools-shield {
      position: fixed;
      inset: 0;
      display: none;
      align-items: center;
      justify-content: center;
      padding: 24px;
      background: rgba(16, 20, 20, 0.78);
      backdrop-filter: blur(10px);
      z-index: 9999;
    }

    .devtools-shield.is-active {
      display: flex;
    }

    .devtools-shield-card {
      max-width: 420px;
      padding: 24px 26px;
      border-radius: 22px;
      border: 1px solid rgba(255, 255, 255, 0.12);
      background: rgba(18, 27, 24, 0.94);
      box-shadow: 0 30px 80px rgba(0, 0, 0, 0.32);
      color: #f6efe5;
      text-align: center;
      display: grid;
      gap: 10px;
    }

    .devtools-shield-card strong {
      font-size: 1.18rem;
      font-weight: 800;
    }

    .devtools-shield-card span {
      color: rgba(246, 239, 229, 0.8);
      line-height: 1.5;
    }

    .tool-menu-panel {
      position: absolute;
      top: calc(100% + 10px);
      right: 0;
      min-width: 220px;
      padding: 14px;
      border-radius: 18px;
      background: rgba(31, 42, 31, 0.96);
      border: 1px solid rgba(255, 255, 255, 0.08);
      box-shadow: 0 20px 45px rgba(20, 24, 20, 0.22);
      display: none;
      z-index: 20;
    }

    .tool-menu-panel.is-open {
      display: grid;
      gap: 10px;
    }

    .tool-link {
      display: block;
      padding: 10px 12px;
      border-radius: 12px;
      color: #f4efe7;
      text-decoration: none;
      transition: background 140ms ease;
    }

    .tool-link:hover,
    .tool-link.is-current {
      background: rgba(255, 255, 255, 0.08);
    }

    .hero {
      padding: 24px;
      margin-bottom: 18px;
    }

    .hero h1 {
      margin: 0 0 8px;
      font-size: clamp(2rem, 3vw, 3rem);
      line-height: 1;
    }

    .hero p {
      margin: 0;
      color: var(--muted);
      max-width: 820px;
      line-height: 1.5;
    }

    .layout {
      display: grid;
      grid-template-columns: minmax(280px, 340px) minmax(0, 1fr);
      gap: 18px;
      align-items: start;
    }

    .sidebar,
    .workspace {
      padding: 16px;
    }

    .workspace {
      display: grid;
      gap: 16px;
      align-content: start;
    }

    .stack {
      display: grid;
      gap: 12px;
    }

    .card {
      border: 1px solid var(--line);
      border-radius: 18px;
      padding: 18px;
      background: rgba(255, 255, 255, 0.55);
      display: grid;
      gap: 12px;
    }

    .card h2,
    .card h3,
    .section-title {
      margin: 0;
      font-size: 1rem;
    }

    .card p,
    .status,
    .note {
      margin: 0;
      color: var(--muted);
      font-size: 0.95rem;
      line-height: 1.5;
    }

    .controls {
      display: grid;
      gap: 14px;
    }

    .field-block {
      display: grid;
      gap: 8px;
    }

    .field-block strong {
      display: block;
      font-size: 0.95rem;
      color: #1f2a1f;
    }

    .field-block span {
      display: block;
      color: var(--muted);
      font-size: 0.86rem;
      line-height: 1.4;
    }

    .file-picker {
      display: grid;
      gap: 8px;
      cursor: pointer;
    }

    .file-input {
      position: absolute;
      width: 1px;
      height: 1px;
      padding: 0;
      margin: -1px;
      overflow: hidden;
      clip: rect(0, 0, 0, 0);
      white-space: nowrap;
      border: 0;
    }

    .file-shell {
      display: flex;
      align-items: center;
      gap: 10px;
      min-width: 0;
      padding: 8px;
      border-radius: 16px;
      border: 1px solid rgba(201, 180, 148, 0.9);
      background: rgba(255, 255, 255, 0.86);
      transition: border-color 140ms ease, box-shadow 140ms ease, transform 140ms ease;
    }

    .file-picker:hover .file-shell,
    .file-picker:focus-within .file-shell,
    .file-shell.has-file {
      border-color: rgba(29, 107, 87, 0.42);
      box-shadow: 0 10px 24px rgba(17, 75, 60, 0.08);
      transform: translateY(-1px);
    }

    .file-trigger {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      flex: 0 0 auto;
      min-height: 40px;
      padding: 0 14px;
      border-radius: 12px;
      background: linear-gradient(180deg, #fffaf3 0%, #f2e4cf 100%);
      border: 1px solid rgba(201, 180, 148, 0.96);
      color: #2f2418;
      font-size: 0.9rem;
      font-weight: 700;
      box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.7);
    }

    .file-name {
      min-width: 0;
      color: #7b7368;
      font-size: 0.92rem;
      line-height: 1.3;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .file-shell.has-file .file-name {
      color: #1f2a1f;
      font-weight: 600;
    }

    .toggle {
      display: flex;
      align-items: start;
      gap: 12px;
      padding: 14px;
      border: 1px solid var(--line);
      border-radius: 16px;
      background: rgba(255, 255, 255, 0.78);
    }

    .toggle input {
      margin-top: 4px;
      accent-color: #1d6b57;
    }

    .toggle strong {
      display: block;
      font-size: 0.95rem;
      margin-bottom: 4px;
    }

    .toggle span {
      display: block;
      color: var(--muted);
      font-size: 0.88rem;
      line-height: 1.45;
    }

    button,
    .button-link {
      appearance: none;
      border: 0;
      border-radius: 14px;
      min-height: 52px;
      padding: 13px 16px;
      background: linear-gradient(135deg, #1d6b57 0%, #144e3f 100%);
      color: #fff8f2;
      font: inherit;
      font-weight: 600;
      cursor: pointer;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      transition: transform 140ms ease, box-shadow 140ms ease, opacity 140ms ease;
      box-shadow: 0 12px 28px rgba(17, 75, 60, 0.18);
    }

    button:hover:not(:disabled),
    .button-link:hover {
      transform: translateY(-1px);
      box-shadow: 0 16px 34px rgba(17, 75, 60, 0.2);
    }

    button:disabled,
    .button-link.is-disabled {
      cursor: default;
      opacity: 0.55;
      pointer-events: none;
      transform: none;
      box-shadow: none;
    }

    .button-link.is-hidden {
      display: none;
    }

    .summary {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-top: 4px;
    }

    .pill {
      padding: 8px 12px;
      border-radius: 999px;
      background: rgba(29, 107, 87, 0.08);
      color: #1d6b57;
      font-size: 0.86rem;
      font-weight: 600;
    }

    .results-grid {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 14px;
      margin-top: 16px;
    }

    .metric {
      padding: 18px;
      border: 1px solid var(--line);
      border-radius: 18px;
      background: rgba(255, 255, 255, 0.7);
    }

    .metric strong {
      display: block;
      font-size: 2rem;
      line-height: 1;
      color: #1f2a1f;
      margin-bottom: 8px;
    }

    .metric span {
      color: var(--muted);
      font-size: 0.92rem;
    }

    .empty {
      padding: 22px;
      border: 1px dashed var(--line);
      border-radius: 18px;
      color: var(--muted);
      background: rgba(255, 255, 255, 0.5);
    }

    .checklist {
      display: grid;
      gap: 10px;
      margin-top: 0;
      color: var(--muted);
      font-size: 0.95rem;
      line-height: 1.45;
    }

    @media (max-width: 960px) {
      .layout {
        grid-template-columns: 1fr;
      }

      .results-grid {
        grid-template-columns: 1fr;
      }

      .appbar {
        flex-direction: column;
        align-items: start;
      }

      .shell {
        padding: 18px 14px 26px;
      }

      .sidebar,
      .workspace {
        padding: 12px;
      }

      .card {
        padding: 16px;
      }

      .file-shell {
        gap: 8px;
        padding: 7px;
      }

      .file-trigger {
        min-height: 38px;
        padding: 0 12px;
      }
    }
  </style>
</head>
<body>
  <div id="devtoolsShield" class="devtools-shield" aria-hidden="true">
    <div class="devtools-shield-card">
      <strong>Protected view active</strong>
      <span>Close developer tools to continue using this page.</span>
    </div>
  </div>
  <div class="shell">
    <header class="appbar panel">
      <div class="appbar-left">
        <div class="brandmark">HBS Alto</div>
      </div>
      <nav class="appbar-right" aria-label="Primary">
        <div class="menu-node">
          <button id="toolMenuTrigger" class="header-action tool-menu-trigger" type="button" aria-expanded="false" aria-controls="toolMenuPanel" aria-label="Correction tools">
            <strong>&#9776;</strong>
          </button>
          <div id="toolMenuPanel" class="tool-menu-panel">
            <a class="tool-link" href="/">DOCX ALT Editor</a>
            <a class="tool-link" href="/color-correction">Color Correction</a>
            <a class="tool-link" href="/excel-merger">Excel Merger</a>
            <a class="tool-link" href="/pdf-alt-editor">PDF ALT Editor</a>
          </div>
        </div>
      </nav>
    </header>

    <section class="hero panel">
      <h1>List Correction</h1>
      <p>Normalize typed list markers into real Word list metadata, clean bullet templates, and download a corrected DOCX without touching the original file.</p>
    </section>

    <section class="layout">
      <aside class="panel sidebar">
        <div class="stack">
          <section class="card">
            <h2>Process File</h2>
            <p>Upload one DOCX to convert typed numbering and bullets into structured Word lists.</p>
            <div class="controls">
              <label class="file-picker" for="listSourceFile">
                <input id="listSourceFile" class="file-input" type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" />
                <span class="field-block">
                  <strong>Source DOCX</strong>
                  <span>Select the document whose typed list markers should become real Word lists.</span>
                </span>
                <span id="listFileShell" class="file-shell">
                  <span class="file-trigger">Choose DOCX</span>
                  <span id="listFileName" class="file-name">No file selected</span>
                </span>
              </label>
              <label class="toggle" for="includeHeaderFooter">
                <input id="includeHeaderFooter" type="checkbox" />
                <span>
                  <strong>Include headers and footers</strong>
                  Process paragraph lists found in header and footer content too.
                </span>
              </label>
              <button id="listProcessBtn" type="button">Correct Lists</button>
              <a id="listDownloadBtn" class="button-link is-hidden is-disabled" href="#" download>Download Corrected DOCX</a>
            </div>
          </section>

          <section class="card">
            <h3>Run Status</h3>
            <p id="listStatus" class="status">Ready.</p>
            <div id="listSummary" class="summary">
              <span class="pill">No session yet</span>
            </div>
          </section>
        </div>
      </aside>

      <main class="panel workspace">
        <h2 class="section-title">Correction Results</h2>
        <div id="listResults" class="empty">Process a DOCX to see how many ordered lists, bullet lists, and bullet templates were corrected.</div>

        <section class="card">
          <h3>What This Does</h3>
          <div class="checklist">
            <div>Detects typed prefixes like <code>1.</code>, <code>a)</code>, and Roman numeral markers and turns them into real Word numbering.</div>
            <div>Converts typed bullets into structured bullet lists with normalized indentation levels.</div>
            <div>Standardizes bullet templates so the corrected DOCX behaves more consistently in Word and downstream processing.</div>
          </div>
        </section>
      </main>
    </section>
  </div>

  <script>
    const listSourceFile = document.getElementById("listSourceFile");
    const includeHeaderFooter = document.getElementById("includeHeaderFooter");
    const listProcessBtn = document.getElementById("listProcessBtn");
    const listDownloadBtn = document.getElementById("listDownloadBtn");
    const listStatus = document.getElementById("listStatus");
    const listSummary = document.getElementById("listSummary");
    const listResults = document.getElementById("listResults");
    const listFileName = document.getElementById("listFileName");
    const listFileShell = document.getElementById("listFileShell");
    const toolMenuTrigger = document.getElementById("toolMenuTrigger");
    const toolMenuPanel = document.getElementById("toolMenuPanel");

    let listSession = null;

    const escapeHtml = (value) =>
      (String(value || "")).replace(/[&<>"']/g, (char) => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;",
      }[char]));

    function enableUiHardening() {
      const devtoolsShield = document.getElementById("devtoolsShield");
      const blockedShiftCombos = new Set(["i", "j", "c", "k", "e", "m"]);
      const blockedAltCombos = new Set(["i", "j", "c"]);
      const blockedPlainCombos = new Set(["u"]);

      const stopEvent = (event) => {
        event.preventDefault();
        event.stopPropagation();
        if (typeof event.stopImmediatePropagation === "function") {
          event.stopImmediatePropagation();
        }
      };

      const setGuardLocked = (locked) => {
        document.body.classList.toggle("guard-locked", locked);
        devtoolsShield?.classList.toggle("is-active", locked);
        devtoolsShield?.setAttribute("aria-hidden", locked ? "false" : "true");
      };

      const stopDevtoolsAccess = (event) => {
        const key = String(event.key || "").toLowerCase();
        const ctrlOrMeta = event.ctrlKey || event.metaKey;
        const shouldBlock =
          key === "f12" ||
          (event.shiftKey && key === "f7") ||
          (ctrlOrMeta && blockedPlainCombos.has(key)) ||
          (ctrlOrMeta && event.shiftKey && blockedShiftCombos.has(key)) ||
          (ctrlOrMeta && event.altKey && blockedAltCombos.has(key));

        if (shouldBlock) {
          stopEvent(event);
        }
      };

      const blockContextAccess = (event) => {
        stopEvent(event);
      };

      const evaluateDevtoolsState = () => {
        const widthGap = Math.max(0, window.outerWidth - window.innerWidth);
        const heightGap = Math.max(0, window.outerHeight - window.innerHeight);
        const devtoolsLikelyOpen = widthGap > 220 || heightGap > 220;
        setGuardLocked(devtoolsLikelyOpen);
      };

      window.addEventListener("keydown", stopDevtoolsAccess, true);
      document.addEventListener("keydown", stopDevtoolsAccess, true);
      document.addEventListener("contextmenu", blockContextAccess, true);
      document.addEventListener("mousedown", (event) => {
        if (event.button === 2) {
          blockContextAccess(event);
        }
      }, true);
      document.addEventListener("auxclick", (event) => {
        if (event.button === 2) {
          blockContextAccess(event);
        }
      }, true);
      window.addEventListener("resize", evaluateDevtoolsState);
      window.addEventListener("focus", evaluateDevtoolsState);
      document.addEventListener("visibilitychange", () => {
        if (!document.hidden) {
          evaluateDevtoolsState();
        }
      });
      evaluateDevtoolsState();
      window.setInterval(evaluateDevtoolsState, 1200);
    }

    function renderListSummary(stats) {
      if (!stats) {
        listSummary.innerHTML = '<span class="pill">No session yet</span>';
        return;
      }
      const pills = [
        `${stats.ordered_tagged || 0} ordered`,
        `${stats.unordered_tagged || 0} unordered`,
        `${stats.bullet_templates_standardized || 0} bullet templates`,
      ];
      listSummary.innerHTML = pills.map((label) => `<span class="pill">${escapeHtml(label)}</span>`).join("");
    }

    function renderListResults(stats) {
      if (!stats) {
        listResults.className = "empty";
        listResults.innerHTML = "Process a DOCX to see how many ordered lists, bullet lists, and bullet templates were corrected.";
        return;
      }
      listResults.className = "results-grid";
      listResults.innerHTML = `
        <article class="metric">
          <strong>${escapeHtml(stats.ordered_tagged || 0)}</strong>
          <span>Ordered lists tagged</span>
        </article>
        <article class="metric">
          <strong>${escapeHtml(stats.unordered_tagged || 0)}</strong>
          <span>Unordered lists tagged</span>
        </article>
        <article class="metric">
          <strong>${escapeHtml(stats.bullet_templates_standardized || 0)}</strong>
          <span>Bullet templates standardized</span>
        </article>
      `;
    }

    function syncListFileState() {
      const [file] = listSourceFile.files || [];
      listFileName.textContent = file ? file.name : "No file selected";
      listFileShell.classList.toggle("has-file", Boolean(file));
    }

    listProcessBtn.addEventListener("click", async () => {
      const [file] = listSourceFile.files || [];
      if (!file) {
        listStatus.textContent = "Choose a DOCX first.";
        return;
      }

      listProcessBtn.disabled = true;
      listDownloadBtn.classList.add("is-hidden", "is-disabled");
      listDownloadBtn.removeAttribute("href");
      listStatus.textContent = "Correcting list structure...";
      renderListSummary(null);
      renderListResults(null);

      try {
        const formData = new FormData();
        formData.append("file", file);
        formData.append("include_hf", includeHeaderFooter.checked ? "true" : "false");

        const response = await fetch("/api/list-correction/process", {
          method: "POST",
          body: formData,
        });
        const payload = await response.json();
        if (!response.ok) {
          throw new Error(payload.detail || "List correction failed.");
        }

        listSession = payload.session_id;
        renderListSummary(payload.stats || null);
        renderListResults(payload.stats || null);
        listStatus.textContent = payload.message || "List correction complete.";

        if (listSession) {
          listDownloadBtn.href = `/api/list-correction/session/${listSession}/download.docx`;
          listDownloadBtn.download = payload.output_filename || "list_corrected.docx";
          listDownloadBtn.classList.remove("is-hidden", "is-disabled");
        }
      } catch (error) {
        listStatus.textContent = error.message;
        listResults.className = "empty";
        listResults.innerHTML = escapeHtml(error.message);
      } finally {
        listProcessBtn.disabled = false;
      }
    });

    listSourceFile.addEventListener("change", syncListFileState);
    syncListFileState();
    enableUiHardening();

    toolMenuTrigger?.addEventListener("click", (event) => {
      event.stopPropagation();
      const isOpen = toolMenuPanel?.classList.contains("is-open");
      toolMenuPanel?.classList.toggle("is-open", !isOpen);
      toolMenuTrigger.setAttribute("aria-expanded", String(!isOpen));
    });

    document.addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof Element)) {
        toolMenuPanel?.classList.remove("is-open");
        toolMenuTrigger?.setAttribute("aria-expanded", "false");
        return;
      }
      if (!target.closest("#toolMenuTrigger") && !target.closest("#toolMenuPanel")) {
        toolMenuPanel?.classList.remove("is-open");
        toolMenuTrigger?.setAttribute("aria-expanded", "false");
      }
    });
  </script>
</body>
</html>
        """
    )


@app.get("/color-correction", response_class=HTMLResponse)
async def color_correction_page():
    return browser_ui_response(
        """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>HBS Alto - Color Correction</title>
  <style>
    :root {
      --bg: #f4efe7;
      --panel: rgba(255, 251, 245, 0.94);
      --line: #d9cbb9;
      --ink: #1f2a1f;
      --muted: #6b7166;
      --accent: #1d6b57;
      --shadow: 0 20px 45px rgba(58, 43, 24, 0.08);
      font-family: "Segoe UI", "Aptos", sans-serif;
    }

    * {
      box-sizing: border-box;
    }

    body {
      margin: 0;
      background:
        radial-gradient(circle at top right, rgba(29, 107, 87, 0.12), transparent 28%),
        radial-gradient(circle at top left, rgba(196, 167, 120, 0.1), transparent 32%),
        linear-gradient(180deg, #f8f3ea 0%, #f3ebdd 100%);
      background-attachment: fixed;
      background-repeat: no-repeat;
      background-size: cover;
      color: var(--ink);
    }

    .shell {
      max-width: 1240px;
      margin: 0 auto;
      padding: 28px 20px 40px;
    }

    .panel {
      background: var(--panel);
      border: 1px solid rgba(217, 203, 185, 0.8);
      border-radius: 24px;
      box-shadow: var(--shadow);
    }

    .appbar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 20px;
      padding: 24px 32px;
      margin-bottom: 18px;
      background: linear-gradient(180deg, rgba(255, 250, 244, 0.96) 0%, rgba(249, 241, 229, 0.98) 100%);
      border: 1px solid rgba(201, 180, 148, 0.58);
      border-radius: 24px;
      box-shadow: 0 22px 54px rgba(77, 55, 24, 0.12);
      color: #35261a;
    }

    .appbar-left,
    .appbar-right {
      display: flex;
      align-items: center;
      gap: 16px;
      min-width: 0;
    }

    .brandmark {
      font-size: 2.48rem;
      line-height: 1;
      font-weight: 800;
      letter-spacing: -0.06em;
      color: #2f2418;
      white-space: nowrap;
    }

    .header-action {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 13px 18px;
      border-radius: 14px;
      color: rgba(61, 42, 24, 0.78);
      text-decoration: none;
      transition: background 140ms ease, color 140ms ease;
    }

    .header-action:hover,
    .header-action.is-current {
      background: rgba(29, 107, 87, 0.1);
      color: #1f2a1f;
    }

    .menu-node {
      position: relative;
    }

    .tool-menu-trigger {
      appearance: none;
      border: 0;
      background: transparent;
      cursor: pointer;
      font: inherit;
      min-width: 50px;
      min-height: 50px;
    }

    .devtools-shield {
      position: fixed;
      inset: 0;
      display: none;
      align-items: center;
      justify-content: center;
      padding: 24px;
      background: rgba(16, 20, 20, 0.78);
      backdrop-filter: blur(10px);
      z-index: 9999;
    }

    .devtools-shield.is-active {
      display: flex;
    }

    .devtools-shield-card {
      max-width: 420px;
      padding: 24px 26px;
      border-radius: 22px;
      border: 1px solid rgba(255, 255, 255, 0.12);
      background: rgba(18, 27, 24, 0.94);
      box-shadow: 0 30px 80px rgba(0, 0, 0, 0.32);
      color: #f6efe5;
      text-align: center;
      display: grid;
      gap: 10px;
    }

    .devtools-shield-card strong {
      font-size: 1.18rem;
      font-weight: 800;
    }

    .devtools-shield-card span {
      color: rgba(246, 239, 229, 0.8);
      line-height: 1.5;
    }

    .tool-menu-panel {
      position: absolute;
      top: calc(100% + 10px);
      right: 0;
      min-width: 220px;
      padding: 14px;
      border-radius: 18px;
      background: rgba(31, 42, 31, 0.96);
      border: 1px solid rgba(255, 255, 255, 0.08);
      box-shadow: 0 20px 45px rgba(20, 24, 20, 0.22);
      display: none;
      z-index: 20;
    }

    .tool-menu-panel.is-open {
      display: grid;
      gap: 10px;
    }

    .tool-link {
      display: block;
      padding: 10px 12px;
      border-radius: 12px;
      color: #f4efe7;
      text-decoration: none;
      transition: background 140ms ease;
    }

    .tool-link:hover,
    .tool-link.is-current {
      background: rgba(255, 255, 255, 0.08);
    }

    .hero {
      padding: 24px;
      margin-bottom: 18px;
    }

    .hero h1 {
      margin: 0 0 8px;
      font-size: clamp(2rem, 3vw, 3rem);
      line-height: 1;
    }

    .hero p {
      margin: 0;
      color: var(--muted);
      max-width: 840px;
      line-height: 1.5;
    }

    .layout {
      display: grid;
      grid-template-columns: minmax(280px, 340px) minmax(0, 1fr);
      gap: 18px;
      align-items: start;
    }

    .sidebar,
    .workspace {
      padding: 16px;
    }

    .workspace {
      display: grid;
      gap: 16px;
      align-content: start;
    }

    .stack {
      display: grid;
      gap: 12px;
    }

    .card {
      border: 1px solid var(--line);
      border-radius: 18px;
      padding: 18px;
      background: rgba(255, 255, 255, 0.55);
      display: grid;
      gap: 12px;
    }

    .card h2,
    .card h3,
    .section-title {
      margin: 0;
      font-size: 1rem;
    }

    .card p,
    .status {
      margin: 0;
      color: var(--muted);
      font-size: 0.95rem;
      line-height: 1.5;
    }

    .controls {
      display: grid;
      gap: 14px;
    }

    .field-block {
      display: grid;
      gap: 8px;
    }

    .field-block strong {
      display: block;
      font-size: 0.95rem;
      color: #1f2a1f;
    }

    .field-block span {
      display: block;
      color: var(--muted);
      font-size: 0.86rem;
      line-height: 1.4;
    }

    .file-picker {
      display: grid;
      gap: 8px;
      cursor: pointer;
    }

    .file-input {
      position: absolute;
      width: 1px;
      height: 1px;
      padding: 0;
      margin: -1px;
      overflow: hidden;
      clip: rect(0, 0, 0, 0);
      white-space: nowrap;
      border: 0;
    }

    .file-shell {
      display: flex;
      align-items: center;
      gap: 10px;
      min-width: 0;
      padding: 8px;
      border-radius: 16px;
      border: 1px solid rgba(201, 180, 148, 0.9);
      background: rgba(255, 255, 255, 0.86);
      transition: border-color 140ms ease, box-shadow 140ms ease, transform 140ms ease;
    }

    .file-picker:hover .file-shell,
    .file-picker:focus-within .file-shell,
    .file-shell.has-file {
      border-color: rgba(29, 107, 87, 0.42);
      box-shadow: 0 10px 24px rgba(17, 75, 60, 0.08);
      transform: translateY(-1px);
    }

    .file-trigger {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      flex: 0 0 auto;
      min-height: 40px;
      padding: 0 14px;
      border-radius: 12px;
      background: linear-gradient(180deg, #fffaf3 0%, #f2e4cf 100%);
      border: 1px solid rgba(201, 180, 148, 0.96);
      color: #2f2418;
      font-size: 0.9rem;
      font-weight: 700;
      box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.7);
    }

    .file-name {
      min-width: 0;
      color: #7b7368;
      font-size: 0.92rem;
      line-height: 1.3;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .file-shell.has-file .file-name {
      color: #1f2a1f;
      font-weight: 600;
    }

    input[type="text"] {
      width: 100%;
      padding: 12px 14px;
      border-radius: 14px;
      border: 1px solid var(--line);
      background: rgba(255, 255, 255, 0.78);
      font: inherit;
      color: var(--ink);
    }

    input[type="text"]#backgroundColor {
      font-family: "Consolas", "Courier New", monospace;
      letter-spacing: 0.03em;
    }

    button,
    .button-link {
      appearance: none;
      border: 0;
      border-radius: 14px;
      min-height: 52px;
      padding: 13px 16px;
      background: linear-gradient(135deg, #1d6b57 0%, #144e3f 100%);
      color: #fff8f2;
      font: inherit;
      font-weight: 600;
      cursor: pointer;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      transition: transform 140ms ease, box-shadow 140ms ease, opacity 140ms ease;
      box-shadow: 0 12px 28px rgba(17, 75, 60, 0.18);
    }

    button:hover:not(:disabled),
    .button-link:hover {
      transform: translateY(-1px);
      box-shadow: 0 16px 34px rgba(17, 75, 60, 0.2);
    }

    button:disabled,
    .button-link.is-disabled {
      cursor: default;
      opacity: 0.55;
      pointer-events: none;
      transform: none;
      box-shadow: none;
    }

    .button-link.is-hidden {
      display: none;
    }

    .summary {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-top: 4px;
    }

    .pill {
      padding: 8px 12px;
      border-radius: 999px;
      background: rgba(29, 107, 87, 0.08);
      color: #1d6b57;
      font-size: 0.86rem;
      font-weight: 600;
    }

    .results-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 14px;
      margin-top: 16px;
    }

    .metric {
      padding: 18px;
      border: 1px solid var(--line);
      border-radius: 18px;
      background: rgba(255, 255, 255, 0.7);
    }

    .metric strong {
      display: block;
      font-size: 2rem;
      line-height: 1;
      color: #1f2a1f;
      margin-bottom: 8px;
    }

    .metric span {
      color: var(--muted);
      font-size: 0.92rem;
    }

    .empty {
      padding: 22px;
      border: 1px dashed var(--line);
      border-radius: 18px;
      color: var(--muted);
      background: rgba(255, 255, 255, 0.5);
    }

    .checklist {
      display: grid;
      gap: 10px;
      margin-top: 0;
      color: var(--muted);
      font-size: 0.95rem;
      line-height: 1.45;
    }

    @media (max-width: 960px) {
      .layout {
        grid-template-columns: 1fr;
      }

      .results-grid {
        grid-template-columns: 1fr;
      }

      .appbar {
        flex-direction: column;
        align-items: start;
      }

      .shell {
        padding: 18px 14px 26px;
      }

      .sidebar,
      .workspace {
        padding: 12px;
      }

      .card {
        padding: 16px;
      }

      .file-shell {
        gap: 8px;
        padding: 7px;
      }

      .file-trigger {
        min-height: 38px;
        padding: 0 12px;
      }
    }
  </style>
</head>
<body>
  <div id="devtoolsShield" class="devtools-shield" aria-hidden="true">
    <div class="devtools-shield-card">
      <strong>Protected view active</strong>
      <span>Close developer tools to continue using this page.</span>
    </div>
  </div>
  <div class="shell">
    <header class="appbar panel">
      <div class="appbar-left">
        <div class="brandmark">HBS Alto</div>
      </div>
      <nav class="appbar-right" aria-label="Primary">
        <div class="menu-node">
          <button id="toolMenuTrigger" class="header-action tool-menu-trigger" type="button" aria-expanded="false" aria-controls="toolMenuPanel" aria-label="Correction tools">
            <strong>&#9776;</strong>
          </button>
          <div id="toolMenuPanel" class="tool-menu-panel">
            <a class="tool-link" href="/">DOCX ALT Editor</a>
            <a class="tool-link" href="/list-correction">List Correction</a>
            <a class="tool-link" href="/excel-merger">Excel Merger</a>
            <a class="tool-link" href="/pdf-alt-editor">PDF ALT Editor</a>
          </div>
        </div>
      </nav>
    </header>

    <section class="hero panel">
      <h1>Color Correction</h1>
      <p>Scan the DOCX package for low-contrast run and style colors, then rewrite them to safer contrast values against the background you choose.</p>
    </section>

    <section class="layout">
      <aside class="panel sidebar">
        <div class="stack">
          <section class="card">
            <h2>Process File</h2>
            <p>Upload one DOCX and generate a corrected copy with improved text contrast in document runs and styles.</p>
            <div class="controls">
              <label class="file-picker" for="colorSourceFile">
                <input id="colorSourceFile" class="file-input" type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" />
                <span class="field-block">
                  <strong>Source DOCX</strong>
                  <span>Select the document whose low-contrast text colors should be normalized.</span>
                </span>
                <span id="colorFileShell" class="file-shell">
                  <span class="file-trigger">Choose DOCX</span>
                  <span id="colorFileName" class="file-name">No file selected</span>
                </span>
              </label>
              <label class="field-block" for="backgroundColor">
                <strong>Background Color</strong>
                <input id="backgroundColor" type="text" value="#FFFFFF" inputmode="text" />
                <span>Use a 6-digit hex value like <code>#FFFFFF</code> for the page/background color you want to validate against.</span>
              </label>
              <button id="colorProcessBtn" type="button">Correct Colors</button>
              <a id="colorDownloadBtn" class="button-link is-hidden is-disabled" href="#" download>Download Corrected DOCX</a>
            </div>
          </section>

          <section class="card">
            <h3>Run Status</h3>
            <p id="colorStatus" class="status">Ready.</p>
            <div id="colorSummary" class="summary">
              <span class="pill">No session yet</span>
            </div>
          </section>
        </div>
      </aside>

      <main class="panel workspace">
        <h2 class="section-title">Correction Results</h2>
        <div id="colorResults" class="empty">Process a DOCX to see how many low-contrast text elements were adjusted.</div>

        <section class="card">
          <h3>What This Does</h3>
          <div class="checklist">
            <div>Checks direct run colors against the effective background from highlights, shading, table cells, paragraph shading, drawings, VML fills, and styles.</div>
            <div>Adjusts low-contrast colors by trying darker or lighter variants first, then falling back to a safer color if needed.</div>
            <div>Updates both document content and style definitions, then packages a corrected DOCX for download.</div>
          </div>
        </section>
      </main>
    </section>
  </div>

  <script>
    const colorSourceFile = document.getElementById("colorSourceFile");
    const backgroundColor = document.getElementById("backgroundColor");
    const colorProcessBtn = document.getElementById("colorProcessBtn");
    const colorDownloadBtn = document.getElementById("colorDownloadBtn");
    const colorStatus = document.getElementById("colorStatus");
    const colorSummary = document.getElementById("colorSummary");
    const colorResults = document.getElementById("colorResults");
    const colorFileName = document.getElementById("colorFileName");
    const colorFileShell = document.getElementById("colorFileShell");
    const toolMenuTrigger = document.getElementById("toolMenuTrigger");
    const toolMenuPanel = document.getElementById("toolMenuPanel");

    let colorSession = null;

    const escapeHtml = (value) =>
      (String(value || "")).replace(/[&<>"']/g, (char) => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;",
      }[char]));

    function enableUiHardening() {
      const devtoolsShield = document.getElementById("devtoolsShield");
      const blockedShiftCombos = new Set(["i", "j", "c", "k", "e", "m"]);
      const blockedAltCombos = new Set(["i", "j", "c"]);
      const blockedPlainCombos = new Set(["u"]);

      const stopEvent = (event) => {
        event.preventDefault();
        event.stopPropagation();
        if (typeof event.stopImmediatePropagation === "function") {
          event.stopImmediatePropagation();
        }
      };

      const setGuardLocked = (locked) => {
        document.body.classList.toggle("guard-locked", locked);
        devtoolsShield?.classList.toggle("is-active", locked);
        devtoolsShield?.setAttribute("aria-hidden", locked ? "false" : "true");
      };

      const stopDevtoolsAccess = (event) => {
        const key = String(event.key || "").toLowerCase();
        const ctrlOrMeta = event.ctrlKey || event.metaKey;
        const shouldBlock =
          key === "f12" ||
          (event.shiftKey && key === "f7") ||
          (ctrlOrMeta && blockedPlainCombos.has(key)) ||
          (ctrlOrMeta && event.shiftKey && blockedShiftCombos.has(key)) ||
          (ctrlOrMeta && event.altKey && blockedAltCombos.has(key));

        if (shouldBlock) {
          stopEvent(event);
        }
      };

      const blockContextAccess = (event) => {
        stopEvent(event);
      };

      const evaluateDevtoolsState = () => {
        const widthGap = Math.max(0, window.outerWidth - window.innerWidth);
        const heightGap = Math.max(0, window.outerHeight - window.innerHeight);
        const devtoolsLikelyOpen = widthGap > 220 || heightGap > 220;
        setGuardLocked(devtoolsLikelyOpen);
      };

      window.addEventListener("keydown", stopDevtoolsAccess, true);
      document.addEventListener("keydown", stopDevtoolsAccess, true);
      document.addEventListener("contextmenu", blockContextAccess, true);
      document.addEventListener("mousedown", (event) => {
        if (event.button === 2) {
          blockContextAccess(event);
        }
      }, true);
      document.addEventListener("auxclick", (event) => {
        if (event.button === 2) {
          blockContextAccess(event);
        }
      }, true);
      window.addEventListener("resize", evaluateDevtoolsState);
      window.addEventListener("focus", evaluateDevtoolsState);
      document.addEventListener("visibilitychange", () => {
        if (!document.hidden) {
          evaluateDevtoolsState();
        }
      });
      evaluateDevtoolsState();
      window.setInterval(evaluateDevtoolsState, 1200);
    }

    function renderColorSummary(result) {
      if (!result) {
        colorSummary.innerHTML = '<span class="pill">No session yet</span>';
        return;
      }
      const pills = [
        `${result.fixed_elements || 0} fixed`,
        result.changed ? "Changes made" : "No changes needed",
        `Background ${result.background || "#FFFFFF"}`,
      ];
      colorSummary.innerHTML = pills.map((label) => `<span class="pill">${escapeHtml(label)}</span>`).join("");
    }

    function renderColorResults(result) {
      if (!result) {
        colorResults.className = "empty";
        colorResults.innerHTML = "Process a DOCX to see how many low-contrast text elements were adjusted.";
        return;
      }
      colorResults.className = "results-grid";
      colorResults.innerHTML = `
        <article class="metric">
          <strong>${escapeHtml(result.fixed_elements || 0)}</strong>
          <span>Elements corrected</span>
        </article>
        <article class="metric">
          <strong>${escapeHtml(result.changed ? "Yes" : "No")}</strong>
          <span>Document changed</span>
        </article>
      `;
    }

    function syncColorFileState() {
      const [file] = colorSourceFile.files || [];
      colorFileName.textContent = file ? file.name : "No file selected";
      colorFileShell.classList.toggle("has-file", Boolean(file));
    }

    colorProcessBtn.addEventListener("click", async () => {
      const [file] = colorSourceFile.files || [];
      if (!file) {
        colorStatus.textContent = "Choose a DOCX first.";
        return;
      }

      colorProcessBtn.disabled = true;
      colorDownloadBtn.classList.add("is-hidden", "is-disabled");
      colorDownloadBtn.removeAttribute("href");
      colorStatus.textContent = "Correcting document colors...";
      renderColorSummary(null);
      renderColorResults(null);

      try {
        const formData = new FormData();
        formData.append("file", file);
        formData.append("background", backgroundColor.value || "#FFFFFF");

        const response = await fetch("/api/color-correction/process", {
          method: "POST",
          body: formData,
        });
        const payload = await response.json();
        if (!response.ok) {
          throw new Error(payload.detail || "Color correction failed.");
        }

        colorSession = payload.session_id;
        renderColorSummary(payload.result || null);
        renderColorResults(payload.result || null);
        colorStatus.textContent = payload.message || "Color correction complete.";

        if (colorSession) {
          colorDownloadBtn.href = `/api/color-correction/session/${colorSession}/download.docx`;
          colorDownloadBtn.download = payload.output_filename || "color_corrected.docx";
          colorDownloadBtn.classList.remove("is-hidden", "is-disabled");
        }
      } catch (error) {
        colorStatus.textContent = error.message;
        colorResults.className = "empty";
        colorResults.innerHTML = escapeHtml(error.message);
      } finally {
        colorProcessBtn.disabled = false;
      }
    });

    colorSourceFile.addEventListener("change", syncColorFileState);
    syncColorFileState();
    enableUiHardening();

    toolMenuTrigger?.addEventListener("click", (event) => {
      event.stopPropagation();
      const isOpen = toolMenuPanel?.classList.contains("is-open");
      toolMenuPanel?.classList.toggle("is-open", !isOpen);
      toolMenuTrigger.setAttribute("aria-expanded", String(!isOpen));
    });

    document.addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof Element)) {
        toolMenuPanel?.classList.remove("is-open");
        toolMenuTrigger?.setAttribute("aria-expanded", "false");
        return;
      }
      if (!target.closest("#toolMenuTrigger") && !target.closest("#toolMenuPanel")) {
        toolMenuPanel?.classList.remove("is-open");
        toolMenuTrigger?.setAttribute("aria-expanded", "false");
      }
    });
  </script>
</body>
</html>
        """
    )


@app.get("/excel-merger", response_class=HTMLResponse)
async def excel_merger_page():
    return browser_ui_response(
        """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Excel Merger - HBS Alto</title>
  <style>
    :root {
      --bg: #f4efe7;
      --panel: rgba(255, 251, 245, 0.94);
      --line: #d9cbb9;
      --ink: #1f2a1f;
      --muted: #6b7166;
      --accent: #1d6b57;
      --accent-soft: #d8efe7;
      --shadow: 0 20px 45px rgba(58, 43, 24, 0.08);
      font-family: "Segoe UI", "Aptos", sans-serif;
    }

    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      background:
        radial-gradient(circle at top right, rgba(29, 107, 87, 0.12), transparent 28%),
        radial-gradient(circle at top left, rgba(196, 167, 120, 0.1), transparent 32%),
        linear-gradient(180deg, #f8f3ea 0%, #f3ebdd 100%);
      background-attachment: fixed;
      background-repeat: no-repeat;
      background-size: cover;
      color: var(--ink);
    }

    .shell { max-width: 1320px; margin: 0 auto; padding: 22px 18px 34px; }
    .panel {
      background: var(--panel);
      border: 1px solid rgba(217, 203, 185, 0.68);
      border-radius: 28px;
      box-shadow: var(--shadow);
      backdrop-filter: blur(10px);
    }
    .appbar, .hero, .sidebar, .workspace { padding: 18px; }
    .appbar {
      position: relative;
      z-index: 20;
      overflow: visible;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      margin-bottom: 18px;
    }
    .brandmark {
      font-size: clamp(2rem, 3vw, 2.8rem);
      line-height: 1;
      font-weight: 900;
      color: #2f2418;
      letter-spacing: -0.04em;
    }
    .appbar-right { display: flex; align-items: center; gap: 12px; }
    .menu-node { position: relative; z-index: 24; }
    .header-action, .tool-menu-panel a, button, .button-link {
      appearance: none;
      border: 1px solid rgba(201, 180, 148, 0.92);
      background: rgba(255, 255, 255, 0.84);
      color: var(--ink);
      border-radius: 14px;
      min-height: 44px;
      padding: 0 16px;
      font: inherit;
      font-weight: 700;
      text-decoration: none;
      cursor: pointer;
    }
    .header-action { background: transparent; border-color: transparent; min-height: 40px; padding: 0 10px; }
    .tool-menu-panel {
      position: absolute;
      right: 0;
      top: calc(100% + 10px);
      width: 280px;
      padding: 14px;
      border-radius: 22px;
      border: 1px solid rgba(201, 180, 148, 0.78);
      background: linear-gradient(180deg, rgba(255, 251, 245, 0.98) 0%, rgba(247, 239, 227, 0.98) 100%);
      color: var(--ink);
      box-shadow: 0 22px 48px rgba(77, 55, 24, 0.16);
      display: none;
      z-index: 60;
    }
    .tool-menu-panel.is-open { display: grid; gap: 10px; }
    .tool-link {
      display: block;
      padding: 12px 14px;
      border-radius: 14px;
      color: #2f2418;
      border: 1px solid rgba(201, 180, 148, 0.8);
      background: rgba(255, 255, 255, 0.82);
      text-decoration: none;
      font-weight: 700;
      box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.86);
    }
    .tool-link:hover {
      background: rgba(29, 107, 87, 0.08);
      border-color: rgba(171, 145, 108, 0.96);
    }
    .hero { margin-bottom: 18px; }
    .hero h1 { margin: 0 0 10px; font-size: clamp(2rem, 3vw, 2.7rem); line-height: 1; color: #2f2418; }
    .hero p { margin: 0; color: var(--muted); font-size: 1rem; line-height: 1.55; max-width: 72ch; }
    .layout { display: grid; grid-template-columns: minmax(300px, 340px) minmax(0, 1fr); gap: 18px; align-items: start; }
    .stack { display: grid; gap: 16px; }
    .card {
      padding: 18px;
      border: 1px solid rgba(217, 203, 185, 0.84);
      border-radius: 22px;
      background: rgba(255,255,255,0.58);
    }
    .card h2, .card h3, .workspace h2 { margin: 0 0 10px; color: #1e2b1f; }
    .card p { margin: 0; color: var(--muted); line-height: 1.5; }
    .controls, .stacked { display: grid; gap: 12px; margin-top: 14px; }
    .file-picker { display: grid; gap: 8px; cursor: pointer; }
    .file-input {
      position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px;
      overflow: hidden; clip: rect(0, 0, 0, 0); white-space: nowrap; border: 0;
    }
    .field-block { display: grid; gap: 4px; color: var(--muted); }
    .field-block strong { color: var(--ink); }
    .file-shell {
      display: flex; align-items: center; gap: 10px; min-width: 0; padding: 8px;
      border-radius: 16px; border: 1px solid rgba(201, 180, 148, 0.9); background: rgba(255,255,255,0.86);
    }
    .file-trigger {
      display: inline-flex; align-items: center; justify-content: center; min-height: 40px; padding: 0 14px;
      border-radius: 12px; background: linear-gradient(180deg, #fffaf3 0%, #f2e4cf 100%);
      border: 1px solid rgba(201,180,148,0.96); color: #2f2418; font-size: 0.9rem; font-weight: 700;
    }
    .file-name { min-width: 0; color: #7b7368; font-size: 0.92rem; line-height: 1.3; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .file-shell.has-file .file-name { color: #1f2a1f; font-weight: 600; }
    .controls > button,
    .controls > .button-link {
      width: 100%;
      justify-content: center;
      text-align: center;
      min-height: 48px;
    }
    .controls > button {
      background: linear-gradient(180deg, #fffdf9 0%, #f4ead9 100%);
      box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.88);
      transition: transform 140ms ease, box-shadow 140ms ease, border-color 140ms ease;
    }
    .controls > button:hover:not(:disabled),
    .controls > .button-link:hover:not(.is-disabled) {
      transform: translateY(-1px);
      border-color: rgba(171, 145, 108, 0.96);
      box-shadow: 0 12px 24px rgba(70, 49, 22, 0.08);
    }
    .controls > #excelMergeProcessBtn {
      background: linear-gradient(135deg, #1d6b57 0%, #2a826c 100%);
      border-color: rgba(29, 107, 87, 0.98);
      color: #fffdf8;
      box-shadow: 0 14px 26px rgba(29, 107, 87, 0.18);
    }
    .controls > #excelMergeProcessBtn:hover:not(:disabled) {
      box-shadow: 0 18px 30px rgba(29, 107, 87, 0.24);
    }
    .controls > #excelMergeDownloadBtn {
      display: inline-flex;
      align-items: center;
      background: linear-gradient(180deg, #fffdf9 0%, #f1e7d5 100%);
      color: #2f2418;
      box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.9);
    }
    .summary { display: flex; flex-wrap: wrap; gap: 8px; }
    .pill { padding: 7px 10px; border-radius: 999px; background: var(--accent-soft); color: var(--accent); font-size: 0.88rem; font-weight: 700; }
    .status {
      padding: 12px 14px; border-radius: 14px; background: rgba(255,255,255,0.6);
      border: 1px solid rgba(217, 203, 185, 0.84); color: var(--ink);
    }
    .workspace { display: grid; gap: 16px; }
    .empty {
      padding: 32px 18px; border: 1px dashed var(--line); border-radius: 18px; color: var(--muted);
      text-align: center; background: rgba(255, 255, 255, 0.45);
    }
    .results-grid { display: grid; gap: 14px; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); }
    .metric {
      padding: 16px; border: 1px solid rgba(217, 203, 185, 0.84); border-radius: 18px; background: rgba(255,255,255,0.62);
      display: grid; gap: 6px;
    }
    .metric strong { font-size: 1.5rem; line-height: 1; color: #2f2418; }
    .metric span { color: var(--muted); font-size: 0.94rem; }
    .sheet-list { display: grid; gap: 8px; color: var(--muted); font-size: 0.95rem; line-height: 1.45; }
    .sheet-chip {
      display: inline-flex; align-items: center; padding: 8px 12px; border-radius: 999px;
      background: rgba(255,255,255,0.72); border: 1px solid rgba(217, 203, 185, 0.84); color: var(--ink);
    }
    .button-link.is-hidden, .is-hidden { display: none !important; }
    .button-link.is-disabled { opacity: 0.52; pointer-events: none; }
    @media (max-width: 960px) {
      .layout { grid-template-columns: 1fr; }
      .appbar { flex-direction: column; align-items: start; }
      .shell { padding: 18px 14px 26px; }
      .sidebar, .workspace { padding: 12px; }
      .card { padding: 16px; }
    }
  </style>
</head>
<body>
  <div class="shell">
    <header class="appbar panel">
      <div class="appbar-left">
        <div class="brandmark">HBS Alto</div>
      </div>
      <nav class="appbar-right" aria-label="Primary">
        <div class="menu-node">
          <button id="toolMenuTrigger" class="header-action tool-menu-trigger" type="button" aria-expanded="false" aria-controls="toolMenuPanel" aria-label="Correction tools">
            <strong>&#9776;</strong>
          </button>
          <div id="toolMenuPanel" class="tool-menu-panel">
            <a class="tool-link" href="/">DOCX ALT Editor</a>
            <a class="tool-link" href="/list-correction">List Correction</a>
            <a class="tool-link" href="/color-correction">Color Correction</a>
            <a class="tool-link" href="/pdf-alt-editor">PDF ALT Editor</a>
          </div>
        </div>
      </nav>
    </header>

    <section class="hero panel">
      <h1>Excel Merger</h1>
      <p>Upload any number of Excel workbooks, append their worksheet rows into one combined workbook, and download the result in a single file.</p>
    </section>

    <section class="layout">
      <aside class="panel sidebar">
        <div class="stack">
          <section class="card">
            <h2>Process Files</h2>
            <p>Select multiple Excel workbooks and concatenate their worksheet rows into one merged workbook.</p>
            <div class="controls">
              <label class="file-picker" for="excelMergeFiles">
                <input id="excelMergeFiles" class="file-input" type="file" multiple accept=".xlsx,.xlsm,.xltx,.xltm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel.sheet.macroEnabled.12" />
                <span class="field-block">
                  <strong>Source Workbooks</strong>
                  <span>Every worksheet from every selected file will be appended into one merged sheet, with images carried forward too.</span>
                </span>
                <span id="excelMergeFileShell" class="file-shell">
                  <span class="file-trigger">Choose Excel Files</span>
                  <span id="excelMergeFileName" class="file-name">No files selected</span>
                </span>
              </label>
              <button id="excelMergeProcessBtn" type="button">Merge Workbooks</button>
              <a id="excelMergeDownloadBtn" class="button-link is-hidden is-disabled" href="#" download>Download Merged Excel</a>
            </div>
          </section>

          <section class="card">
            <h3>Run Status</h3>
            <p id="excelMergeStatus" class="status">Ready.</p>
            <div id="excelMergeSummary" class="summary">
              <span class="pill">No session yet</span>
            </div>
          </section>
        </div>
      </aside>

      <main class="panel workspace">
        <h2 class="section-title">Merge Results</h2>
        <div id="excelMergeResults" class="empty">Upload Excel workbooks to preview the merged sheet summary.</div>

        <section class="card">
          <h3>What This Does</h3>
          <div class="sheet-list">
            <div>Appends every worksheet row from every uploaded workbook into one combined workbook sheet.</div>
            <div>Keeps every non-empty source row in order, including repeated headers when they exist in the uploaded files.</div>
            <div>Packages the merged workbook as one downloadable <code>.xlsx</code> file.</div>
          </div>
        </section>
      </main>
    </section>
  </div>

  <script>
    const excelMergeFiles = document.getElementById("excelMergeFiles");
    const excelMergeProcessBtn = document.getElementById("excelMergeProcessBtn");
    const excelMergeDownloadBtn = document.getElementById("excelMergeDownloadBtn");
    const excelMergeStatus = document.getElementById("excelMergeStatus");
    const excelMergeSummary = document.getElementById("excelMergeSummary");
    const excelMergeResults = document.getElementById("excelMergeResults");
    const excelMergeFileName = document.getElementById("excelMergeFileName");
    const excelMergeFileShell = document.getElementById("excelMergeFileShell");
    const toolMenuTrigger = document.getElementById("toolMenuTrigger");
    const toolMenuPanel = document.getElementById("toolMenuPanel");

    let excelMergeSession = null;

    const escapeHtml = (value) =>
      (String(value || "")).replace(/[&<>"']/g, (char) => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;",
      }[char]));

    function syncExcelMergeFileState() {
      const files = Array.from(excelMergeFiles.files || []);
      if (!files.length) {
        excelMergeFileName.textContent = "No files selected";
        excelMergeFileShell.classList.remove("has-file");
        return;
      }
      excelMergeFileName.textContent = files.length === 1 ? files[0].name : `${files.length} files selected`;
      excelMergeFileShell.classList.add("has-file");
    }

    function renderExcelMergeSummary(data) {
      if (!data) {
        excelMergeSummary.innerHTML = '<span class="pill">No session yet</span>';
        return;
      }
      const pills = [
        `${data.files || 0} workbook(s)`,
        `${data.sheets || 0} source sheet(s)`,
        `${data.rows || 0} merged row(s)`,
      ];
      excelMergeSummary.innerHTML = pills.map((label) => `<span class="pill">${escapeHtml(label)}</span>`).join("");
    }

    function renderExcelMergeResults(data) {
      if (!data) {
        excelMergeResults.className = "empty";
        excelMergeResults.innerHTML = "Upload Excel workbooks to preview the merged sheet summary.";
        return;
      }
      const sheetTitles = Array.isArray(data.sheet_titles) ? data.sheet_titles : [];
      excelMergeResults.className = "";
      excelMergeResults.innerHTML = `
        <div class="results-grid">
          <article class="metric">
            <strong>${escapeHtml(data.files || 0)}</strong>
            <span>Source workbook(s)</span>
          </article>
          <article class="metric">
            <strong>${escapeHtml(data.sheets || 0)}</strong>
            <span>Source worksheet(s)</span>
          </article>
          <article class="metric">
            <strong>${escapeHtml(data.rows || 0)}</strong>
            <span>Merged row(s)</span>
          </article>
        </div>
        <section class="card">
          <h3>Output Sheet</h3>
          <div class="sheet-list">
            ${sheetTitles.length ? sheetTitles.map((title) => `<span class="sheet-chip">${escapeHtml(title)}</span>`).join("") : "<div>No sheets were copied.</div>"}
          </div>
        </section>
      `;
    }

    excelMergeProcessBtn.addEventListener("click", async () => {
      const files = Array.from(excelMergeFiles.files || []);
      if (!files.length) {
        excelMergeStatus.textContent = "Choose at least one Excel workbook first.";
        return;
      }

      excelMergeProcessBtn.disabled = true;
      excelMergeDownloadBtn.classList.add("is-hidden", "is-disabled");
      excelMergeDownloadBtn.removeAttribute("href");
      excelMergeStatus.textContent = "Merging workbooks...";
      renderExcelMergeSummary(null);
      renderExcelMergeResults(null);

      try {
        const formData = new FormData();
        files.forEach((file) => formData.append("files", file));
        const response = await fetch("/api/excel-merger/process", {
          method: "POST",
          body: formData,
        });
        const payload = await response.json();
        if (!response.ok) {
          throw new Error(payload.detail || "Excel merge failed.");
        }

        excelMergeSession = payload.session_id;
        renderExcelMergeSummary(payload.summary || null);
        renderExcelMergeResults(payload.summary || null);
        excelMergeStatus.textContent = payload.message || "Excel merge complete.";

        if (excelMergeSession) {
          excelMergeDownloadBtn.href = `/api/excel-merger/session/${excelMergeSession}/download.xlsx`;
          excelMergeDownloadBtn.download = payload.output_filename || "merged.xlsx";
          excelMergeDownloadBtn.classList.remove("is-hidden", "is-disabled");
        }
      } catch (error) {
        excelMergeStatus.textContent = error.message;
        excelMergeResults.className = "empty";
        excelMergeResults.innerHTML = escapeHtml(error.message);
      } finally {
        excelMergeProcessBtn.disabled = false;
      }
    });

    excelMergeFiles.addEventListener("change", syncExcelMergeFileState);
    syncExcelMergeFileState();

    toolMenuTrigger?.addEventListener("click", (event) => {
      event.stopPropagation();
      const isOpen = toolMenuPanel?.classList.contains("is-open");
      toolMenuPanel?.classList.toggle("is-open", !isOpen);
      toolMenuTrigger.setAttribute("aria-expanded", String(!isOpen));
    });

    document.addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof Element)) {
        toolMenuPanel?.classList.remove("is-open");
        toolMenuTrigger?.setAttribute("aria-expanded", "false");
        return;
      }
      if (!target.closest("#toolMenuTrigger") && !target.closest("#toolMenuPanel")) {
        toolMenuPanel?.classList.remove("is-open");
        toolMenuTrigger?.setAttribute("aria-expanded", "false");
      }
    });
  </script>
</body>
</html>
        """
    )


@app.get("/pdf-alt-editor", response_class=HTMLResponse)
async def pdf_alt_editor_page():
    return browser_ui_response(
        """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>PDF ALT Editor - HBS Alto</title>
  <style>
    :root {
      --bg: #f4efe7;
      --panel: rgba(255, 251, 245, 0.94);
      --line: #d9cbb9;
      --ink: #1f2a1f;
      --muted: #6b7166;
      --accent: #1d6b57;
      --accent-soft: #d8efe7;
      --warn: #9a3412;
      --warn-soft: #ffedd5;
      --shadow: 0 20px 45px rgba(58, 43, 24, 0.08);
      font-family: "Segoe UI", "Aptos", sans-serif;
    }

    * {
      box-sizing: border-box;
    }

    .is-hidden {
      display: none !important;
    }

    body {
      margin: 0;
      background:
        radial-gradient(circle at top right, rgba(29, 107, 87, 0.12), transparent 28%),
        radial-gradient(circle at top left, rgba(196, 167, 120, 0.1), transparent 32%),
        linear-gradient(180deg, #f8f3ea 0%, #f3ebdd 100%);
      background-attachment: fixed;
      background-repeat: no-repeat;
      background-size: cover;
      color: var(--ink);
    }

    .shell {
      max-width: 1280px;
      margin: 0 auto;
      padding: 28px 20px 40px;
    }

    .panel {
      border: 1px solid rgba(201, 180, 148, 0.58);
      background: var(--panel);
      box-shadow: var(--shadow);
      border-radius: 24px;
    }

    .appbar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 20px;
      padding: 25px 34px;
      margin-bottom: 18px;
      background: linear-gradient(180deg, rgba(255, 250, 244, 0.96) 0%, rgba(249, 241, 229, 0.98) 100%);
      color: #35261a;
      overflow: visible;
    }

    .brandmark {
      font-size: 2.6rem;
      line-height: 1;
      font-weight: 800;
      letter-spacing: -0.06em;
      color: #2f2418;
      white-space: nowrap;
    }

    .appbar-right {
      display: flex;
      align-items: center;
      gap: 14px;
    }

    .header-actions {
      flex-wrap: wrap;
      justify-content: flex-end;
      gap: 14px;
    }

    .header-action,
    .tool-link,
    .button-link,
    button {
      appearance: none;
      border: 0;
      border-radius: 999px;
      padding: 11px 16px;
      background: var(--accent);
      color: white;
      font: inherit;
      font-weight: 700;
      cursor: pointer;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      justify-content: center;
    }

    .header-action {
      position: relative;
      background: transparent;
      color: rgba(61, 42, 24, 0.72);
      border: 1px solid transparent;
    }

    .header-action::after {
      content: "";
      position: absolute;
      left: 12px;
      right: 12px;
      bottom: 6px;
      height: 2px;
      border-radius: 999px;
      background: transparent;
      transition: background 140ms ease;
    }

    .header-action:hover:not(:disabled)::after {
      background: rgba(29, 107, 87, 0.32);
    }

    .header-action.is-open {
      border-color: rgba(201, 180, 148, 0.7);
      background: rgba(255, 255, 255, 0.45);
    }

    .header-action:disabled {
      opacity: 0.48;
      cursor: not-allowed;
    }

    .menu-node {
      position: relative;
    }

    .menu-caret {
      margin-left: 4px;
      color: var(--muted);
      font-size: 0.82rem;
    }

    .menu-panel,
    .tool-menu-panel {
      position: absolute;
      top: calc(100% + 10px);
      right: 0;
      min-width: 280px;
      z-index: 30;
      display: none;
      padding: 14px;
      border-radius: 20px;
      border: 1px solid rgba(78, 83, 96, 0.72);
      background: rgba(9, 11, 16, 0.98);
      box-shadow: 0 28px 60px rgba(0, 0, 0, 0.4);
    }

    .menu-panel.is-open,
    .tool-menu-panel.is-open {
      display: grid;
      gap: 12px;
    }

    .menu-kicker {
      margin: 0;
      font-size: 0.76rem;
      letter-spacing: 0.08em;
      color: rgba(223, 213, 198, 0.52);
      font-weight: 700;
      text-transform: uppercase;
    }

    .menu-note {
      margin: 0;
      color: rgba(230, 220, 206, 0.74);
      font-size: 0.9rem;
      line-height: 1.45;
    }

    .menu-divider {
      height: 1px;
      background: rgba(255, 255, 255, 0.08);
    }

    .menu-panel button,
    .menu-panel .button-link,
    .tool-link {
      width: 100%;
      justify-content: flex-start;
      border-radius: 14px;
      background: transparent;
      border: 1px solid rgba(255, 255, 255, 0.08);
      color: #f7f1e8;
    }

    .menu-panel button:hover,
    .menu-panel .button-link:hover,
    .tool-link.is-current {
      background: rgba(255, 255, 255, 0.1);
      border-color: rgba(255, 255, 255, 0.2);
    }

    .layout {
      display: grid;
      grid-template-columns: minmax(260px, 320px) minmax(0, 1fr);
      gap: 20px;
      align-items: start;
    }

    .sidebar,
    .workspace {
      padding: 20px;
    }

    .card {
      border: 1px solid var(--line);
      border-radius: 20px;
      padding: 16px;
      background: rgba(255, 255, 255, 0.48);
      display: grid;
      gap: 12px;
    }

    .stack {
      display: grid;
      gap: 14px;
    }

    h1,
    h2,
    h3,
    p {
      margin: 0;
    }

    p {
      color: var(--muted);
      line-height: 1.45;
    }

    input[type="file"],
    select,
    textarea {
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 11px 12px;
      background: rgba(255, 255, 255, 0.88);
      font: inherit;
      color: var(--ink);
    }

    input[type="file"] {
      border-style: dashed;
    }

    .file-picker {
      display: grid;
      gap: 8px;
      cursor: pointer;
    }

    .file-input {
      position: absolute;
      width: 1px;
      height: 1px;
      padding: 0;
      margin: -1px;
      overflow: hidden;
      clip: rect(0, 0, 0, 0);
      white-space: nowrap;
      border: 0;
    }

    .file-shell {
      display: flex;
      align-items: center;
      gap: 10px;
      min-width: 0;
      padding: 8px;
      border-radius: 16px;
      border: 1px solid rgba(201, 180, 148, 0.9);
      background: rgba(255, 255, 255, 0.86);
      transition: border-color 140ms ease, box-shadow 140ms ease, transform 140ms ease;
    }

    .file-picker:hover .file-shell,
    .file-picker:focus-within .file-shell,
    .file-shell.has-file {
      border-color: rgba(29, 107, 87, 0.42);
      box-shadow: 0 10px 24px rgba(17, 75, 60, 0.08);
      transform: translateY(-1px);
    }

    .file-trigger {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      flex: 0 0 auto;
      min-height: 40px;
      padding: 0 14px;
      border-radius: 12px;
      background: linear-gradient(180deg, #fffaf3 0%, #f2e4cf 100%);
      border: 1px solid rgba(201, 180, 148, 0.96);
      color: #2f2418;
      font-size: 0.9rem;
      font-weight: 700;
      box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.7);
    }

    .file-name {
      min-width: 0;
      color: #7b7368;
      font-size: 0.92rem;
      line-height: 1.3;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .file-shell.has-file .file-name {
      color: #1f2a1f;
      font-weight: 600;
    }

    textarea {
      min-height: 118px;
      resize: vertical;
    }

    button:disabled {
      opacity: 0.58;
      cursor: wait;
    }

    .button-link.is-disabled {
      opacity: 0.52;
      pointer-events: none;
    }

    .summary {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
    }

    .pill {
      padding: 7px 10px;
      border-radius: 999px;
      background: var(--accent-soft);
      color: var(--accent);
      font-size: 0.88rem;
      font-weight: 700;
    }

    .status {
      padding: 12px 14px;
      border-radius: 14px;
      background: rgba(255, 255, 255, 0.6);
      border: 1px solid rgba(217, 203, 185, 0.84);
      color: var(--ink);
    }

    .items-toolbar {
      display: flex;
      flex-wrap: wrap;
      align-items: end;
      gap: 12px;
      margin: 12px 0 16px;
      padding: 12px;
      border: 1px solid rgba(217, 203, 185, 0.82);
      border-radius: 14px;
      background: rgba(255, 255, 255, 0.42);
    }

    .filter-field {
      display: grid;
      gap: 5px;
      min-width: 160px;
    }

    .filter-field label,
    .meta label {
      color: var(--muted);
      font-size: 0.78rem;
      font-weight: 800;
      text-transform: uppercase;
    }

    .filter-count {
      margin-left: auto;
      color: var(--muted);
      font-size: 0.9rem;
      padding: 8px 2px;
    }

    .find-replace-launch {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 42px;
      padding: 0 18px;
      border-radius: 12px;
      border: 1px solid rgba(184, 168, 145, 0.96);
      background: linear-gradient(180deg, rgba(255, 255, 255, 0.98) 0%, rgba(243, 234, 221, 0.98) 100%);
      color: #2f2418;
      font-weight: 700;
      box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.82);
      white-space: nowrap;
    }

    .find-replace-launch:hover:not(:disabled) {
      transform: translateY(-1px);
      box-shadow: 0 10px 22px rgba(58, 43, 24, 0.1);
    }

    .find-replace-dialog[hidden] {
      display: none !important;
    }

    .find-replace-dialog {
      position: fixed;
      inset: 0;
      z-index: 90;
      display: grid;
      place-items: center;
      padding: 24px;
    }

    .find-replace-backdrop {
      position: absolute;
      inset: 0;
      background: rgba(34, 27, 18, 0.2);
      backdrop-filter: blur(2px);
    }

    .find-replace-card {
      position: relative;
      width: min(900px, calc(100vw - 48px));
      border: 1px solid #cfd4dc;
      border-radius: 0;
      background: #f8f8f8;
      box-shadow: 0 28px 54px rgba(34, 27, 18, 0.22);
      overflow: hidden;
      color: #111827;
    }

    .find-replace-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      min-height: 48px;
      padding: 10px 14px;
      background: linear-gradient(180deg, #fcfcfd 0%, #f1f3f6 100%);
      border-bottom: 1px solid #d9dde3;
    }

    .find-replace-title {
      font-size: 0.98rem;
      font-weight: 400;
      color: #1f2937;
    }

    .find-replace-close {
      min-width: 34px;
      min-height: 34px;
      padding: 0;
      border-radius: 8px;
      background: transparent;
      border: 1px solid transparent;
      color: #374151;
      font-size: 1.35rem;
      line-height: 1;
    }

    .find-replace-close:hover:not(:disabled) {
      background: rgba(217, 48, 37, 0.1);
      border-color: rgba(217, 48, 37, 0.16);
      transform: none;
    }

    .find-replace-tabs {
      display: flex;
      align-items: end;
      gap: 0;
      padding: 12px 12px 0;
      background: #f8f8f8;
    }

    .find-replace-tab {
      min-height: 34px;
      min-width: 92px;
      margin-right: 4px;
      padding: 0 16px;
      border-radius: 0;
      border: 1px solid #d9dde3;
      border-bottom: 0;
      background: #eceff3;
      color: #1f2937;
      font-weight: 400;
    }

    .find-replace-tab.is-active {
      background: #ffffff;
      position: relative;
      top: 1px;
    }

    .find-replace-body {
      display: grid;
      gap: 28px;
      padding: 18px 22px 26px;
      background: #ffffff;
      border-top: 1px solid #d9dde3;
    }

    .find-replace-row {
      display: grid;
      grid-template-columns: 96px minmax(0, 1fr);
      align-items: center;
      gap: 12px;
    }

    .find-replace-row label {
      color: #111827;
      font-size: 0.96rem;
      font-weight: 400;
    }

    .find-replace-row input {
      width: 100%;
      min-height: 40px;
      border: 1px solid #c8cdd5;
      border-radius: 0;
      padding: 8px 10px;
      background: #ffffff;
      color: #111827;
      box-shadow: inset 0 1px 2px rgba(17, 24, 39, 0.06);
    }

    .find-replace-row input:focus {
      outline: 1px solid #8fa9d8;
      border-color: #8fa9d8;
    }

    .find-replace-actions {
      display: flex;
      justify-content: flex-end;
      gap: 14px;
      padding: 0 22px 22px;
      background: #ffffff;
    }

    .find-replace-actions button {
      width: auto;
      min-width: 124px;
      min-height: 36px;
      padding: 0 16px;
      border-radius: 4px;
      border: 1px solid #c9ced6;
      background: linear-gradient(180deg, #ffffff 0%, #f2f4f7 100%);
      color: #1f2937;
      box-shadow: none;
      font-weight: 400;
    }

    .find-replace-actions button:hover:not(:disabled) {
      background: linear-gradient(180deg, #ffffff 0%, #e9edf3 100%);
      transform: none;
    }

    .find-replace-primary {
      border-color: #9fb5d8;
      background: linear-gradient(180deg, #fdfefe 0%, #e8f0fe 100%);
    }

    .find-replace-cancel {
      min-width: 110px;
    }

    @media (max-width: 760px) {
      .find-replace-dialog {
        padding: 16px;
      }

      .find-replace-card {
        width: min(100vw - 32px, 900px);
      }

      .find-replace-body {
        gap: 16px;
        padding: 16px;
      }

      .find-replace-row {
        grid-template-columns: 1fr;
        gap: 8px;
      }

      .find-replace-actions {
        flex-wrap: wrap;
        justify-content: stretch;
        padding: 0 16px 16px;
      }

      .find-replace-actions button {
        flex: 1 1 160px;
      }
    }

    .find-replace-launch {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 42px;
      padding: 0 18px;
      border-radius: 12px;
      border: 1px solid rgba(184, 168, 145, 0.96);
      background: linear-gradient(180deg, rgba(255, 255, 255, 0.98) 0%, rgba(243, 234, 221, 0.98) 100%);
      color: #2f2418;
      font-weight: 700;
      box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.82);
      white-space: nowrap;
    }

    .find-replace-launch:hover:not(:disabled) {
      transform: translateY(-1px);
      box-shadow: 0 10px 22px rgba(58, 43, 24, 0.1);
    }

    .find-replace-dialog[hidden] {
      display: none !important;
    }

    .find-replace-dialog {
      position: fixed;
      inset: 0;
      z-index: 90;
      display: grid;
      place-items: center;
      padding: 24px;
    }

    .find-replace-backdrop {
      position: absolute;
      inset: 0;
      background: rgba(34, 27, 18, 0.2);
      backdrop-filter: blur(2px);
    }

    .find-replace-card {
      position: relative;
      width: min(900px, calc(100vw - 48px));
      border: 1px solid #cfd4dc;
      border-radius: 0;
      background: #f8f8f8;
      box-shadow: 0 28px 54px rgba(34, 27, 18, 0.22);
      overflow: hidden;
      color: #111827;
    }

    .find-replace-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      min-height: 48px;
      padding: 10px 14px;
      background: linear-gradient(180deg, #fcfcfd 0%, #f1f3f6 100%);
      border-bottom: 1px solid #d9dde3;
    }

    .find-replace-title {
      font-size: 0.98rem;
      font-weight: 400;
      color: #1f2937;
    }

    .find-replace-close {
      min-width: 34px;
      min-height: 34px;
      padding: 0;
      border-radius: 8px;
      background: transparent;
      border: 1px solid transparent;
      color: #374151;
      font-size: 1.35rem;
      line-height: 1;
    }

    .find-replace-close:hover:not(:disabled) {
      background: rgba(217, 48, 37, 0.1);
      border-color: rgba(217, 48, 37, 0.16);
      transform: none;
    }

    .find-replace-tabs {
      display: flex;
      align-items: end;
      gap: 0;
      padding: 12px 12px 0;
      background: #f8f8f8;
    }

    .find-replace-tab {
      min-height: 34px;
      min-width: 92px;
      margin-right: 4px;
      padding: 0 16px;
      border-radius: 0;
      border: 1px solid #d9dde3;
      border-bottom: 0;
      background: #eceff3;
      color: #1f2937;
      font-weight: 400;
    }

    .find-replace-tab.is-active {
      background: #ffffff;
      position: relative;
      top: 1px;
    }

    .find-replace-body {
      display: grid;
      gap: 28px;
      padding: 18px 22px 26px;
      background: #ffffff;
      border-top: 1px solid #d9dde3;
    }

    .find-replace-row {
      display: grid;
      grid-template-columns: 96px minmax(0, 1fr);
      align-items: center;
      gap: 12px;
    }

    .find-replace-row label {
      color: #111827;
      font-size: 0.96rem;
      font-weight: 400;
    }

    .find-replace-row input {
      width: 100%;
      min-height: 40px;
      border: 1px solid #c8cdd5;
      border-radius: 0;
      padding: 8px 10px;
      background: #ffffff;
      color: #111827;
      box-shadow: inset 0 1px 2px rgba(17, 24, 39, 0.06);
    }

    .find-replace-row input:focus {
      outline: 1px solid #8fa9d8;
      border-color: #8fa9d8;
    }

    .find-replace-actions {
      display: flex;
      justify-content: flex-end;
      gap: 14px;
      padding: 0 22px 22px;
      background: #ffffff;
    }

    .find-replace-actions button {
      width: auto;
      min-width: 124px;
      min-height: 36px;
      padding: 0 16px;
      border-radius: 4px;
      border: 1px solid #c9ced6;
      background: linear-gradient(180deg, #ffffff 0%, #f2f4f7 100%);
      color: #1f2937;
      box-shadow: none;
      font-weight: 400;
    }

    .find-replace-actions button:hover:not(:disabled) {
      background: linear-gradient(180deg, #ffffff 0%, #e9edf3 100%);
      transform: none;
    }

    .find-replace-primary {
      border-color: #9fb5d8;
      background: linear-gradient(180deg, #fdfefe 0%, #e8f0fe 100%);
    }

    .find-replace-cancel {
      min-width: 110px;
    }

    @media (max-width: 760px) {
      .find-replace-dialog {
        padding: 16px;
      }

      .find-replace-card {
        width: min(100vw - 32px, 900px);
      }

      .find-replace-body {
        gap: 16px;
        padding: 16px;
      }

      .find-replace-row {
        grid-template-columns: 1fr;
        gap: 8px;
      }

      .find-replace-actions {
        flex-wrap: wrap;
        justify-content: stretch;
        padding: 0 16px 16px;
      }

      .find-replace-actions button {
        flex: 1 1 160px;
      }
    }

    .items {
      display: grid;
      gap: 16px;
    }

    .item {
      display: grid;
      grid-template-columns: 300px minmax(0, 1fr);
      gap: 16px;
      border: 1px solid var(--line);
      border-radius: 18px;
      padding: 16px;
      background: rgba(255, 255, 255, 0.55);
    }

    .preview {
      min-height: 190px;
      border-radius: 14px;
      overflow: auto;
      border: 1px solid rgba(217, 203, 185, 0.9);
      background: #fff;
      display: flex;
      align-items: center;
      justify-content: flex-start;
      padding: 10px;
    }

    .preview img {
      width: auto;
      height: auto;
      max-width: 100%;
      max-height: 210px;
      display: block;
      background: white;
    }

    .item.formula-item .preview img {
      max-width: none;
    }

    .meta {
      display: grid;
      gap: 8px;
    }

    .meta strong {
      font-size: 1.05rem;
    }

    .chips {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin: 4px 0;
    }

    .chip {
      padding: 6px 10px;
      border-radius: 999px;
      font-size: 0.82rem;
      font-weight: 800;
    }

    .chip.good {
      background: var(--accent-soft);
      color: var(--accent);
    }

    .chip.warn {
      background: var(--warn-soft);
      color: var(--warn);
    }

    .empty {
      padding: 32px 18px;
      border: 1px dashed var(--line);
      border-radius: 18px;
      color: var(--muted);
      text-align: center;
      background: rgba(255, 255, 255, 0.45);
    }

    @media (max-width: 980px) {
      .appbar {
        flex-direction: column;
        align-items: stretch;
      }

      .appbar-right {
        justify-content: flex-start;
        flex-wrap: wrap;
      }

      .menu-panel,
      .tool-menu-panel {
        position: static;
        margin-top: 10px;
      }

      .layout,
      .item {
        grid-template-columns: 1fr;
      }

      .filter-count {
        width: 100%;
        margin-left: 0;
      }
    }
  </style>
</head>
<body>
  <div class="shell">
    <header class="appbar panel">
      <div class="brandmark">HBS Alto</div>
      <nav class="appbar-right header-actions" aria-label="Primary">
        <div class="menu-node">
          <button class="header-action menu-trigger" type="button" data-menu-trigger="pdfUploadMenu" aria-expanded="false" aria-controls="pdfUploadMenu">
            <strong>Upload</strong>
            <span class="menu-caret">&#9662;</span>
          </button>
          <div id="pdfUploadMenu" class="menu-panel" data-menu-panel>
            <p class="menu-kicker">Upload</p>
            <p class="menu-note">Choose a tagged PDF and start a PDF ALT editing session.</p>
            <div class="menu-divider"></div>
            <button id="quickPdfChooseBtn" type="button">Choose PDF</button>
            <button id="quickPdfProcessBtn" type="button">Process PDF</button>
            <button id="quickPdfImportExcelBtn" type="button" disabled>Upload ALT Excel</button>
          </div>
        </div>

        <div class="menu-node">
          <button class="header-action menu-trigger" type="button" data-menu-trigger="pdfDownloadMenu" aria-expanded="false" aria-controls="pdfDownloadMenu">
            <strong>Download</strong>
            <span class="menu-caret">&#9662;</span>
          </button>
          <div id="pdfDownloadMenu" class="menu-panel" data-menu-panel>
            <p class="menu-kicker">Download</p>
            <p class="menu-note">Save the edited PDF with ALT text written into Figure and Formula tag properties.</p>
            <div class="menu-divider"></div>
            <button id="quickPdfDownloadExcelBtn" type="button" disabled>Download Excel</button>
            <button id="quickPdfDownloadBtn" type="button" disabled>Download Updated PDF</button>
          </div>
        </div>

        <div class="menu-node">
          <button class="header-action menu-trigger" type="button" data-menu-trigger="pdfImagesMenu" aria-expanded="false" aria-controls="pdfImagesMenu">
            <strong>Images</strong>
            <span class="menu-caret">&#9662;</span>
          </button>
          <div id="pdfImagesMenu" class="menu-panel" data-menu-panel>
            <p class="menu-kicker">Images</p>
            <p class="menu-note">Generate missing ALT from the collected Figure and Formula previews.</p>
            <div class="menu-divider"></div>
            <button id="quickPdfGenerateAltBtn" type="button" disabled>Generate ALT</button>
            <button id="quickPdfClearBtn" type="button" disabled>Clear ALT</button>
          </div>
        </div>

        <div class="menu-node">
          <button class="header-action menu-trigger" type="button" data-menu-trigger="pdfToolsMenu" aria-expanded="false" aria-controls="pdfToolsMenu" aria-label="Tools">
            <strong>&#9776;</strong>
          </button>
          <div id="pdfToolsMenu" class="menu-panel" data-menu-panel>
            <p class="menu-kicker">Tools</p>
            <p class="menu-note">Switch between HBS Alto document tools.</p>
            <div class="menu-divider"></div>
            <a class="button-link" href="/">DOCX ALT Editor</a>
            <a class="button-link" href="/list-correction">List Correction</a>
            <a class="button-link" href="/color-correction">Color Correction</a>
            <a class="button-link" href="/excel-merger">Excel Merger</a>
          </div>
        </div>
      </nav>
    </header>

    <div class="is-hidden" aria-hidden="true">
      <button id="generatePdfAltBtn" type="button" disabled>Generate Missing ALT</button>
      <button id="downloadPdfBtn" type="button" disabled>Download Updated PDF</button>
      <button id="clearPdfAltBtn" type="button" disabled>Clear ALT</button>
      <input id="pdfExcelFile" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" />
    </div>

    <section class="layout">
      <aside class="panel sidebar">
        <div class="stack">
          <section class="card">
            <h3>Process PDF</h3>
            <p>Upload a tagged PDF to collect Figure and Formula structure tags.</p>
            <label class="file-picker" for="pdfSourceFile">
              <input id="pdfSourceFile" class="file-input" type="file" accept=".pdf,application/pdf" />
              <span id="pdfFileShell" class="file-shell">
                <span class="file-trigger">Choose PDF</span>
                <span id="pdfFileName" class="file-name">No file selected</span>
              </span>
            </label>
            <button id="processPdfBtn" type="button">Process PDF ALT Inventory</button>
          </section>

          <section class="card">
            <h3>ALT Provider</h3>
            <p>Choose which model generates missing ALT text for this PDF session.</p>
            <select id="pdfProviderSelect">
              <option value="claude">Claude</option>
              <option value="copilot">Copilot</option>
              <option value="gemini">Gemini</option>
              <option value="groq">Groq</option>
              <option value="openrouter">OpenRouter</option>
              <option value="claude_fallback_groq">Claude then Groq fallback</option>
            </select>
          </section>

          <section class="card">
            <h3>Session Summary</h3>
            <div class="summary" id="pdfSummary"><span class="pill">No session yet</span></div>
            <h3>Session Status</h3>
            <p class="status" id="pdfStatus">Ready.</p>
          </section>
        </div>
      </aside>

      <main class="panel workspace">
        <h2>PDF Tagged Items</h2>
        <div class="items-toolbar" aria-label="PDF item filters">
          <div class="filter-field">
            <label for="pdfTypeFilter">Type</label>
            <select id="pdfTypeFilter">
              <option value="all">All types</option>
              <option value="figure">Figures</option>
              <option value="formula">Formulas</option>
            </select>
          </div>
          <div class="filter-field">
            <label for="pdfSplitFilter">Split</label>
            <select id="pdfSplitFilter">
              <option value="all">All items</option>
              <option value="with-alt">With ALT</option>
              <option value="without-alt">Without ALT</option>
            </select>
          </div>
          <div class="filter-field">
            <label for="pdfFindReplaceOpenBtn">Find &amp; Replace</label>
            <button id="pdfFindReplaceOpenBtn" class="find-replace-launch" type="button">Find &amp; Replace</button>
          </div>
          <div id="pdfFilterCount" class="filter-count">0 shown</div>
        </div>
        <div id="pdfFindReplaceDialog" class="find-replace-dialog" hidden>
          <div class="find-replace-backdrop" data-find-replace-close="pdf"></div>
          <div class="find-replace-card" role="dialog" aria-modal="true" aria-labelledby="pdfFindReplaceTitle">
            <div class="find-replace-head">
              <strong id="pdfFindReplaceTitle" class="find-replace-title">Find and Replace</strong>
              <button id="pdfFindReplaceCloseBtn" class="find-replace-close" type="button" aria-label="Close dialog">&times;</button>
            </div>
            <div class="find-replace-tabs" aria-hidden="true">
              <button class="find-replace-tab is-active" type="button" tabindex="-1">Find</button>
              <button class="find-replace-tab" type="button" tabindex="-1">Replace</button>
            </div>
            <div class="find-replace-body">
              <div class="find-replace-row">
                <label for="pdfFindText">Find what:</label>
                <input id="pdfFindText" type="text" placeholder="Search exact text" />
              </div>
              <div class="find-replace-row">
                <label for="pdfReplaceText">Replace with:</label>
                <input id="pdfReplaceText" type="text" placeholder="Optional replacement" />
              </div>
            </div>
            <div class="find-replace-actions">
              <button id="pdfFindMatchesBtn" type="button">Find Matches</button>
              <button id="pdfReplaceBtn" class="find-replace-primary" type="button">Replace All</button>
              <button class="find-replace-cancel" type="button" data-find-replace-close="pdf">Cancel</button>
            </div>
          </div>
        </div>
        <div id="pdfResults" class="items">
          <div class="empty">Process a tagged PDF to inspect Figure and Formula ALT items.</div>
        </div>
      </main>
    </section>
  </div>

  <script>
    const pdfSourceFile = document.getElementById("pdfSourceFile");
    const processPdfBtn = document.getElementById("processPdfBtn");
    const generatePdfAltBtn = document.getElementById("generatePdfAltBtn");
    const downloadPdfBtn = document.getElementById("downloadPdfBtn");
    const clearPdfAltBtn = document.getElementById("clearPdfAltBtn");
    const pdfExcelFile = document.getElementById("pdfExcelFile");
    const pdfProviderSelect = document.getElementById("pdfProviderSelect");
    const pdfSummary = document.getElementById("pdfSummary");
    const pdfStatus = document.getElementById("pdfStatus");
    const pdfResults = document.getElementById("pdfResults");
    const pdfTypeFilter = document.getElementById("pdfTypeFilter");
    const pdfSplitFilter = document.getElementById("pdfSplitFilter");
    const pdfFindReplaceOpenBtn = document.getElementById("pdfFindReplaceOpenBtn");
    const pdfFindReplaceDialog = document.getElementById("pdfFindReplaceDialog");
    const pdfFindReplaceCloseBtn = document.getElementById("pdfFindReplaceCloseBtn");
    const pdfFindReplaceDismissButtons = Array.from(pdfFindReplaceDialog.querySelectorAll("[data-find-replace-close]"));
    const pdfFindText = document.getElementById("pdfFindText");
    const pdfReplaceText = document.getElementById("pdfReplaceText");
    const pdfFindMatchesBtn = document.getElementById("pdfFindMatchesBtn");
    const pdfReplaceBtn = document.getElementById("pdfReplaceBtn");
    const pdfFilterCount = document.getElementById("pdfFilterCount");
    const pdfFileShell = document.getElementById("pdfFileShell");
    const pdfFileName = document.getElementById("pdfFileName");
    const quickPdfChooseBtn = document.getElementById("quickPdfChooseBtn");
    const quickPdfProcessBtn = document.getElementById("quickPdfProcessBtn");
    const quickPdfDownloadBtn = document.getElementById("quickPdfDownloadBtn");
    const quickPdfDownloadExcelBtn = document.getElementById("quickPdfDownloadExcelBtn");
    const quickPdfImportExcelBtn = document.getElementById("quickPdfImportExcelBtn");
    const quickPdfGenerateAltBtn = document.getElementById("quickPdfGenerateAltBtn");
    const quickPdfClearBtn = document.getElementById("quickPdfClearBtn");
    const menuTriggers = Array.from(document.querySelectorAll("[data-menu-trigger]"));
    const menuPanels = Array.from(document.querySelectorAll("[data-menu-panel]"));

    let pdfSession = null;
    let pdfRows = [];
    let pdfTextSearch = "";
    const pdfSaveTimers = new Map();
    const pdfSavePromises = new Map();

    const escapeHtml = (value) =>
      String(value || "").replace(/[&<>"']/g, (char) => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;",
      }[char]));

    function providerLabel() {
      const provider = String(pdfProviderSelect?.value || "claude").toLowerCase();
      if (provider === "copilot") return "Copilot";
      if (provider === "gemini") return "Gemini";
      if (provider === "groq") return "Groq";
      if (provider === "openrouter") return "OpenRouter";
      if (provider === "claude_fallback_groq") return "Claude with Groq fallback";
      return "Claude";
    }

    function closeAllMenus() {
      menuTriggers.forEach((trigger) => trigger.classList.remove("is-open"));
      menuTriggers.forEach((trigger) => trigger.setAttribute("aria-expanded", "false"));
      menuPanels.forEach((panel) => panel.classList.remove("is-open"));
    }

    function toggleMenu(menuId) {
      const panel = document.getElementById(menuId);
      const trigger = menuTriggers.find((item) => item.getAttribute("data-menu-trigger") === menuId);
      if (!panel || !trigger) {
        return;
      }
      const isOpen = panel.classList.contains("is-open");
      closeAllMenus();
      if (!isOpen) {
        panel.classList.add("is-open");
        trigger.classList.add("is-open");
        trigger.setAttribute("aria-expanded", "true");
      }
    }

    function refreshPdfQuickActions() {
      const hasSession = Boolean(pdfSession);
      quickPdfDownloadBtn.disabled = !hasSession;
      quickPdfDownloadExcelBtn.disabled = !hasSession;
      quickPdfImportExcelBtn.disabled = !hasSession;
      quickPdfGenerateAltBtn.disabled = !hasSession;
      quickPdfClearBtn.disabled = !hasSession;
      clearPdfAltBtn.disabled = !hasSession;
      pdfFindReplaceOpenBtn.disabled = !hasSession;
      pdfFindMatchesBtn.disabled = !hasSession;
      pdfReplaceBtn.disabled = !hasSession;
    }

    function syncPdfFileState() {
      const [file] = pdfSourceFile.files || [];
      pdfFileName.textContent = file ? file.name : "No file selected";
      pdfFileShell.classList.toggle("has-file", Boolean(file));
    }

    function effectiveAltText(row) {
      return String(row?.alt_text || row?.generated_alt_text || row?.existing_alt_text || "").trim();
    }

    function renderSummary(data) {
      if (!data) {
        pdfSummary.innerHTML = '<span class="pill">No session yet</span>';
        return;
      }
      const pills = [
        `${data.total_items || 0} item(s)`,
        `${data.images || 0} figure(s)`,
        `${data.equations || 0} formula(s)`,
        `${data.with_alt_text || 0} original ALT`,
        `${data.without_alt_text || 0} missing`,
      ];
      pdfSummary.innerHTML = pills.map((label) => `<span class="pill">${escapeHtml(label)}</span>`).join("");
    }

    function rowMatchesFilters(row) {
      const typeValue = String(pdfTypeFilter?.value || "all").toLowerCase();
      const splitValue = String(pdfSplitFilter?.value || "all").toLowerCase();
      const standardRole = String(row?.standard_role || "").toLowerCase();
      if (typeValue !== "all" && standardRole !== typeValue) {
        return false;
      }
      const hasAlt = Boolean(effectiveAltText(row));
      if (splitValue === "with-alt" && !hasAlt) return false;
      if (splitValue === "without-alt" && hasAlt) return false;
      if (pdfTextSearch) {
        const currentAlt = String(row?.alt_text || "").toLowerCase();
        if (!currentAlt.includes(pdfTextSearch)) {
          return false;
        }
      }
      return true;
    }

    function visiblePdfRows() {
      return pdfRows.filter(rowMatchesFilters);
    }

    function updateFilterCount(visibleCount) {
      const total = pdfRows.length;
      pdfFilterCount.textContent = total ? `${visibleCount} of ${total} shown` : "0 shown";
    }

    function filenameFromDisposition(headerValue, fallbackName) {
      const value = String(headerValue || "");
      const encodedMatch = value.match(/filename\\*=UTF-8''([^;]+)/i);
      if (encodedMatch) {
        try {
          return decodeURIComponent(encodedMatch[1]);
        } catch (error) {
          return fallbackName;
        }
      }
      const basicMatch = value.match(/filename="?([^";]+)"?/i);
      return basicMatch ? basicMatch[1] : fallbackName;
    }

    async function downloadSessionFile(url, fallbackName) {
      const response = await fetch(url);
      if (!response.ok) {
        let message = "Download failed.";
        try {
          const payload = await response.json();
          message = payload.detail || message;
        } catch (error) {
          const text = await response.text();
          message = text || message;
        }
        throw new Error(message);
      }
      const blob = await response.blob();
      const objectUrl = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = objectUrl;
      link.download = filenameFromDisposition(response.headers.get("Content-Disposition"), fallbackName);
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
    }

    function previewSrc(itemId) {
      return `/api/pdf-alt/session/${pdfSession}/item/${itemId}/preview.png`;
    }

    function renderResults() {
      if (!pdfRows.length) {
        updateFilterCount(0);
        pdfResults.innerHTML = '<div class="empty">No Figure or Formula tags were found in this PDF structure tree.</div>';
        return;
      }

      const visibleRows = visiblePdfRows();
      updateFilterCount(visibleRows.length);
      if (!visibleRows.length) {
        pdfResults.innerHTML = '<div class="empty">No PDF items match the selected filters.</div>';
        return;
      }

      pdfResults.innerHTML = visibleRows.map((row, index) => `
        <article class="item ${String(row.standard_role || "").toLowerCase() === "formula" ? "formula-item" : "figure-item"}">
          <div class="preview">
            <img src="${previewSrc(row.id)}" alt="${escapeHtml(row.type)} preview ${index + 1}" loading="lazy" />
          </div>
          <div class="meta">
            <strong>#${index + 1} ${escapeHtml(row.type)}</strong>
            <p>${escapeHtml(row.source_part || "PDF structure tree")} &middot; Page ${row.page || "-"}</p>
            <p>${row.existing_alt_text ? `Original ALT: ${escapeHtml(row.existing_alt_text)}` : "Original ALT missing for this tag."}</p>
            <div class="chips">
              <span class="chip ${row.has_alt_text ? "good" : "warn"}">${row.has_alt_text ? "Original ALT" : "Original ALT missing"}</span>
              ${row.generated_alt_text ? '<span class="chip good">Generated ALT</span>' : ''}
              ${row.alt_source === "manual" ? '<span class="chip good">Dashboard edit</span>' : ''}
            </div>
            <label for="pdf-alt-editor-${row.id}">Current ALT text</label>
            <textarea id="pdf-alt-editor-${row.id}" data-pdf-alt-editor="${row.id}" placeholder="Write ALT text for this PDF tag.">${escapeHtml(row.alt_text || "")}</textarea>
            <p data-pdf-save-state="${row.id}" style="margin-top: 4px;">${row.alt_source === "manual" ? "Saved" : ""}</p>
          </div>
        </article>
      `).join("");

      document.querySelectorAll("[data-pdf-alt-editor]").forEach((field) => {
        field.addEventListener("input", () => {
          const itemId = Number(field.getAttribute("data-pdf-alt-editor"));
          scheduleSave(itemId, field.value);
        });
        field.addEventListener("blur", async () => {
          const itemId = Number(field.getAttribute("data-pdf-alt-editor"));
          await flushPendingSaves();
          setSaveState(itemId, "Saved");
        });
      });
    }

    function setSaveState(itemId, message) {
      const el = document.querySelector(`[data-pdf-save-state="${itemId}"]`);
      if (el) el.textContent = message;
    }

    async function saveRow(itemId, altText) {
      if (!pdfSession) return;
      const promise = fetch(`/api/pdf-alt/session/${pdfSession}/item/${itemId}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ alt_text: altText }),
      })
        .then(async (response) => {
          const payload = await response.json();
          if (!response.ok) throw new Error(payload.detail || "PDF ALT save failed.");
          pdfRows = pdfRows.map((row) => row.id === itemId ? payload.row : row);
          renderSummary(payload.summary || null);
          setSaveState(itemId, "Saved");
        })
        .catch((error) => {
          setSaveState(itemId, error.message);
          throw error;
        })
        .finally(() => {
          pdfSavePromises.delete(itemId);
        });
      pdfSavePromises.set(itemId, promise);
      return promise;
    }

    function scheduleSave(itemId, altText) {
      const row = pdfRows.find((entry) => entry.id === itemId);
      if (row) row.alt_text = altText;
      const timerId = pdfSaveTimers.get(itemId);
      if (timerId) clearTimeout(timerId);
      setSaveState(itemId, "Saving...");
      pdfSaveTimers.set(itemId, window.setTimeout(() => {
        pdfSaveTimers.delete(itemId);
        void saveRow(itemId, altText);
      }, 450));
    }

    async function flushPendingSaves() {
      Array.from(pdfSaveTimers.entries()).forEach(([itemId, timerId]) => {
        clearTimeout(timerId);
        pdfSaveTimers.delete(itemId);
        const row = pdfRows.find((entry) => entry.id === itemId);
        void saveRow(itemId, row?.alt_text || "");
      });
      const pending = Array.from(pdfSavePromises.values());
      if (pending.length) {
        const settled = await Promise.allSettled(pending);
        const rejected = settled.find((entry) => entry.status === "rejected");
        if (rejected && rejected.status === "rejected") throw rejected.reason;
      }
    }

    async function importPdfExcel(file) {
      if (!pdfSession || !file) {
        pdfStatus.textContent = pdfSession ? "Choose an ALT Excel workbook first." : "Process a PDF before importing ALT Excel.";
        return;
      }
      quickPdfImportExcelBtn.disabled = true;
      closeAllMenus();
      pdfStatus.textContent = "Importing ALT text from Excel...";
      try {
        await flushPendingSaves();
        const formData = new FormData();
        formData.append("workbook", file);
        const response = await fetch(`/api/pdf-alt/session/${pdfSession}/import-excel`, {
          method: "POST",
          body: formData,
        });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.detail || "PDF ALT Excel import failed.");
        pdfRows = payload.rows || [];
        renderSummary(payload.summary || null);
        renderResults();
        pdfStatus.textContent = payload.message || "Imported PDF ALT text from Excel.";
      } catch (error) {
        pdfStatus.textContent = error.message;
      } finally {
        pdfExcelFile.value = "";
        refreshPdfQuickActions();
      }
    }

    async function clearPdfAltSession() {
      if (!pdfSession) {
        pdfStatus.textContent = "Process a PDF before clearing ALT text.";
        return;
      }
      clearPdfAltBtn.disabled = true;
      quickPdfClearBtn.disabled = true;
      closeAllMenus();
      pdfStatus.textContent = "Clearing PDF ALT text in this session...";
      try {
        await flushPendingSaves();
        const response = await fetch(`/api/pdf-alt/session/${pdfSession}/clear-alt`, {
          method: "POST",
        });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.detail || "Clear PDF ALT failed.");
        pdfRows = payload.rows || [];
        renderSummary(payload.summary || null);
        renderResults();
        pdfStatus.textContent = payload.message || "Cleared PDF ALT text in this session.";
      } catch (error) {
        pdfStatus.textContent = error.message;
      } finally {
        refreshPdfQuickActions();
      }
    }

    async function replacePdfAltText() {
      if (!pdfSession) {
        pdfStatus.textContent = "Process a PDF before replacing ALT text.";
        return;
      }
      const searchText = String(pdfFindText?.value || "");
      const replaceText = String(pdfReplaceText?.value || "");
      if (!searchText) {
        pdfStatus.textContent = "Enter the exact text you want to find first.";
        return;
      }
      const targetIds = visiblePdfRows()
        .filter((row) => String(row?.alt_text || "").includes(searchText))
        .map((row) => Number(row.id))
        .filter((itemId) => Number.isInteger(itemId));
      if (!targetIds.length) {
        pdfStatus.textContent = "No visible PDF ALT text matches that search string.";
        return;
      }
      pdfReplaceBtn.disabled = true;
      closeAllMenus();
      pdfStatus.textContent = "Replacing ALT text in the visible PDF items...";
      try {
        await flushPendingSaves();
        const response = await fetch(`/api/pdf-alt/session/${pdfSession}/replace-alt-text`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            search_text: searchText,
            replace_text: replaceText,
            item_ids: targetIds,
          }),
        });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.detail || "Replace ALT text failed.");
        pdfRows = payload.rows || [];
        renderSummary(payload.summary || null);
        renderResults();
        pdfStatus.textContent = payload.message || "Updated PDF ALT text.";
      } catch (error) {
        pdfStatus.textContent = error.message;
      } finally {
        refreshPdfQuickActions();
      }
    }

    function openPdfFindReplaceDialog() {
      pdfFindReplaceDialog.hidden = false;
      window.setTimeout(() => pdfFindText.focus(), 0);
    }

    function closePdfFindReplaceDialog() {
      pdfFindReplaceDialog.hidden = true;
    }

    function applyPdfFindMatches() {
      const searchText = String(pdfFindText?.value || "").trim().toLowerCase();
      pdfTextSearch = searchText;
      renderResults();
      if (!searchText) {
        pdfStatus.textContent = "Showing all PDF items.";
      } else {
        pdfStatus.textContent = `Showing ${visiblePdfRows().length} PDF item(s) matching "${pdfFindText.value}".`;
      }
      closePdfFindReplaceDialog();
    }

    processPdfBtn.addEventListener("click", async () => {
      const [file] = pdfSourceFile.files || [];
      if (!file) {
        pdfStatus.textContent = "Choose a tagged PDF first.";
        return;
      }

      processPdfBtn.disabled = true;
      quickPdfChooseBtn.disabled = true;
      quickPdfProcessBtn.disabled = true;
      generatePdfAltBtn.disabled = true;
      downloadPdfBtn.disabled = true;
      clearPdfAltBtn.disabled = true;
      pdfFindMatchesBtn.disabled = true;
      pdfFindReplaceOpenBtn.disabled = true;
      pdfReplaceBtn.disabled = true;
      pdfSession = null;
      pdfRows = [];
      pdfTextSearch = "";
      closeAllMenus();
      refreshPdfQuickActions();
      renderSummary(null);
      pdfResults.innerHTML = '<div class="empty">Collecting tagged PDF Figure and Formula items...</div>';
      pdfStatus.textContent = "Reading PDF structure tree...";

      try {
        const formData = new FormData();
        formData.append("file", file);
        const response = await fetch("/api/pdf-alt/analyze", { method: "POST", body: formData });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.detail || "PDF ALT inventory failed.");

        pdfSession = payload.session_id;
        pdfRows = payload.rows || [];
        renderSummary(payload.summary || null);
        renderResults();
        generatePdfAltBtn.disabled = !pdfSession;
        downloadPdfBtn.disabled = !pdfSession;
        refreshPdfQuickActions();
        pdfStatus.textContent = payload.message || `Collected ${pdfRows.length} PDF ALT item(s).`;
      } catch (error) {
        pdfStatus.textContent = error.message;
        pdfResults.innerHTML = `<div class="empty">${escapeHtml(error.message)}</div>`;
      } finally {
        processPdfBtn.disabled = false;
        quickPdfChooseBtn.disabled = false;
        quickPdfProcessBtn.disabled = false;
        refreshPdfQuickActions();
      }
    });

    generatePdfAltBtn.addEventListener("click", async () => {
      if (!pdfSession) {
        pdfStatus.textContent = "Process a PDF before generating ALT text.";
        return;
      }
      generatePdfAltBtn.disabled = true;
      quickPdfGenerateAltBtn.disabled = true;
      closeAllMenus();
      pdfStatus.textContent = `Generating missing ALT text with ${providerLabel()}...`;
      try {
        await flushPendingSaves();
        const response = await fetch(`/api/pdf-alt/session/${pdfSession}/generate-alt`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ provider: String(pdfProviderSelect.value || "claude") }),
        });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.detail || "PDF ALT generation failed.");
        pdfRows = payload.rows || [];
        renderSummary(payload.summary || null);
        renderResults();
        pdfStatus.textContent = payload.message || "Generated PDF ALT text.";
      } catch (error) {
        pdfStatus.textContent = error.message;
      } finally {
        generatePdfAltBtn.disabled = false;
        refreshPdfQuickActions();
      }
    });

    downloadPdfBtn.addEventListener("click", async () => {
      if (!pdfSession) return;
      downloadPdfBtn.disabled = true;
      quickPdfDownloadBtn.disabled = true;
      closeAllMenus();
      pdfStatus.textContent = "Saving PDF tag ALT text and preparing updated PDF...";
      try {
        await flushPendingSaves();
        await downloadSessionFile(`/api/pdf-alt/session/${pdfSession}/updated.pdf`, "updated_alt.pdf");
        pdfStatus.textContent = "Updated PDF download is ready.";
      } catch (error) {
        pdfStatus.textContent = error.message;
      } finally {
        downloadPdfBtn.disabled = false;
        refreshPdfQuickActions();
      }
    });

    quickPdfDownloadExcelBtn.addEventListener("click", async () => {
      if (!pdfSession) return;
      quickPdfDownloadExcelBtn.disabled = true;
      closeAllMenus();
      pdfStatus.textContent = "Saving changes and preparing PDF ALT Excel...";
      try {
        await flushPendingSaves();
        window.location.href = `/api/pdf-alt/session/${pdfSession}/download.xlsx`;
        pdfStatus.textContent = "PDF ALT Excel download is ready.";
      } catch (error) {
        pdfStatus.textContent = error.message;
      } finally {
        refreshPdfQuickActions();
      }
    });

    pdfTypeFilter.addEventListener("change", renderResults);
    pdfSplitFilter.addEventListener("change", renderResults);
    pdfSourceFile.addEventListener("change", syncPdfFileState);

    quickPdfChooseBtn.addEventListener("click", () => {
      closeAllMenus();
      pdfSourceFile.click();
    });
    quickPdfProcessBtn.addEventListener("click", () => {
      closeAllMenus();
      processPdfBtn.click();
    });
    quickPdfImportExcelBtn.addEventListener("click", () => {
      closeAllMenus();
      pdfExcelFile.click();
    });
    pdfExcelFile.addEventListener("change", () => {
      const [file] = pdfExcelFile.files || [];
      void importPdfExcel(file);
    });
    quickPdfGenerateAltBtn.addEventListener("click", () => {
      closeAllMenus();
      generatePdfAltBtn.click();
    });
    quickPdfClearBtn.addEventListener("click", () => {
      closeAllMenus();
      clearPdfAltBtn.click();
    });
    pdfFindReplaceOpenBtn.addEventListener("click", () => {
      openPdfFindReplaceDialog();
    });
    pdfFindReplaceCloseBtn.addEventListener("click", () => {
      closePdfFindReplaceDialog();
    });
    pdfFindReplaceDismissButtons.forEach((button) => {
      button.addEventListener("click", () => {
        closePdfFindReplaceDialog();
      });
    });
    pdfFindMatchesBtn.addEventListener("click", () => {
      applyPdfFindMatches();
    });
    quickPdfDownloadBtn.addEventListener("click", () => {
      closeAllMenus();
      downloadPdfBtn.click();
    });
    pdfReplaceBtn.addEventListener("click", async () => {
      await replacePdfAltText();
    });
    clearPdfAltBtn.addEventListener("click", async () => {
      await clearPdfAltSession();
    });
    menuTriggers.forEach((trigger) => {
      trigger.addEventListener("click", (event) => {
        event.stopPropagation();
        toggleMenu(trigger.getAttribute("data-menu-trigger"));
      });
    });

    document.addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof Element)) {
        closeAllMenus();
        return;
      }
      if (!target.closest("[data-menu-trigger]") && !target.closest("[data-menu-panel]")) {
        closeAllMenus();
      }
    });
    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape" && !pdfFindReplaceDialog.hidden) {
        closePdfFindReplaceDialog();
      }
    });
    syncPdfFileState();
    refreshPdfQuickActions();
  </script>
</body>
</html>
        """
    )


@app.get("/", response_class=HTMLResponse)
async def index():
    return browser_ui_response(
        """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>HBS Alto</title>
  <style>
    :root {
      --bg: #f4efe7;
      --panel: rgba(255, 251, 245, 0.94);
      --line: #d9cbb9;
      --ink: #1f2a1f;
      --muted: #6b7166;
      --accent: #1d6b57;
      --accent-soft: #d8efe7;
      --warn: #9a3412;
      --warn-soft: #ffedd5;
      --shadow: 0 20px 45px rgba(58, 43, 24, 0.08);
      font-family: "Segoe UI", "Aptos", sans-serif;
    }

    * {
      box-sizing: border-box;
    }

    .is-hidden {
      display: none !important;
    }

    body {
      margin: 0;
      background:
        radial-gradient(circle at top right, rgba(29, 107, 87, 0.12), transparent 28%),
        radial-gradient(circle at top left, rgba(196, 167, 120, 0.1), transparent 32%),
        linear-gradient(180deg, #f8f3ea 0%, #f3ebdd 100%);
      background-attachment: fixed;
      background-repeat: no-repeat;
      background-size: cover;
      color: var(--ink);
    }

    .shell {
      max-width: 1280px;
      margin: 0 auto;
      padding: 28px 20px 40px;
    }

    .hero {
      background: var(--panel);
      border: 1px solid rgba(217, 203, 185, 0.8);
      border-radius: 24px;
      padding: 24px;
      box-shadow: var(--shadow);
      margin-bottom: 20px;
    }

    .hero h1 {
      margin: 0 0 8px;
      font-size: clamp(2rem, 3vw, 3rem);
      line-height: 1;
    }

    .hero p {
      margin: 0;
      color: var(--muted);
      max-width: 760px;
    }

    .appbar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 20px;
      padding: 25px 34px;
      margin-bottom: 18px;
      background: linear-gradient(180deg, rgba(255, 250, 244, 0.96) 0%, rgba(249, 241, 229, 0.98) 100%);
      border: 1px solid rgba(201, 180, 148, 0.58);
      border-radius: 24px;
      box-shadow: 0 22px 54px rgba(77, 55, 24, 0.12);
      color: #35261a;
      overflow: visible;
    }

    .appbar-left,
    .appbar-right {
      display: flex;
      align-items: center;
      gap: 24px;
      min-width: 0;
    }

    .brandmark {
      font-size: 2.6rem;
      line-height: 1;
      font-weight: 800;
      letter-spacing: -0.06em;
      color: #2f2418;
      white-space: nowrap;
    }

    .header-action {
      width: auto;
      appearance: none;
      border: 0;
      background: transparent;
      color: rgba(61, 42, 24, 0.72);
      font: inherit;
      padding: 13px 18px;
      border-radius: 14px;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 8px;
      text-decoration: none;
      white-space: nowrap;
      transition: background 140ms ease, color 140ms ease;
    }

    .header-action:hover:not(:disabled),
    .header-action.is-open {
      background: rgba(29, 107, 87, 0.1);
      color: #1f2a1f;
    }

    .header-action strong {
      font-size: 1.08rem;
      font-weight: 700;
    }

    .header-action .menu-caret {
      color: rgba(61, 42, 24, 0.5);
      font-size: 0.98rem;
    }

    .panel-action {
      width: 100%;
      justify-content: flex-start;
      padding: 14px 16px;
      border-radius: 14px;
      background: rgba(255, 255, 255, 0.04);
      border: 1px solid rgba(255, 255, 255, 0.08);
      color: #f7f1e8;
    }

    .panel-action:hover {
      background: rgba(255, 255, 255, 0.08);
    }

    .layout {
      display: grid;
      grid-template-columns: 300px minmax(0, 1fr);
      gap: 20px;
      align-items: start;
    }

    .panel {
      background: var(--panel);
      border: 1px solid rgba(217, 203, 185, 0.8);
      border-radius: 24px;
      box-shadow: var(--shadow);
    }

    .sidebar,
    .workspace {
      padding: 20px;
    }

    .stack {
      display: grid;
      gap: 14px;
    }

    .card {
      border: 1px solid var(--line);
      border-radius: 18px;
      padding: 16px;
      background: rgba(255, 255, 255, 0.55);
    }

    .card h3,
    .section-title {
      margin: 0 0 8px;
      font-size: 1rem;
    }

    .card p,
    .note,
    .status {
      margin: 0;
      color: var(--muted);
      font-size: 0.95rem;
      line-height: 1.45;
    }

    .controls {
      display: grid;
      gap: 12px;
      margin-top: 14px;
    }

    .control-shell {
      display: grid;
      gap: 14px;
      margin-top: 14px;
    }

    .command-surface {
      position: relative;
      overflow: visible;
      background:
        linear-gradient(180deg, rgba(28, 32, 39, 0.98) 0%, rgba(17, 20, 25, 0.98) 100%);
      border-color: rgba(72, 78, 90, 0.75);
      box-shadow: 0 24px 60px rgba(14, 18, 24, 0.28);
      color: #f5efe6;
    }

    .command-topline {
      display: flex;
      align-items: center;
      gap: 12px;
      color: rgba(239, 227, 210, 0.84);
      font-size: 0.94rem;
      letter-spacing: 0.01em;
      margin-bottom: 16px;
    }

    .command-slash {
      color: rgba(239, 227, 210, 0.42);
    }

    .command-title {
      display: flex;
      align-items: start;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 16px;
    }

    .command-title h3 {
      margin: 0;
      font-size: 1.1rem;
      color: #f7f2eb;
    }

    .command-title p {
      margin: 6px 0 0;
      color: rgba(230, 220, 206, 0.7);
      max-width: 230px;
    }

    .command-badge {
      padding: 7px 11px;
      border-radius: 999px;
      background: rgba(255, 255, 255, 0.08);
      border: 1px solid rgba(255, 255, 255, 0.08);
      font-size: 0.76rem;
      font-weight: 700;
      color: rgba(247, 242, 235, 0.9);
      white-space: nowrap;
    }

    .menu-grid {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 10px;
    }

    .menu-node {
      position: relative;
    }

    .menu-trigger {
      width: 100%;
      border-radius: 18px;
      padding: 11px 14px;
      border: 1px solid transparent;
      background: transparent;
      color: rgba(61, 42, 24, 0.86);
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
      font: inherit;
      cursor: pointer;
      text-align: left;
      transition: background 140ms ease, border-color 140ms ease, transform 140ms ease;
    }

    .menu-trigger:hover,
    .menu-trigger.is-open {
      background: rgba(29, 107, 87, 0.1);
      border-color: rgba(201, 180, 148, 0.7);
      transform: translateY(-1px);
    }

    .menu-label strong {
      display: block;
      font-size: 0.98rem;
      color: #f7f2eb;
    }

    .menu-label span {
      display: block;
      margin-top: 4px;
      color: rgba(231, 221, 208, 0.7);
      font-size: 0.82rem;
      line-height: 1.35;
    }

    .menu-panel {
      position: absolute;
      top: calc(100% + 10px);
      right: 0;
      min-width: 280px;
      z-index: 30;
      display: none;
      padding: 14px;
      border-radius: 20px;
      border: 1px solid rgba(78, 83, 96, 0.72);
      background: rgba(9, 11, 16, 0.98);
      box-shadow: 0 28px 60px rgba(0, 0, 0, 0.4);
    }

    .menu-panel.is-open {
      display: grid;
      gap: 12px;
    }

    .menu-kicker {
      margin: 0;
      font-size: 0.76rem;
      letter-spacing: 0.08em;
      color: rgba(223, 213, 198, 0.52);
      font-weight: 700;
      text-transform: uppercase;
    }

    .menu-note {
      margin: 0;
      color: rgba(230, 220, 206, 0.74);
      font-size: 0.9rem;
      line-height: 1.45;
    }

    .menu-divider {
      height: 1px;
      background: rgba(255, 255, 255, 0.08);
    }

    .menu-panel button,
    .menu-panel .button-link {
      width: 100%;
      justify-content: flex-start;
      border-radius: 14px;
      background: transparent;
      border: 1px solid rgba(255, 255, 255, 0.08);
      padding: 13px 14px;
      color: #f7f1e8;
      font-weight: 600;
    }

    .menu-panel button:hover,
    .menu-panel .button-link:hover {
      background: rgba(255, 255, 255, 0.08);
      border-color: rgba(255, 255, 255, 0.12);
    }

    .devtools-shield {
      position: fixed;
      inset: 0;
      display: none;
      align-items: center;
      justify-content: center;
      padding: 24px;
      background: rgba(12, 15, 18, 0.82);
      backdrop-filter: blur(11px);
      z-index: 9999;
    }

    .devtools-shield.is-active {
      display: flex;
    }

    .devtools-shield-card {
      max-width: 430px;
      padding: 24px 26px;
      border-radius: 22px;
      border: 1px solid rgba(255, 255, 255, 0.12);
      background: rgba(12, 18, 23, 0.96);
      box-shadow: 0 30px 80px rgba(0, 0, 0, 0.36);
      color: #f6efe5;
      text-align: center;
      display: grid;
      gap: 10px;
    }

    .devtools-shield-card strong {
      font-size: 1.2rem;
      font-weight: 800;
    }

    .devtools-shield-card span {
      color: rgba(246, 239, 229, 0.8);
      line-height: 1.5;
    }

    .meta-card {
      display: grid;
      gap: 14px;
    }

    .summary-card {
      display: grid;
      gap: 16px;
    }

    .action-button {
      width: 100%;
      justify-content: flex-start;
      border-radius: 14px;
      padding: 12px 14px;
      background: rgba(255, 255, 255, 0.82);
      color: var(--ink);
      border: 1px solid rgba(217, 203, 185, 0.9);
      box-shadow: none;
    }

    .action-button:hover:not(:disabled) {
      background: rgba(29, 107, 87, 0.1);
    }

    .action-button:disabled {
      opacity: 0.48;
      cursor: not-allowed;
    }

    .summary-card h3 {
      margin-bottom: 8px;
    }

    .summary-card .status {
      padding: 12px 14px;
      border-radius: 14px;
      background: rgba(255, 255, 255, 0.58);
      border: 1px solid rgba(217, 203, 185, 0.84);
      color: var(--ink);
    }

    .header-actions {
      flex-wrap: wrap;
      justify-content: flex-end;
      gap: 14px;
    }

    .header-action {
      position: relative;
      border: 1px solid transparent;
    }

    .header-action::after {
      content: "";
      position: absolute;
      left: 12px;
      right: 12px;
      bottom: 6px;
      height: 2px;
      border-radius: 999px;
      background: transparent;
      transition: background 140ms ease;
    }

    .header-action:hover:not(:disabled)::after {
      background: rgba(29, 107, 87, 0.32);
    }

    .header-action.is-open {
      border-color: rgba(201, 180, 148, 0.7);
      background: rgba(255, 255, 255, 0.45);
    }

    .header-action:disabled {
      opacity: 0.48;
      cursor: not-allowed;
    }

    input[type="file"] {
      width: 100%;
      padding: 10px;
      border: 1px dashed var(--line);
      border-radius: 14px;
      background: rgba(255, 255, 255, 0.7);
    }

    .file-picker {
      display: grid;
      gap: 8px;
      cursor: pointer;
    }

    .file-input {
      position: absolute;
      width: 1px;
      height: 1px;
      padding: 0;
      margin: -1px;
      overflow: hidden;
      clip: rect(0, 0, 0, 0);
      white-space: nowrap;
      border: 0;
    }

    .file-shell {
      display: flex;
      align-items: center;
      gap: 10px;
      min-width: 0;
      padding: 8px;
      border-radius: 16px;
      border: 1px solid rgba(201, 180, 148, 0.9);
      background: rgba(255, 255, 255, 0.86);
      transition: border-color 140ms ease, box-shadow 140ms ease, transform 140ms ease;
    }

    .file-picker:hover .file-shell,
    .file-picker:focus-within .file-shell,
    .file-shell.has-file {
      border-color: rgba(29, 107, 87, 0.42);
      box-shadow: 0 10px 24px rgba(17, 75, 60, 0.08);
      transform: translateY(-1px);
    }

    .file-trigger {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      flex: 0 0 auto;
      min-height: 40px;
      padding: 0 14px;
      border-radius: 12px;
      background: linear-gradient(180deg, #fffaf3 0%, #f2e4cf 100%);
      border: 1px solid rgba(201, 180, 148, 0.96);
      color: #2f2418;
      font-size: 0.9rem;
      font-weight: 700;
      box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.7);
    }

    .file-name {
      min-width: 0;
      color: #7b7368;
      font-size: 0.92rem;
      line-height: 1.3;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .file-shell.has-file .file-name {
      color: #1f2a1f;
      font-weight: 600;
    }

    button,
    .button-link {
      appearance: none;
      border: 0;
      border-radius: 999px;
      padding: 11px 16px;
      background: var(--accent);
      color: white;
      font: inherit;
      font-weight: 600;
      cursor: pointer;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 100%;
    }

    button:disabled {
      opacity: 0.6;
      cursor: wait;
    }

    .button-link.is-hidden,
    button.is-hidden,
    input.is-hidden {
      display: none;
    }

    .summary {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 16px;
    }

    .pill {
      padding: 7px 10px;
      border-radius: 999px;
      background: var(--accent-soft);
      color: var(--accent);
      font-size: 0.88rem;
      font-weight: 600;
    }

    .items {
      display: grid;
      gap: 16px;
      margin-top: 16px;
    }

    .items-toolbar {
      display: flex;
      flex-wrap: wrap;
      align-items: end;
      gap: 12px;
      margin-top: 10px;
      padding: 12px;
      border: 1px solid rgba(217, 203, 185, 0.82);
      border-radius: 14px;
      background: rgba(255, 255, 255, 0.42);
    }

    .filter-field {
      display: grid;
      gap: 5px;
      min-width: 150px;
    }

    .filter-field label {
      color: var(--muted);
      font-size: 0.78rem;
      font-weight: 800;
      text-transform: uppercase;
      letter-spacing: 0;
    }

    .filter-field select {
      min-height: 38px;
      padding: 8px 10px;
      border-radius: 10px;
      font-size: 0.92rem;
    }

    .filter-count {
      margin-left: auto;
      color: var(--muted);
      font-size: 0.9rem;
      padding: 8px 2px;
    }

    .item {
      display: grid;
      grid-template-columns: 230px minmax(0, 1fr);
      gap: 16px;
      border: 1px solid var(--line);
      border-radius: 18px;
      padding: 16px;
      background: rgba(255, 255, 255, 0.55);
    }

    .item.equation-item {
      grid-template-columns: 320px minmax(0, 1fr);
    }

    .preview {
      min-height: 150px;
      border-radius: 14px;
      overflow: hidden;
      border: 1px solid rgba(217, 203, 185, 0.9);
      background: #fff;
      display: grid;
      place-items: center;
    }

    .item.equation-item .preview {
      min-height: 190px;
      padding: 10px;
      overflow-x: auto;
      overflow-y: hidden;
      display: flex;
      align-items: center;
      justify-content: flex-start;
      position: relative;
    }

    .preview img {
      width: auto;
      height: auto;
      max-width: 100%;
      max-height: 140px;
      display: block;
      background: white;
    }

    .item.equation-item .preview img {
      max-width: none;
      max-height: 170px;
    }

    .meta strong {
      display: block;
      margin-bottom: 6px;
      font-size: 1.05rem;
    }

    .meta p {
      margin: 4px 0;
      color: var(--muted);
    }

    .chips {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin: 10px 0 14px;
    }

    .chip {
      padding: 6px 10px;
      border-radius: 999px;
      font-size: 0.82rem;
      font-weight: 700;
    }

    .chip.good {
      background: var(--accent-soft);
      color: var(--accent);
    }

    .chip.warn {
      background: var(--warn-soft);
      color: var(--warn);
    }

    textarea {
      width: 100%;
      min-height: 120px;
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 12px;
      font-family: "Trebuchet MS", "Gill Sans", sans-serif;
      font-size: 0.98rem;
      resize: vertical;
      background: rgba(255, 255, 255, 0.9);
    }

    select {
      width: 100%;
      padding: 11px 12px;
      border: 1px solid var(--line);
      border-radius: 14px;
      font-family: "Trebuchet MS", "Gill Sans", sans-serif;
      font-size: 0.98rem;
      background: rgba(255, 255, 255, 0.9);
      color: var(--ink);
    }

    .empty {
      padding: 32px 18px;
      border: 1px dashed var(--line);
      border-radius: 18px;
      color: var(--muted);
      text-align: center;
      background: rgba(255, 255, 255, 0.45);
    }

    @media (max-width: 980px) {
      .appbar {
        flex-direction: column;
        align-items: stretch;
      }

      .appbar-left,
      .appbar-right {
        flex-wrap: wrap;
      }

      .layout {
        grid-template-columns: 1fr;
      }

      .menu-grid {
        grid-template-columns: 1fr;
      }

      .menu-panel {
        position: static;
        margin-top: 10px;
      }

      .item {
        grid-template-columns: 1fr;
      }

      .items-toolbar {
        align-items: stretch;
      }

      .filter-field {
        min-width: min(100%, 180px);
        flex: 1 1 150px;
      }

      .filter-count {
        width: 100%;
        margin-left: 0;
      }
    }
  </style>
</head>
<body>
  <div id="devtoolsShield" class="devtools-shield" aria-hidden="true">
    <div class="devtools-shield-card">
      <strong>Protected view active</strong>
      <span>Close developer tools to continue using this page.</span>
    </div>
  </div>
  <div class="shell">
    <header class="appbar panel">
      <div class="appbar-left">
        <div class="brandmark">HBS Alto</div>
      </div>

      <nav class="appbar-right header-actions" aria-label="Primary">
        <div class="menu-node">
          <button class="header-action menu-trigger" type="button" data-menu-trigger="uploadMenu" aria-expanded="false" aria-controls="uploadMenu">
            <strong>Upload</strong>
            <span class="menu-caret">&#9662;</span>
          </button>
          <div id="uploadMenu" class="menu-panel" data-menu-panel>
            <p class="menu-kicker">Upload</p>
            <p class="menu-note">Start a new document session or bring in a revised ALT workbook.</p>
            <div class="menu-divider"></div>
            <button id="quickProcessBtn" type="button">Process Document</button>
            <button id="quickImportBtn" type="button" disabled>Upload ALT Excel</button>
          </div>
        </div>

        <div class="menu-node">
          <button class="header-action menu-trigger" type="button" data-menu-trigger="downloadMenu" aria-expanded="false" aria-controls="downloadMenu">
            <strong>Download</strong>
            <span class="menu-caret">&#9662;</span>
          </button>
          <div id="downloadMenu" class="menu-panel" data-menu-panel>
            <p class="menu-kicker">Download</p>
            <p class="menu-note">Export the workbook for editing or produce the updated DOCX.</p>
            <div class="menu-divider"></div>
            <button id="quickDownloadBtn" type="button" disabled>Download Excel</button>
            <button id="quickUpdatedDocxBtn" type="button" disabled>Download Updated DOCX</button>
          </div>
        </div>

        <div class="menu-node">
          <button class="header-action menu-trigger" type="button" data-menu-trigger="imagesMenu" aria-expanded="false" aria-controls="imagesMenu">
            <strong>Images</strong>
            <span class="menu-caret">&#9662;</span>
          </button>
          <div id="imagesMenu" class="menu-panel" data-menu-panel>
            <p class="menu-kicker">Images</p>
            <p class="menu-note">Work with previews, generated ALT, and cleanup tasks for the active session.</p>
            <div class="menu-divider"></div>
            <button id="quickGridBtn" type="button" disabled>Download Preview Grid ZIP</button>
            <button id="quickClearBtn" type="button" disabled>Clear ALT</button>
            <button id="quickGenerateAltBtn" type="button" disabled>Generate ALT</button>
          </div>
        </div>
        <div class="menu-node">
          <button class="header-action menu-trigger" type="button" data-menu-trigger="toolsMenu" aria-expanded="false" aria-controls="toolsMenu" aria-label="Correction tools">
            <strong>&#9776;</strong>
          </button>
          <div id="toolsMenu" class="menu-panel" data-menu-panel>
            <p class="menu-kicker">Corrections</p>
            <p class="menu-note">Open document repair tools for list structure and color contrast cleanup.</p>
            <div class="menu-divider"></div>
            <a class="button-link" href="/list-correction">List Correction</a>
            <a class="button-link" href="/color-correction">Color Correction</a>
            <a class="button-link" href="/excel-merger">Excel Merger</a>
            <a class="button-link" href="/pdf-alt-editor">PDF ALT Editor</a>
          </div>
        </div>
      </nav>
    </header>

    <div class="is-hidden" aria-hidden="true">
      <button id="generateAltBtn" class="is-hidden" type="button">Generate ALT Text</button>
      <p id="generateAltState"></p>
      <a id="downloadBtn" class="button-link is-hidden" href="#" download>Download Excel</a>
      <a id="gridBtn" class="button-link is-hidden" href="#" download>Download Preview Grid ZIP</a>
      <input id="importExcelInput" class="is-hidden" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" />
      <button id="importExcelBtn" class="is-hidden" type="button">Import Updated Excel</button>
      <p id="importExcelState"></p>
      <button id="updatedDocxBtn" class="is-hidden" type="button">Download Updated DOCX</button>
      <button id="clearBtn" class="is-hidden" type="button">Clear ALT</button>
    </div>

    <section class="layout">
      <aside class="panel sidebar">
        <div class="card">
          <h3>Process File</h3>
          <p>Upload one DOCX to start an ALT session.</p>
          <div class="controls">
            <label class="file-picker" for="sourceFile">
              <input id="sourceFile" class="file-input" type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" />
              <span id="sourceFileShell" class="file-shell">
                <span class="file-trigger">Choose DOCX</span>
                <span id="sourceFileName" class="file-name">No file selected</span>
              </span>
            </label>
            <button id="processBtn" type="button">Process ALT Inventory</button>
          </div>
        </div>

        <div class="control-shell">
          <section class="card">
            <h3>ALT Provider</h3>
            <p>Choose which model generates missing ALT text for this session.</p>
            <div class="controls">
              <select id="altProviderSelect">
                <option value="claude">Claude</option>
                <option value="copilot">Copilot</option>
                <option value="gemini">Gemini</option>
                <option value="groq">Groq</option>
                <option value="openrouter">OpenRouter</option>
                <option value="claude_fallback_groq">Claude then Groq fallback</option>
              </select>
            </div>
          </section>

          <section class="card summary-card">
            <div>
              <h3>Session Summary</h3>
              <div class="summary" id="summary">
                <span class="pill">No session yet</span>
              </div>
            </div>
            <div>
              <h3>Session Status</h3>
              <p class="status" id="status">Ready.</p>
            </div>
          </section>
        </div>
      </aside>

      <main class="panel workspace">
        <h2 class="section-title">Collected Items</h2>
        <div class="items-toolbar" aria-label="Collected item filters">
          <div class="filter-field">
            <label for="typeFilter">Type</label>
            <select id="typeFilter">
              <option value="all">All types</option>
              <option value="equation">Equations</option>
              <option value="image">Images</option>
            </select>
          </div>
          <div class="filter-field">
            <label for="splitFilter">Split</label>
            <select id="splitFilter">
              <option value="all">All items</option>
              <option value="with-alt">With ALT</option>
              <option value="without-alt">Without ALT</option>
            </select>
          </div>
          <div class="filter-field">
            <label for="findReplaceOpenBtn">Find &amp; Replace</label>
            <button id="findReplaceOpenBtn" class="find-replace-launch" type="button">Find &amp; Replace</button>
          </div>
          <div id="filterCount" class="filter-count">0 shown</div>
        </div>
        <div id="findReplaceDialog" class="find-replace-dialog" hidden>
          <div class="find-replace-backdrop" data-find-replace-close="docx"></div>
          <div class="find-replace-card" role="dialog" aria-modal="true" aria-labelledby="findReplaceTitle">
            <div class="find-replace-head">
              <strong id="findReplaceTitle" class="find-replace-title">Find and Replace</strong>
              <button id="findReplaceCloseBtn" class="find-replace-close" type="button" aria-label="Close dialog">&times;</button>
            </div>
            <div class="find-replace-tabs" aria-hidden="true">
              <button class="find-replace-tab is-active" type="button" tabindex="-1">Find</button>
              <button class="find-replace-tab" type="button" tabindex="-1">Replace</button>
            </div>
            <div class="find-replace-body">
              <div class="find-replace-row">
                <label for="findText">Find what:</label>
                <input id="findText" type="text" placeholder="Search exact text" />
              </div>
              <div class="find-replace-row">
                <label for="replaceText">Replace with:</label>
                <input id="replaceText" type="text" placeholder="Optional replacement" />
              </div>
            </div>
            <div class="find-replace-actions">
              <button id="findMatchesBtn" type="button">Find Matches</button>
              <button id="replaceAllBtn" class="find-replace-primary" type="button">Replace All</button>
              <button class="find-replace-cancel" type="button" data-find-replace-close="docx">Cancel</button>
            </div>
          </div>
        </div>
        <div id="results" class="items">
          <div class="empty">Process a file to inspect its ALT items.</div>
        </div>
      </main>
    </section>
  </div>

  <script>
    const sourceFile = document.getElementById("sourceFile");
    const altProviderSelect = document.getElementById("altProviderSelect");
    const processBtn = document.getElementById("processBtn");
    const generateAltBtn = document.getElementById("generateAltBtn");
    const generateAltState = document.getElementById("generateAltState");
    const downloadBtn = document.getElementById("downloadBtn");
    const gridBtn = document.getElementById("gridBtn");
    const importExcelInput = document.getElementById("importExcelInput");
    const importExcelBtn = document.getElementById("importExcelBtn");
    const importExcelState = document.getElementById("importExcelState");
    const updatedDocxBtn = document.getElementById("updatedDocxBtn");
    const clearBtn = document.getElementById("clearBtn");
    const summary = document.getElementById("summary");
    const status = document.getElementById("status");
    const results = document.getElementById("results");
    const typeFilter = document.getElementById("typeFilter");
    const splitFilter = document.getElementById("splitFilter");
    const findReplaceOpenBtn = document.getElementById("findReplaceOpenBtn");
    const findReplaceDialog = document.getElementById("findReplaceDialog");
    const findReplaceCloseBtn = document.getElementById("findReplaceCloseBtn");
    const findReplaceDismissButtons = Array.from(findReplaceDialog.querySelectorAll("[data-find-replace-close]"));
    const findText = document.getElementById("findText");
    const replaceText = document.getElementById("replaceText");
    const findMatchesBtn = document.getElementById("findMatchesBtn");
    const replaceAllBtn = document.getElementById("replaceAllBtn");
    const filterCount = document.getElementById("filterCount");
    const sourceFileShell = document.getElementById("sourceFileShell");
    const sourceFileName = document.getElementById("sourceFileName");
    const quickProcessBtn = document.getElementById("quickProcessBtn");
    const quickDownloadBtn = document.getElementById("quickDownloadBtn");
    const quickGridBtn = document.getElementById("quickGridBtn");
    const quickClearBtn = document.getElementById("quickClearBtn");
    const quickImportBtn = document.getElementById("quickImportBtn");
    const quickUpdatedDocxBtn = document.getElementById("quickUpdatedDocxBtn");
    const quickGenerateAltBtn = document.getElementById("quickGenerateAltBtn");
    const menuTriggers = Array.from(document.querySelectorAll("[data-menu-trigger]"));
    const menuPanels = Array.from(document.querySelectorAll("[data-menu-panel]"));

    let altSession = null;
    let altRows = [];
    let altSourceKind = null;
    let altTextSearch = "";
    const saveTimers = new Map();
    const savePromises = new Map();

    const escapeHtml = (value) =>
      (value || "").replace(/[&<>"']/g, (char) => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;",
      }[char]));

    function enableUiHardening() {
      const devtoolsShield = document.getElementById("devtoolsShield");
      const blockedShiftCombos = new Set(["i", "j", "c", "k", "e", "m"]);
      const blockedAltCombos = new Set(["i", "j", "c"]);
      const blockedPlainCombos = new Set(["u"]);

      const stopEvent = (event) => {
        event.preventDefault();
        event.stopPropagation();
        if (typeof event.stopImmediatePropagation === "function") {
          event.stopImmediatePropagation();
        }
      };

      const setGuardLocked = (locked) => {
        document.body.classList.toggle("guard-locked", locked);
        devtoolsShield?.classList.toggle("is-active", locked);
        devtoolsShield?.setAttribute("aria-hidden", locked ? "false" : "true");
      };

      const stopDevtoolsAccess = (event) => {
        const key = String(event.key || "").toLowerCase();
        const ctrlOrMeta = event.ctrlKey || event.metaKey;
        const shouldBlock =
          key === "f12" ||
          (event.shiftKey && key === "f7") ||
          (ctrlOrMeta && blockedPlainCombos.has(key)) ||
          (ctrlOrMeta && event.shiftKey && blockedShiftCombos.has(key)) ||
          (ctrlOrMeta && event.altKey && blockedAltCombos.has(key));

        if (shouldBlock) {
          stopEvent(event);
        }
      };

      const blockContextAccess = (event) => {
        stopEvent(event);
      };

      const evaluateDevtoolsState = () => {
        const widthGap = Math.max(0, window.outerWidth - window.innerWidth);
        const heightGap = Math.max(0, window.outerHeight - window.innerHeight);
        const devtoolsLikelyOpen = widthGap > 220 || heightGap > 220;
        setGuardLocked(devtoolsLikelyOpen);
      };

      window.addEventListener("keydown", stopDevtoolsAccess, true);
      document.addEventListener("keydown", stopDevtoolsAccess, true);
      document.addEventListener("contextmenu", blockContextAccess, true);
      document.addEventListener("mousedown", (event) => {
        if (event.button === 2) {
          blockContextAccess(event);
        }
      }, true);
      document.addEventListener("auxclick", (event) => {
        if (event.button === 2) {
          blockContextAccess(event);
        }
      }, true);
      window.addEventListener("resize", evaluateDevtoolsState);
      window.addEventListener("focus", evaluateDevtoolsState);
      document.addEventListener("visibilitychange", () => {
        if (!document.hidden) {
          evaluateDevtoolsState();
        }
      });
      evaluateDevtoolsState();
      window.setInterval(evaluateDevtoolsState, 1200);
    }

    function closeAllMenus() {
      menuTriggers.forEach((trigger) => trigger.classList.remove("is-open"));
      menuTriggers.forEach((trigger) => trigger.setAttribute("aria-expanded", "false"));
      menuPanels.forEach((panel) => panel.classList.remove("is-open"));
    }

    function toggleMenu(menuId) {
      const panel = document.getElementById(menuId);
      const trigger = menuTriggers.find((item) => item.getAttribute("data-menu-trigger") === menuId);
      if (!panel || !trigger) {
        return;
      }
      const isOpen = panel.classList.contains("is-open");
      closeAllMenus();
      if (!isOpen) {
        panel.classList.add("is-open");
        trigger.classList.add("is-open");
        trigger.setAttribute("aria-expanded", "true");
      }
    }

    function currentDownloadHref(link) {
      const href = link?.getAttribute("href");
      if (!href || href === "#") {
        return "";
      }
      return href;
    }

    function selectedAltProvider() {
      const value = String(altProviderSelect?.value || "claude").trim().toLowerCase();
      if (value === "copilot" || value === "gemini" || value === "groq" || value === "openrouter" || value === "claude_fallback_groq") {
        return value;
      }
      return "claude";
    }

    function selectedAltProviderLabel() {
      const provider = selectedAltProvider();
      if (provider === "copilot") {
        return "Copilot";
      }
      if (provider === "gemini") {
        return "Gemini";
      }
      if (provider === "groq") {
        return "Groq";
      }
      if (provider === "openrouter") {
        return "OpenRouter";
      }
      if (provider === "claude_fallback_groq") {
        return "Claude with Groq fallback";
      }
      return "Claude";
    }

    function refreshQuickActions() {
      const hasSession = Boolean(altSession);
      const docxSession = hasSession && altSourceKind === "docx";
      quickDownloadBtn.disabled = !currentDownloadHref(downloadBtn);
      quickGridBtn.disabled = !currentDownloadHref(gridBtn);
      quickClearBtn.disabled = !docxSession;
      quickImportBtn.disabled = !docxSession;
      quickUpdatedDocxBtn.disabled = !docxSession;
      quickGenerateAltBtn.disabled = !hasSession;
      findReplaceOpenBtn.disabled = !docxSession;
      findMatchesBtn.disabled = !docxSession;
      replaceAllBtn.disabled = !docxSession;
    }

    function syncSourceFileState() {
      const [file] = sourceFile.files || [];
      sourceFileName.textContent = file ? file.name : "No file selected";
      sourceFileShell.classList.toggle("has-file", Boolean(file));
    }

    function filenameFromDisposition(headerValue, fallbackName) {
      if (!headerValue) {
        return fallbackName;
      }
      const utf8Match = headerValue.match(/filename\\*=UTF-8''([^;]+)/i);
      if (utf8Match && utf8Match[1]) {
        return decodeURIComponent(utf8Match[1]);
      }
      const basicMatch = headerValue.match(/filename="?([^";]+)"?/i);
      if (basicMatch && basicMatch[1]) {
        return basicMatch[1];
      }
      return fallbackName;
    }

    function triggerBlobDownload(blob, filename) {
      const objectUrl = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = objectUrl;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(objectUrl);
    }

    function altPreviewImage(itemId, attempt = 0) {
      const cacheKey = attempt ? `?v=${attempt}` : "";
      return `/api/alt/session/${altSession}/item/${itemId}/preview.png${cacheKey}`;
    }

    function effectiveAltText(row) {
      return String(row?.alt_text || row?.generated_alt_text || row?.existing_alt_text || "").trim();
    }

    function rowMatchesFilters(row) {
      const typeValue = String(typeFilter?.value || "all").toLowerCase();
      const splitValue = String(splitFilter?.value || "all").toLowerCase();
      const role = String(row?.role || "").toLowerCase();

      if (typeValue !== "all" && role !== typeValue) {
        return false;
      }

      const hasAlt = Boolean(effectiveAltText(row));
      if (splitValue === "with-alt" && !hasAlt) {
        return false;
      }
      if (splitValue === "without-alt" && hasAlt) {
        return false;
      }
      if (altTextSearch) {
        const currentAlt = String(row?.alt_text || "").toLowerCase();
        if (!currentAlt.includes(altTextSearch)) {
          return false;
        }
      }
      return true;
    }

    function filteredAltRows() {
      return altRows.filter(rowMatchesFilters);
    }

    function baseVisibleAltRows() {
      return altRows.filter((row) => {
        const typeValue = String(typeFilter?.value || "all").toLowerCase();
        const splitValue = String(splitFilter?.value || "all").toLowerCase();
        const role = String(row?.role || "").toLowerCase();
        if (typeValue !== "all" && role !== typeValue) {
          return false;
        }
        const hasAlt = Boolean(effectiveAltText(row));
        if (splitValue === "with-alt" && !hasAlt) {
          return false;
        }
        if (splitValue === "without-alt" && hasAlt) {
          return false;
        }
        return true;
      });
    }

    function updateFilterCount(visibleCount) {
      if (!filterCount) {
        return;
      }
      const total = altRows.length;
      filterCount.textContent = total ? `${visibleCount} of ${total} shown` : "0 shown";
    }

    function renderSummary(data) {
      if (!data) {
        summary.innerHTML = '<span class="pill">No session yet</span>';
        return;
      }
      const pills = [
        `${data.total_items || 0} item(s)`,
        `${data.images || 0} image(s)`,
        `${data.equations || 0} equation(s)`,
        `${data.with_alt_text || 0} original ALT`,
        `${data.without_alt_text || 0} missing`,
      ];
      summary.innerHTML = pills.map((label) => `<span class="pill">${escapeHtml(label)}</span>`).join("");
    }

    function renderResults() {
      if (!altRows.length) {
        updateFilterCount(0);
        results.innerHTML = '<div class="empty">No ALT items were found in this file.</div>';
        return;
      }

      const visibleRows = filteredAltRows();
      updateFilterCount(visibleRows.length);

      if (!visibleRows.length) {
        results.innerHTML = '<div class="empty">No ALT items match the selected filters.</div>';
        return;
      }

      results.innerHTML = visibleRows.map((row, index) => `
        <article class="item ${String(row.role || "").toLowerCase() === "equation" ? "equation-item" : ""}">
          <div class="preview">
            <img src="${altPreviewImage(row.id)}" alt="${escapeHtml(row.type)} preview ${index + 1}" loading="lazy" />
          </div>
          <div class="meta">
            <strong>#${index + 1} ${escapeHtml(row.type)}</strong>
            <p>${escapeHtml(row.source_part || "body")} · Page ${row.page || "-"}</p>
            <p>${row.existing_alt_text ? `Original ALT: ${escapeHtml(row.existing_alt_text)}` : "Original ALT missing for this item."}</p>
            <div class="chips">
              <span class="chip ${row.has_alt_text ? "good" : "warn"}">${row.has_alt_text ? "Original ALT" : "Original ALT missing"}</span>
              ${row.generated_alt_text ? '<span class="chip good">Generated ALT</span>' : ''}
              ${String(row.alt_source || "").startsWith("generated_claude") ? '<span class="chip good">Claude AI</span>' : ''}
              ${String(row.alt_source || "").startsWith("generated_copilot") ? '<span class="chip good">Copilot AI</span>' : ''}
              ${String(row.alt_source || "").startsWith("generated_gemini") ? '<span class="chip good">Gemini AI</span>' : ''}
              ${String(row.alt_source || "").startsWith("generated_openrouter") ? '<span class="chip good">OpenRouter AI</span>' : ''}
              ${String(row.alt_source || "").startsWith("generated_groq") ? '<span class="chip good">Groq AI</span>' : ''}
              ${row.alt_source === "manual" ? '<span class="chip good">Dashboard edit</span>' : ''}
              ${row.alt_source === "excel_import" ? '<span class="chip good">Excel import</span>' : ''}
            </div>
            <label for="alt-editor-${row.id}">Current ALT text</label>
            <textarea id="alt-editor-${row.id}" data-alt-editor="${row.id}" placeholder="Write ALT text for this item.">${escapeHtml(row.alt_text || "")}</textarea>
            <p class="note" data-save-state="${row.id}" style="margin-top: 8px;">${row.alt_source === "manual" ? "Saved" : ""}</p>
          </div>
        </article>
      `).join("");

      document.querySelectorAll("[data-alt-editor]").forEach((field) => {
        field.addEventListener("input", () => {
          const itemId = Number(field.getAttribute("data-alt-editor"));
          scheduleSave(itemId, field.value);
        });
        field.addEventListener("blur", async () => {
          const itemId = Number(field.getAttribute("data-alt-editor"));
          await flushPendingSaves();
          setSaveState(itemId, "Saved");
        });
      });
    }

    function setSaveState(itemId, message) {
      const el = document.querySelector(`[data-save-state="${itemId}"]`);
      if (el) {
        el.textContent = message;
      }
    }

    async function saveRow(itemId, altText) {
      if (!altSession || altSourceKind !== "docx") {
        return;
      }
      const promise = fetch(`/api/alt/session/${altSession}/item/${itemId}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ alt_text: altText }),
      })
        .then(async (response) => {
          const payload = await response.json();
          if (!response.ok) {
            throw new Error(payload.detail || "ALT save failed.");
          }
          altRows = altRows.map((row) => row.id === itemId ? payload.row : row);
          renderSummary(payload.summary || null);
          setSaveState(itemId, "Saved");
        })
        .catch((error) => {
          setSaveState(itemId, error.message);
          throw error;
        })
        .finally(() => {
          savePromises.delete(itemId);
        });

      savePromises.set(itemId, promise);
      return promise;
    }

    function scheduleSave(itemId, altText) {
      const row = altRows.find((entry) => entry.id === itemId);
      if (row) {
        row.alt_text = altText;
      }
      const timerId = saveTimers.get(itemId);
      if (timerId) {
        clearTimeout(timerId);
      }
      setSaveState(itemId, "Saving...");
      saveTimers.set(itemId, window.setTimeout(() => {
        saveTimers.delete(itemId);
        void saveRow(itemId, altText);
      }, 450));
    }

    async function flushPendingSaves() {
      const pendingTimers = Array.from(saveTimers.entries());
      pendingTimers.forEach(([itemId, timerId]) => {
        clearTimeout(timerId);
        saveTimers.delete(itemId);
        const row = altRows.find((entry) => entry.id === itemId);
        void saveRow(itemId, row?.alt_text || "");
      });

      const pendingPromises = Array.from(savePromises.values());
      if (pendingPromises.length) {
        const settled = await Promise.allSettled(pendingPromises);
        const rejected = settled.find((entry) => entry.status === "rejected");
        if (rejected && rejected.status === "rejected") {
          throw rejected.reason;
        }
      }
    }

    async function downloadUpdatedDocx() {
      if (!altSession || altSourceKind !== "docx") {
        status.textContent = "Updated DOCX download is available only for DOCX sessions.";
        return;
      }
      updatedDocxBtn.disabled = true;
      status.textContent = "Saving ALT changes and preparing the updated DOCX...";
      try {
        await flushPendingSaves();
        const response = await fetch(`/api/alt/session/${altSession}/updated.docx`);
        if (!response.ok) {
          let detail = "Updated DOCX download failed.";
          try {
            const payload = await response.json();
            detail = payload.detail || detail;
          } catch (_error) {
          }
          throw new Error(detail);
        }
        const blob = await response.blob();
        const filename = filenameFromDisposition(response.headers.get("content-disposition"), "updated_alt.docx");
        triggerBlobDownload(blob, filename);
        status.textContent = "Updated DOCX is ready.";
      } catch (error) {
        status.textContent = error.message;
      } finally {
        updatedDocxBtn.disabled = false;
      }
    }

    async function importUpdatedExcel(file) {
      if (!altSession || altSourceKind !== "docx") {
        status.textContent = "Excel import is available only for DOCX sessions.";
        return;
      }
      if (!file) {
        return;
      }
      importExcelBtn.disabled = true;
      importExcelState.textContent = "Importing updated workbook...";
      status.textContent = "Refreshing the ALT session from Excel...";
      const formData = new FormData();
      formData.append("workbook", file);
      try {
        const response = await fetch(`/api/alt/session/${altSession}/import-excel`, {
          method: "POST",
          body: formData,
        });
        const payload = await response.json();
        if (!response.ok) {
          throw new Error(payload.detail || "Excel import failed.");
        }
        altRows = payload.rows || [];
        renderSummary(payload.summary || null);
        renderResults();
        importExcelState.textContent = "Updated Excel imported successfully.";
        status.textContent = "ALT session updated from Excel.";
      } catch (error) {
        importExcelState.textContent = error.message;
        status.textContent = error.message;
      } finally {
        importExcelBtn.disabled = false;
        importExcelInput.value = "";
      }
    }

    async function generateAltText() {
      if (!altSession) {
        status.textContent = "Process a file before generating ALT text.";
        return;
      }
      const provider = selectedAltProvider();
      const providerLabel = selectedAltProviderLabel();
      generateAltBtn.disabled = true;
      generateAltState.textContent = `Generating ALT text with ${providerLabel}...`;
      status.textContent = `Saving changes and generating missing ALT text with ${providerLabel}...`;
      try {
        await flushPendingSaves();
        const response = await fetch(`/api/alt/session/${altSession}/generate-alt`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ provider }),
        });
        const payload = await response.json();
        if (!response.ok) {
          throw new Error(payload.detail || "ALT generation failed.");
        }
        altRows = payload.rows || [];
        renderSummary(payload.summary || null);
        renderResults();
        generateAltState.textContent = payload.message || "Generated ALT text.";
        status.textContent = payload.message || "Generated ALT text.";
      } catch (error) {
        generateAltState.textContent = error.message;
        status.textContent = error.message;
      } finally {
        generateAltBtn.disabled = false;
      }
    }

    async function clearAltSession() {
      if (!altSession || altSourceKind !== "docx") {
        status.textContent = "Clear ALT is available only for DOCX sessions.";
        return;
      }

      clearBtn.disabled = true;
      quickClearBtn.disabled = true;
      status.textContent = "Clearing ALT text in the current session...";
      try {
        await flushPendingSaves();
        const response = await fetch(`/api/alt/session/${altSession}/clear-alt`, {
          method: "POST",
        });
        const payload = await response.json();
        if (!response.ok) {
          throw new Error(payload.detail || "Clear ALT failed.");
        }
        altRows = payload.rows || [];
        renderSummary(payload.summary || null);
        renderResults();
        status.textContent = payload.message || "Cleared ALT text in this session.";
      } catch (error) {
        status.textContent = error.message;
      } finally {
        clearBtn.disabled = false;
        refreshQuickActions();
      }
    }

    function openFindReplaceDialog() {
      findReplaceDialog.hidden = false;
      window.setTimeout(() => findText.focus(), 0);
    }

    function closeFindReplaceDialog() {
      findReplaceDialog.hidden = true;
    }

    function applyFindMatches() {
      const searchText = String(findText?.value || "").trim().toLowerCase();
      altTextSearch = searchText;
      renderResults();
      if (!searchText) {
        status.textContent = "Showing all ALT items.";
      } else {
        status.textContent = `Showing ${filteredAltRows().length} ALT item(s) matching "${findText.value}".`;
      }
      closeFindReplaceDialog();
    }

    async function replaceAltTextAcrossVisibleRows() {
      if (!altSession || altSourceKind !== "docx") {
        status.textContent = "Find and Replace is available only for DOCX sessions.";
        return;
      }
      const searchTextValue = String(findText?.value || "");
      const replaceTextValue = String(replaceText?.value || "");
      if (!searchTextValue) {
        status.textContent = "Enter the exact text you want to find first.";
        return;
      }
      const targetIds = baseVisibleAltRows()
        .filter((row) => String(row?.alt_text || "").includes(searchTextValue))
        .map((row) => Number(row.id))
        .filter((itemId) => Number.isInteger(itemId));
      if (!targetIds.length) {
        status.textContent = "No visible ALT text matches that search string.";
        return;
      }
      replaceAllBtn.disabled = true;
      try {
        await flushPendingSaves();
        const response = await fetch(`/api/alt/session/${altSession}/replace-alt-text`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            search_text: searchTextValue,
            replace_text: replaceTextValue,
            item_ids: targetIds,
          }),
        });
        const payload = await response.json();
        if (!response.ok) {
          throw new Error(payload.detail || "Find and Replace failed.");
        }
        altRows = payload.rows || [];
        altTextSearch = "";
        renderSummary(payload.summary || null);
        renderResults();
        status.textContent = payload.message || "Updated ALT text.";
        closeFindReplaceDialog();
      } catch (error) {
        status.textContent = error.message;
      } finally {
        refreshQuickActions();
      }
    }

    processBtn.addEventListener("click", async () => {
      const [file] = sourceFile.files || [];
      if (!file) {
        status.textContent = "Choose a DOCX first.";
        return;
      }

      processBtn.disabled = true;
      status.textContent = "Collecting ALT inventory...";
      results.innerHTML = '<div class="empty">Working on the ALT inventory...</div>';
      renderSummary(null);
      closeAllMenus();
      refreshQuickActions();

      generateAltBtn.classList.add("is-hidden");
      generateAltBtn.disabled = false;
      generateAltState.textContent = "Process a file, then generate missing ALT text with the selected provider.";
      downloadBtn.classList.add("is-hidden");
      downloadBtn.removeAttribute("href");
      gridBtn.classList.add("is-hidden");
      gridBtn.removeAttribute("href");
      importExcelBtn.classList.add("is-hidden");
      importExcelBtn.disabled = false;
      updatedDocxBtn.classList.add("is-hidden");
      updatedDocxBtn.disabled = false;
      clearBtn.classList.add("is-hidden");
      clearBtn.disabled = false;
      findReplaceOpenBtn.disabled = true;
      findMatchesBtn.disabled = true;
      replaceAllBtn.disabled = true;
      altTextSearch = "";
      saveTimers.forEach((timerId) => clearTimeout(timerId));
      saveTimers.clear();
      savePromises.clear();

      try {
        const formData = new FormData();
        formData.append("file", file);
        const response = await fetch("/api/alt/analyze", {
          method: "POST",
          body: formData,
        });
        const payload = await response.json();
        if (!response.ok) {
          throw new Error(payload.detail || "ALT inventory failed.");
        }

        altSession = payload.session_id;
        altRows = payload.rows || [];
        altSourceKind = payload.source_kind || null;

        renderSummary(payload.summary || null);
        renderResults();
        status.textContent = payload.message || `Collected ${altRows.length} ALT item(s).`;

        if (altSession) {
          generateAltBtn.classList.remove("is-hidden");
          downloadBtn.href = `/api/alt/session/${altSession}/download.xlsx`;
          downloadBtn.classList.remove("is-hidden");
          gridBtn.href = `/api/alt/session/${altSession}/grids.zip`;
          gridBtn.classList.remove("is-hidden");
          if (altSourceKind === "docx") {
            importExcelBtn.classList.remove("is-hidden");
            updatedDocxBtn.classList.remove("is-hidden");
            clearBtn.classList.remove("is-hidden");
          }
          refreshQuickActions();
        }
      } catch (error) {
        status.textContent = error.message;
        results.innerHTML = `<div class="empty">${escapeHtml(error.message)}</div>`;
      } finally {
        processBtn.disabled = false;
        refreshQuickActions();
      }
    });

    quickProcessBtn.addEventListener("click", () => {
      closeAllMenus();
      processBtn.click();
    });
    quickDownloadBtn.addEventListener("click", () => {
      closeAllMenus();
      const href = currentDownloadHref(downloadBtn);
      if (href) {
        window.location.href = href;
      }
    });
    quickGridBtn.addEventListener("click", () => {
      closeAllMenus();
      const href = currentDownloadHref(gridBtn);
      if (href) {
        window.location.href = href;
      }
    });
    quickClearBtn.addEventListener("click", () => {
      closeAllMenus();
      void clearAltSession();
    });
    quickImportBtn.addEventListener("click", () => {
      closeAllMenus();
      importExcelInput.click();
    });
    quickUpdatedDocxBtn.addEventListener("click", async () => {
      closeAllMenus();
      await downloadUpdatedDocx();
    });
    quickGenerateAltBtn.addEventListener("click", async () => {
      closeAllMenus();
      await generateAltText();
    });
    importExcelBtn.addEventListener("click", () => importExcelInput.click());
    importExcelInput.addEventListener("change", async () => {
      const [file] = importExcelInput.files || [];
      await importUpdatedExcel(file);
    });
    findReplaceOpenBtn.addEventListener("click", () => {
      openFindReplaceDialog();
    });
    findReplaceCloseBtn.addEventListener("click", () => {
      closeFindReplaceDialog();
    });
    findReplaceDismissButtons.forEach((button) => {
      button.addEventListener("click", () => {
        closeFindReplaceDialog();
      });
    });
    findMatchesBtn.addEventListener("click", () => {
      applyFindMatches();
    });
    replaceAllBtn.addEventListener("click", async () => {
      await replaceAltTextAcrossVisibleRows();
    });
    typeFilter.addEventListener("change", renderResults);
    splitFilter.addEventListener("change", renderResults);
    sourceFile.addEventListener("change", syncSourceFileState);
    menuTriggers.forEach((trigger) => {
      trigger.addEventListener("click", (event) => {
        event.stopPropagation();
        toggleMenu(trigger.getAttribute("data-menu-trigger"));
      });
    });
    document.addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof Element)) {
        closeAllMenus();
        return;
      }
      if (!target.closest("[data-menu-trigger]") && !target.closest("[data-menu-panel]")) {
        closeAllMenus();
      }
    });
    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        closeFindReplaceDialog();
        closeAllMenus();
      }
    });
    generateAltBtn.addEventListener("click", async () => {
      await generateAltText();
    });
    updatedDocxBtn.addEventListener("click", async () => {
      await downloadUpdatedDocx();
    });
    clearBtn.addEventListener("click", async () => {
      await clearAltSession();
    });
    syncSourceFileState();
    refreshQuickActions();
    enableUiHardening();
  </script>
</body>
</html>
        """
    )


@app.post("/api/alt/analyze")
async def analyze_alt_document(file: UploadFile = File(...)):
    filename = Path(file.filename or "upload.bin").name
    suffix = Path(filename).suffix.lower()
    if suffix != ".docx":
        raise HTTPException(status_code=400, detail="HBS Alto supports DOCX files only.")

    session_id = uuid.uuid4().hex
    session_dir = TEMP_DIR / session_id
    session_dir.mkdir(parents=True, exist_ok=True)
    source_path = write_upload(session_dir, file)
    validate_uploaded_docx(source_path)

    try:
        inventory = build_alt_inventory(source_path, None)

        if not inventory.get("available"):
            raise HTTPException(status_code=422, detail=inventory.get("message", "ALT inventory could not be generated."))

        pdf_path = None
        preview_images = build_alt_preview_images(inventory["rows"], source_path, pdf_path, roles={"image"})
        ALT_SESSIONS[session_id] = {
            "source_path": source_path,
            "pdf_path": pdf_path,
            "source_kind": suffix.lstrip("."),
            "source_filename": source_path.name,
            "rows": inventory["rows"],
            "summary": inventory["summary"],
            "preview_images": preview_images,
            "equation_previews_status": "queued",
        }
        asyncio.create_task(asyncio.to_thread(build_equation_previews_for_session, session_id))

        return JSONResponse(
            {
                "session_id": session_id,
                "rows": inventory["rows"],
                "summary": inventory["summary"],
                "source_kind": "docx",
                "message": f"Collected {inventory['summary']['total_items']} ALT item(s).",
            }
        )
    except HTTPException:
        raise
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.get("/api/alt/session/{session_id}/item/{item_id}/preview.png")
async def get_alt_item_preview(session_id: str, item_id: int):
    session = ALT_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="ALT session not found.")

    preview_images = session.setdefault("preview_images", {})
    preview_entry = preview_images.get(item_id)
    if not isinstance(preview_entry, dict) or not preview_entry.get("bytes"):
        row = next((entry for entry in session.get("rows", []) if entry.get("id") == item_id), None)
        if row and str(row.get("role", "")).lower() == "equation":
            try:
                preview_entry = await asyncio.to_thread(
                    build_alt_preview_entry,
                    row,
                    session.get("source_path"),
                    session.get("pdf_path"),
                )
            except Exception as exc:
                session.setdefault("equation_preview_errors", {})[item_id] = str(exc)
                preview_entry = None
            if isinstance(preview_entry, dict) and preview_entry.get("bytes"):
                preview_images[item_id] = preview_entry

    if not isinstance(preview_entry, dict) or not preview_entry.get("bytes"):
        raise HTTPException(status_code=404, detail="Preview not found.")
    return preview_response(preview_entry)


@app.post("/api/alt/session/{session_id}/item/{item_id}")
async def update_alt_item(session_id: str, item_id: int, request: Request):
    session = ALT_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="ALT session not found.")

    if session.get("source_kind") != "docx":
        raise HTTPException(status_code=400, detail="Dashboard editing is available only for DOCX ALT sessions.")

    row = next((entry for entry in session.get("rows", []) if entry.get("id") == item_id), None)
    if row is None:
        raise HTTPException(status_code=404, detail="ALT item not found.")

    payload = await request.json()
    alt_text = normalize_alt_text(str(payload.get("alt_text", "") or ""))
    existing_alt = normalize_alt_text(str(row.get("existing_alt_text", "") or ""))
    row["alt_text"] = alt_text
    row["effective_alt_text"] = alt_text
    if alt_text == existing_alt:
        row["alt_source"] = "existing" if existing_alt else "missing"
    elif alt_text:
        row["alt_source"] = "manual"
    else:
        row["alt_source"] = "missing"

    session["summary"] = summarize_alt_rows(session.get("rows", []))
    return JSONResponse({"row": row, "summary": session["summary"]})


@app.get("/api/alt/session/{session_id}/download.xlsx")
async def download_alt_inventory_excel(session_id: str):
    session = ALT_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="ALT session not found.")

    workbook_bytes = build_alt_excel(
        session.get("rows", []),
        session.get("source_filename", "document.docx"),
        preview_images=session.get("preview_images") or {},
    )
    filename = f"{safe_download_stem(session.get('source_filename', 'document.docx'))}_alt_inventory.xlsx"
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return Response(
        content=workbook_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.get("/api/alt/session/{session_id}/grids.zip")
async def download_alt_grid_archive(session_id: str):
    session = ALT_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="ALT session not found.")

    try:
        workbook_bytes = build_alt_excel(
            session.get("rows", []),
            session.get("source_filename", "document.docx"),
            preview_images=session.get("preview_images") or {},
        )
        images = extract_excel_images(workbook_bytes)
        if not images:
            raise HTTPException(status_code=422, detail="No embedded images were found in the ALT workbook.")
        grid_files = make_grids(images)
    except HTTPException:
        raise
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    archive_buffer = io.BytesIO()
    with zipfile.ZipFile(archive_buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, image_bytes in grid_files:
            archive.writestr(name, image_bytes)

    filename = f"{safe_download_stem(session.get('source_filename', 'document.docx'))}_alt_grids.zip"
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return Response(content=archive_buffer.getvalue(), media_type="application/zip", headers=headers)


@app.post("/api/alt/session/{session_id}/import-excel")
async def import_alt_inventory_excel(session_id: str, workbook: UploadFile = File(...)):
    session = ALT_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="ALT session not found.")

    if session.get("source_kind") != "docx":
        raise HTTPException(status_code=400, detail="Excel import is available only for DOCX ALT sessions.")

    try:
        workbook_bytes = await workbook.read()
        imported_alt_texts = parse_alt_injection_workbook(
            workbook_bytes,
            session.get("rows", []),
            session.get("source_filename"),
        )

        for index, row in enumerate(session.get("rows", [])):
            imported_alt = normalize_alt_text(imported_alt_texts[index] if index < len(imported_alt_texts) else "")
            existing_alt = normalize_alt_text(str(row.get("existing_alt_text", "") or ""))
            row["alt_text"] = imported_alt
            row["effective_alt_text"] = imported_alt
            if imported_alt == existing_alt:
                row["alt_source"] = "existing" if existing_alt else "missing"
            elif imported_alt:
                row["alt_source"] = "excel_import"
            else:
                row["alt_source"] = "missing"

        session["summary"] = summarize_alt_rows(session.get("rows", []))
        return JSONResponse({"rows": session.get("rows", []), "summary": session["summary"]})
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.post("/api/alt/session/{session_id}/generate-alt")
async def generate_alt_text(session_id: str, request: Request):
    session = ALT_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="ALT session not found.")

    try:
        source_path = session.get("source_path") if isinstance(session.get("source_path"), Path) else None
        rows = session.get("rows", [])
        preview_images = session.setdefault("preview_images", {})
        missing_equation_preview = any(
            isinstance(row.get("id"), int)
            and str(row.get("role", "")).lower() == "equation"
            and row.get("id") not in preview_images
            for row in rows
        )
        if missing_equation_preview:
            await asyncio.to_thread(build_equation_previews_for_session, session_id)
        preview_images = session.get("preview_images") or {}
        payload = {}
        try:
            payload = await request.json()
        except Exception:
            payload = {}

        requested_provider = normalize_alt_text(str((payload or {}).get("provider", "") or "")).lower()
        generation_mode = requested_provider or (os.getenv("MATCHA_ALT_GENERATION_MODE") or "groq").strip().lower()
        if generation_mode not in {"claude", "copilot", "gemini", "groq", "openrouter", "claude_fallback_groq"}:
            generation_mode = "groq"

        if generation_mode == "claude":
            claude_generation = generate_missing_alt_rows_with_claude(
                rows,
                preview_images,
                None,
                source_path,
            )
            if not claude_generation.get("available"):
                raise HTTPException(status_code=422, detail=claude_generation.get("message", "Claude ALT generation is unavailable."))
            session["rows"] = claude_generation.get("rows", rows)
            session["summary"] = claude_generation.get("summary", summarize_alt_rows(session.get("rows", [])))
            message = claude_generation.get("message", "Generated ALT text with Claude.")
            provider = claude_generation.get("provider", "claude")
            generated_count = int(claude_generation.get("generated_count", 0) or 0)
        elif generation_mode == "copilot":
            copilot_generation = generate_missing_alt_rows_with_copilot(
                rows,
                preview_images,
                None,
                source_path,
            )
            if not copilot_generation.get("available"):
                raise HTTPException(status_code=422, detail=copilot_generation.get("message", "Copilot ALT generation is unavailable."))
            session["rows"] = copilot_generation.get("rows", rows)
            session["summary"] = copilot_generation.get("summary", summarize_alt_rows(session.get("rows", [])))
            message = copilot_generation.get("message", "Generated ALT text with Copilot.")
            provider = copilot_generation.get("provider", "copilot")
            generated_count = int(copilot_generation.get("generated_count", 0) or 0)
        elif generation_mode == "gemini":
            gemini_generation = generate_missing_alt_rows_with_gemini(
                rows,
                preview_images,
                None,
                source_path,
            )
            if not gemini_generation.get("available"):
                raise HTTPException(status_code=422, detail=gemini_generation.get("message", "Gemini ALT generation is unavailable."))
            session["rows"] = gemini_generation.get("rows", rows)
            session["summary"] = gemini_generation.get("summary", summarize_alt_rows(session.get("rows", [])))
            message = gemini_generation.get("message", "Generated ALT text with Gemini.")
            provider = gemini_generation.get("provider", "gemini")
            generated_count = int(gemini_generation.get("generated_count", 0) or 0)
        elif generation_mode == "openrouter":
            openrouter_generation = generate_missing_alt_rows_with_openrouter(
                rows,
                preview_images,
                None,
                source_path,
            )
            if not openrouter_generation.get("available"):
                raise HTTPException(status_code=422, detail=openrouter_generation.get("message", "OpenRouter ALT generation is unavailable."))
            session["rows"] = openrouter_generation.get("rows", rows)
            session["summary"] = openrouter_generation.get("summary", summarize_alt_rows(session.get("rows", [])))
            message = openrouter_generation.get("message", "Generated ALT text with OpenRouter.")
            provider = openrouter_generation.get("provider", "openrouter")
            generated_count = int(openrouter_generation.get("generated_count", 0) or 0)
        elif generation_mode == "claude_fallback_groq":
            claude_generation = generate_missing_alt_rows_with_claude(
                rows,
                preview_images,
                None,
                source_path,
            )
            if not claude_generation.get("available"):
                groq_generation = generate_missing_alt_rows_with_groq(
                    rows,
                    preview_images,
                    None,
                    source_path,
                )
                if not groq_generation.get("available"):
                    raise HTTPException(status_code=422, detail=groq_generation.get("message", "Groq ALT generation is unavailable."))
                session["rows"] = groq_generation.get("rows", rows)
                session["summary"] = groq_generation.get("summary", summarize_alt_rows(session.get("rows", [])))
                message = f"{claude_generation.get('message', 'Claude ALT generation is unavailable.')} Fallback: {groq_generation.get('message', 'Generated ALT text with Groq.')}"
                provider = "claude_fallback_groq"
                generated_count = int(groq_generation.get("generated_count", 0) or 0)
            else:
                rows_after_claude = claude_generation.get("rows", rows)
                claude_generated = int(claude_generation.get("generated_count", 0) or 0)
                groq_generation = generate_missing_alt_rows_with_groq(
                    rows_after_claude,
                    preview_images,
                    None,
                    source_path,
                )
                groq_generated = 0
                if groq_generation.get("available"):
                    rows_after_claude = groq_generation.get("rows", rows_after_claude)
                    groq_generated = int(groq_generation.get("generated_count", 0) or 0)
                session["rows"] = rows_after_claude
                session["summary"] = summarize_alt_rows(session.get("rows", []))
                provider = "claude_fallback_groq" if groq_generated else "claude"
                generated_count = claude_generated + groq_generated
                message = claude_generation.get("message", "Generated ALT text with Claude.")
                if groq_generated:
                    provider = "claude_fallback_groq"
                    message += f" Fallback Groq generated {groq_generated} additional item(s)."
        else:
            groq_generation = generate_missing_alt_rows_with_groq(
                rows,
                preview_images,
                None,
                source_path,
            )
            if not groq_generation.get("available"):
                raise HTTPException(status_code=422, detail=groq_generation.get("message", "Groq ALT generation is unavailable."))
            session["rows"] = groq_generation.get("rows", rows)
            session["summary"] = groq_generation.get("summary", summarize_alt_rows(session.get("rows", [])))
            message = groq_generation.get("message", "Generated ALT text with Groq.")
            provider = groq_generation.get("provider", "groq")
            generated_count = int(groq_generation.get("generated_count", 0) or 0)

        return JSONResponse(
            {
                "rows": session.get("rows", []),
                "summary": session["summary"],
                "message": message,
                "generated_count": generated_count,
                "provider": provider,
            }
        )
    except HTTPException:
        raise
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.get("/api/alt/session/{session_id}/updated.docx")
async def download_updated_docx(session_id: str):
    session = ALT_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="ALT session not found.")

    source_path = session.get("source_path")
    if not isinstance(source_path, Path) or source_path.suffix.lower() != ".docx":
        raise HTTPException(status_code=400, detail="Updated DOCX download is available only for DOCX ALT sessions.")

    try:
        rows = session.get("rows", [])
        alt_texts = [normalize_alt_text(str(row.get("alt_text", "") or "")) for row in rows]
        docx_bytes, _applied_count = inject_alt_texts_into_docx(source_path, alt_texts, rows)
    except ValueError as exc:
        print(f"[Altomizer] Updated DOCX rejected for session {session_id}: {exc}")
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    filename = f"{safe_download_stem(session.get('source_filename', source_path.name))}_updated_alt.docx"
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return Response(
        content=docx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers,
    )


@app.post("/api/alt/session/{session_id}/clear-alt")
async def clear_alt_text_from_session(session_id: str):
    session = ALT_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="ALT session not found.")

    source_path = session.get("source_path")
    if not isinstance(source_path, Path) or source_path.suffix.lower() != ".docx":
        raise HTTPException(status_code=400, detail="Clear ALT is available only for DOCX ALT sessions.")

    rows = session.get("rows", [])
    cleared_count = 0
    for row in rows:
        previous_alt = normalize_alt_text(str(row.get("alt_text", "") or ""))
        if previous_alt:
            cleared_count += 1
        row["alt_text"] = ""
        row["effective_alt_text"] = ""
        row["alt_source"] = "missing"

    session["summary"] = summarize_alt_rows(rows)
    message = f"Cleared ALT text for {cleared_count} item(s) in this session."
    return JSONResponse({"rows": rows, "summary": session["summary"], "message": message})


@app.post("/api/alt/session/{session_id}/replace-alt-text")
async def replace_alt_text_in_session(session_id: str, request: Request):
    session = ALT_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="ALT session not found.")

    source_path = session.get("source_path")
    if not isinstance(source_path, Path) or source_path.suffix.lower() != ".docx":
        raise HTTPException(status_code=400, detail="Find and Replace is available only for DOCX ALT sessions.")

    payload = await request.json()
    search_text = str(payload.get("search_text", "") or "")
    replace_text = str(payload.get("replace_text", "") or "")
    raw_item_ids = payload.get("item_ids", [])
    if not search_text:
        raise HTTPException(status_code=422, detail="Search text is required.")

    target_ids = {
        int(item_id)
        for item_id in raw_item_ids
        if isinstance(item_id, int) or (isinstance(item_id, str) and item_id.strip().isdigit())
    }
    if not target_ids:
        raise HTTPException(status_code=422, detail="Choose at least one ALT item to update.")

    rows = session.get("rows", [])
    replaced_count = 0
    touched_count = 0
    for row in rows:
        row_id = row.get("id")
        if not isinstance(row_id, int) or row_id not in target_ids:
            continue
        current_alt = str(row.get("alt_text", "") or "")
        if search_text not in current_alt:
            continue
        updated_alt = current_alt.replace(search_text, replace_text)
        if updated_alt == current_alt:
            continue
        touched_count += 1
        replaced_count += current_alt.count(search_text)
        normalized_alt = normalize_alt_text(updated_alt)
        existing_alt = normalize_alt_text(str(row.get("existing_alt_text", "") or ""))
        row["alt_text"] = normalized_alt
        row["effective_alt_text"] = normalized_alt
        if normalized_alt == existing_alt:
            row["alt_source"] = "existing" if existing_alt else "missing"
        elif normalized_alt:
            row["alt_source"] = "manual"
        else:
            row["alt_source"] = "missing"

    session["summary"] = summarize_alt_rows(rows)
    message = f"Replaced {replaced_count} occurrence(s) across {touched_count} item(s)."
    return JSONResponse({"rows": rows, "summary": session["summary"], "message": message})


@app.post("/api/pdf-alt/analyze")
async def analyze_pdf_alt_document(file: UploadFile = File(...)):
    filename = Path(file.filename or "upload.bin").name
    suffix = Path(filename).suffix.lower()
    if suffix != ".pdf":
        raise HTTPException(status_code=400, detail="PDF ALT Editor supports PDF files only.")

    session_id = uuid.uuid4().hex
    session_dir = TEMP_DIR / f"pdf_alt_{session_id}"
    session_dir.mkdir(parents=True, exist_ok=True)
    source_path = write_upload(session_dir, file)
    validation_error = validate_pdf_file(source_path)
    if validation_error:
        raise HTTPException(status_code=422, detail=validation_error)

    try:
        inventory = build_pdf_alt_inventory(source_path)
        if not inventory.get("available"):
            raise HTTPException(status_code=422, detail=inventory.get("message", "PDF ALT inventory could not be generated."))

        preview_images = build_pdf_preview_images(inventory["rows"], source_path)
        PDF_ALT_SESSIONS[session_id] = {
            "source_path": source_path,
            "source_kind": "pdf",
            "source_filename": source_path.name,
            "rows": inventory["rows"],
            "summary": inventory["summary"],
            "preview_images": preview_images,
        }

        return JSONResponse(
            {
                "session_id": session_id,
                "rows": inventory["rows"],
                "summary": inventory["summary"],
                "source_kind": "pdf",
                "message": inventory.get("message", f"Collected {inventory['summary']['total_items']} PDF ALT item(s)."),
            }
        )
    except HTTPException:
        raise
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.get("/api/pdf-alt/session/{session_id}/item/{item_id}/preview.png")
async def get_pdf_alt_item_preview(session_id: str, item_id: int):
    session = PDF_ALT_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="PDF ALT session not found.")

    preview_images = session.setdefault("preview_images", {})
    preview_entry = preview_images.get(item_id)
    if preview_entry is None:
        row = next((entry for entry in session.get("rows", []) if entry.get("id") == item_id), None)
        source_path = session.get("source_path") if isinstance(session.get("source_path"), Path) else None
        if row is not None and isinstance(source_path, Path):
            generated = build_pdf_preview_images([row], source_path)
            preview_images.update(generated)
            preview_entry = preview_images.get(item_id)
    if preview_entry is None:
        raise HTTPException(status_code=404, detail="Preview not found.")
    return preview_response(preview_entry)


@app.post("/api/pdf-alt/session/{session_id}/item/{item_id}")
async def update_pdf_alt_item(session_id: str, item_id: int, request: Request):
    session = PDF_ALT_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="PDF ALT session not found.")

    row = next((entry for entry in session.get("rows", []) if entry.get("id") == item_id), None)
    if row is None:
        raise HTTPException(status_code=404, detail="PDF ALT item not found.")

    payload = await request.json()
    alt_text = normalize_alt_text(str(payload.get("alt_text", "") or ""))
    existing_alt = normalize_alt_text(str(row.get("existing_alt_text", "") or ""))
    row["alt_text"] = alt_text
    row["effective_alt_text"] = alt_text
    if alt_text == existing_alt:
        row["alt_source"] = "existing" if existing_alt else "missing"
    elif alt_text:
        row["alt_source"] = "manual"
    else:
        row["alt_source"] = "missing"

    session["summary"] = summarize_alt_rows(session.get("rows", []))
    return JSONResponse({"row": row, "summary": session["summary"]})


@app.post("/api/pdf-alt/session/{session_id}/clear-alt")
async def clear_pdf_alt_text_from_session(session_id: str):
    session = PDF_ALT_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="PDF ALT session not found.")

    rows = session.get("rows", [])
    cleared_count = 0
    for row in rows:
        previous_alt = normalize_alt_text(str(row.get("alt_text", "") or ""))
        if previous_alt:
            cleared_count += 1
        row["alt_text"] = ""
        row["generated_alt_text"] = ""
        row["effective_alt_text"] = ""
        row["alt_source"] = "missing"

    session["summary"] = summarize_alt_rows(rows)
    message = f"Cleared ALT text for {cleared_count} PDF item(s) in this session."
    return JSONResponse({"rows": rows, "summary": session["summary"], "message": message})


@app.post("/api/pdf-alt/session/{session_id}/replace-alt-text")
async def replace_pdf_alt_text_in_session(session_id: str, request: Request):
    session = PDF_ALT_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="PDF ALT session not found.")

    payload = await request.json()
    search_text = str(payload.get("search_text", "") or "")
    replace_text = str(payload.get("replace_text", "") or "")
    raw_item_ids = payload.get("item_ids", [])
    if not search_text:
        raise HTTPException(status_code=422, detail="Search text is required.")

    target_ids = {
        int(item_id)
        for item_id in raw_item_ids
        if isinstance(item_id, int) or (isinstance(item_id, str) and item_id.strip().isdigit())
    }
    if not target_ids:
        raise HTTPException(status_code=422, detail="Choose at least one PDF ALT item to update.")

    rows = session.get("rows", [])
    replaced_count = 0
    touched_count = 0
    for row in rows:
        row_id = row.get("id")
        if not isinstance(row_id, int) or row_id not in target_ids:
            continue
        current_alt = str(row.get("alt_text", "") or "")
        if search_text not in current_alt:
            continue
        updated_alt = current_alt.replace(search_text, replace_text)
        if updated_alt == current_alt:
            continue
        touched_count += 1
        replaced_count += current_alt.count(search_text)
        normalized_alt = normalize_alt_text(updated_alt)
        existing_alt = normalize_alt_text(str(row.get("existing_alt_text", "") or ""))
        row["alt_text"] = normalized_alt
        row["effective_alt_text"] = normalized_alt
        if normalized_alt == existing_alt:
            row["alt_source"] = "existing" if existing_alt else "missing"
        elif normalized_alt:
            row["alt_source"] = "manual"
        else:
            row["alt_source"] = "missing"

    session["summary"] = summarize_alt_rows(rows)
    message = f"Replaced {replaced_count} occurrence(s) across {touched_count} PDF item(s)."
    return JSONResponse({"rows": rows, "summary": session["summary"], "message": message})


@app.get("/api/pdf-alt/session/{session_id}/download.xlsx")
async def download_pdf_alt_inventory_excel(session_id: str):
    session = PDF_ALT_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="PDF ALT session not found.")

    if session.get("source_kind") != "pdf":
        raise HTTPException(status_code=400, detail="Excel download is available only for PDF ALT sessions.")

    workbook_bytes = build_alt_excel(
        session.get("rows", []),
        session.get("source_filename", "document.pdf"),
        preview_images=session.get("preview_images") or {},
    )
    filename = f"{safe_download_stem(session.get('source_filename', 'document.pdf'))}_pdf_alt_inventory.xlsx"
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return Response(
        content=workbook_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.post("/api/pdf-alt/session/{session_id}/import-excel")
async def import_pdf_alt_inventory_excel(session_id: str, workbook: UploadFile = File(...)):
    session = PDF_ALT_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="PDF ALT session not found.")

    if session.get("source_kind") != "pdf":
        raise HTTPException(status_code=400, detail="Excel import is available only for PDF ALT sessions.")

    try:
        workbook_bytes = await workbook.read()
        imported_alt_texts = parse_alt_injection_workbook(
            workbook_bytes,
            session.get("rows", []),
            session.get("source_filename"),
        )

        for index, row in enumerate(session.get("rows", [])):
            imported_alt = normalize_alt_text(imported_alt_texts[index] if index < len(imported_alt_texts) else "")
            existing_alt = normalize_alt_text(str(row.get("existing_alt_text", "") or ""))
            row["alt_text"] = imported_alt
            row["effective_alt_text"] = imported_alt
            if imported_alt == existing_alt:
                row["alt_source"] = "existing" if existing_alt else "missing"
            elif imported_alt:
                row["alt_source"] = "excel_import"
            else:
                row["alt_source"] = "missing"

        session["summary"] = summarize_alt_rows(session.get("rows", []))
        return JSONResponse(
            {
                "rows": session.get("rows", []),
                "summary": session["summary"],
                "message": "Imported PDF ALT text from Excel.",
            }
        )
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(exc)) from exc


def run_pdf_alt_generation(session: dict, generation_mode: str) -> dict:
    source_path = session.get("source_path") if isinstance(session.get("source_path"), Path) else None
    if not isinstance(source_path, Path):
        raise HTTPException(status_code=400, detail="PDF source file is no longer available for this session.")

    rows = session.get("rows", [])
    preview_images = session.setdefault("preview_images", {})
    if generation_mode == "claude":
        result = generate_missing_alt_rows_with_claude(rows, preview_images, source_path, None)
    elif generation_mode == "copilot":
        result = generate_missing_alt_rows_with_copilot(rows, preview_images, source_path, None)
    elif generation_mode == "gemini":
        result = generate_missing_alt_rows_with_gemini(rows, preview_images, source_path, None)
    elif generation_mode == "openrouter":
        result = generate_missing_alt_rows_with_openrouter(rows, preview_images, source_path, None)
    else:
        result = generate_missing_alt_rows_with_groq(rows, preview_images, source_path, None)

    if not result.get("available"):
        provider_label = generation_mode.replace("_", " ").title()
        raise HTTPException(status_code=422, detail=result.get("message", f"{provider_label} PDF ALT generation is unavailable."))
    return result


@app.post("/api/pdf-alt/session/{session_id}/generate-alt")
async def generate_pdf_alt_text(session_id: str, request: Request):
    session = PDF_ALT_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="PDF ALT session not found.")

    try:
        payload = {}
        try:
            payload = await request.json()
        except Exception:
            payload = {}

        generation_mode = normalize_alt_text(str((payload or {}).get("provider", "") or "")).lower()
        if generation_mode not in {"claude", "copilot", "gemini", "groq", "openrouter", "claude_fallback_groq"}:
            generation_mode = "groq"

        if generation_mode == "claude_fallback_groq":
            try:
                claude_result = run_pdf_alt_generation(session, "claude")
            except HTTPException as claude_error:
                groq_result = run_pdf_alt_generation(session, "groq")
                session["rows"] = groq_result.get("rows", session.get("rows", []))
                session["summary"] = groq_result.get("summary", summarize_alt_rows(session.get("rows", [])))
                message = f"{claude_error.detail} Fallback: {groq_result.get('message', 'Generated PDF ALT text with Groq.')}"
                provider = "claude_fallback_groq"
                generated_count = int(groq_result.get("generated_count", 0) or 0)
            else:
                rows_after_claude = claude_result.get("rows", session.get("rows", []))
                session["rows"] = rows_after_claude
                session["summary"] = claude_result.get("summary", summarize_alt_rows(rows_after_claude))
                groq_result = run_pdf_alt_generation(session, "groq")
                session["rows"] = groq_result.get("rows", rows_after_claude)
                session["summary"] = summarize_alt_rows(session.get("rows", []))
                claude_count = int(claude_result.get("generated_count", 0) or 0)
                groq_count = int(groq_result.get("generated_count", 0) or 0)
                generated_count = claude_count + groq_count
                provider = "claude_fallback_groq" if groq_count else "claude"
                message = claude_result.get("message", "Generated PDF ALT text with Claude.")
                if groq_count:
                    message += f" Fallback Groq generated {groq_count} additional item(s)."
        else:
            result = run_pdf_alt_generation(session, generation_mode)
            session["rows"] = result.get("rows", session.get("rows", []))
            session["summary"] = result.get("summary", summarize_alt_rows(session.get("rows", [])))
            message = result.get("message", "Generated PDF ALT text.")
            provider = result.get("provider", generation_mode)
            generated_count = int(result.get("generated_count", 0) or 0)

        return JSONResponse(
            {
                "rows": session.get("rows", []),
                "summary": session["summary"],
                "message": message,
                "generated_count": generated_count,
                "provider": provider,
            }
        )
    except HTTPException:
        raise
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.get("/api/pdf-alt/session/{session_id}/updated.pdf")
async def download_updated_pdf(session_id: str):
    session = PDF_ALT_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="PDF ALT session not found.")

    source_path = session.get("source_path")
    if not isinstance(source_path, Path) or source_path.suffix.lower() != ".pdf":
        raise HTTPException(status_code=400, detail="Updated PDF download is available only for PDF ALT sessions.")

    try:
        pdf_bytes, _applied_count = inject_pdf_alt_texts(source_path, session.get("rows", []))
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    filename = f"{safe_download_stem(session.get('source_filename', source_path.name))}_updated_alt.pdf"
    headers = {"Content-Disposition": f'attachment; filename="{filename}"; filename*=UTF-8\'\'{quote(filename)}'}
    return Response(content=pdf_bytes, media_type="application/pdf", headers=headers)


@app.post("/api/list-correction/process")
async def process_list_correction_document(
    file: UploadFile = File(...),
    include_hf: bool = Form(False),
):
    filename = Path(file.filename or "upload.bin").name
    suffix = Path(filename).suffix.lower()
    if suffix != ".docx":
        raise HTTPException(status_code=400, detail="List Correction supports DOCX files only.")

    session_id = uuid.uuid4().hex
    session_dir = TEMP_DIR / f"list_correction_{session_id}"
    session_dir.mkdir(parents=True, exist_ok=True)
    source_path = write_upload(session_dir, file)
    validate_uploaded_docx(source_path)
    output_filename = build_list_correction_output_name(filename)
    output_path = session_dir / output_filename

    try:
        processor = ListProcessor()
        stats = processor.process_document(source_path, output_path, include_hf=include_hf)
        stats_payload = stats.to_dict()
        LIST_CORRECTION_SESSIONS[session_id] = {
            "source_path": source_path,
            "output_path": output_path,
            "output_filename": output_filename,
            "include_hf": include_hf,
            "stats": stats_payload,
        }
        return JSONResponse(
            {
                "session_id": session_id,
                "output_filename": output_filename,
                "stats": stats_payload,
                "message": f"Corrected lists in {filename} and prepared a normalized DOCX copy.",
            }
        )
    except HTTPException:
        raise
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.get("/api/list-correction/session/{session_id}/download.docx")
async def download_list_correction_docx(session_id: str):
    session = LIST_CORRECTION_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="List Correction session not found.")

    output_path = session.get("output_path")
    if not isinstance(output_path, Path) or not output_path.exists():
        raise HTTPException(status_code=404, detail="Corrected DOCX not found.")

    filename = str(session.get("output_filename") or output_path.name)
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return Response(
        content=output_path.read_bytes(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers,
    )


@app.post("/api/color-correction/process")
async def process_color_correction_document(
    file: UploadFile = File(...),
    background: str = Form("#FFFFFF"),
):
    filename = Path(file.filename or "upload.bin").name
    suffix = Path(filename).suffix.lower()
    if suffix != ".docx":
        raise HTTPException(status_code=400, detail="Color Correction supports DOCX files only.")

    try:
        normalized_background = normalize_hex_color(background)
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc

    session_id = uuid.uuid4().hex
    session_dir = TEMP_DIR / f"color_correction_{session_id}"
    session_dir.mkdir(parents=True, exist_ok=True)
    source_path = write_upload(session_dir, file)
    validate_uploaded_docx(source_path)
    output_filename = build_color_correction_output_name(filename)
    output_path = session_dir / output_filename

    try:
        processed = process_docx_bytes(filename, source_path.read_bytes(), background=normalized_background)
        output_path.write_bytes(processed.output_bytes)
        result_payload = {
            "fixed_elements": processed.fixed_elements,
            "changed": processed.changed,
            "background": normalized_background,
        }
        COLOR_CORRECTION_SESSIONS[session_id] = {
            "source_path": source_path,
            "output_path": output_path,
            "output_filename": output_filename,
            "background": normalized_background,
            "result": result_payload,
        }
        return JSONResponse(
            {
                "session_id": session_id,
                "output_filename": output_filename,
                "result": result_payload,
                "message": f"Processed {filename} and checked document colors against {normalized_background}.",
            }
        )
    except HTTPException:
        raise
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.get("/api/color-correction/session/{session_id}/download.docx")
async def download_color_correction_docx(session_id: str):
    session = COLOR_CORRECTION_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Color Correction session not found.")

    output_path = session.get("output_path")
    if not isinstance(output_path, Path) or not output_path.exists():
        raise HTTPException(status_code=404, detail="Corrected DOCX not found.")

    filename = str(session.get("output_filename") or output_path.name)
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return Response(
        content=output_path.read_bytes(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers,
    )


@app.post("/api/excel-merger/process")
async def process_excel_merger(files: list[UploadFile] = File(...)):
    if not files:
        raise HTTPException(status_code=400, detail="Excel Merger requires at least one workbook.")

    session_id = uuid.uuid4().hex
    session_dir = TEMP_DIR / f"excel_merger_{session_id}"
    session_dir.mkdir(parents=True, exist_ok=True)

    file_payloads: list[tuple[str, bytes]] = []
    filenames: list[str] = []
    for upload in files:
        filename = Path(upload.filename or "workbook.xlsx").name
        if not supported_excel_filename(filename):
            raise HTTPException(status_code=400, detail=f"{filename} is not a supported Excel workbook.")
        workbook_bytes = await upload.read()
        if not workbook_bytes:
            raise HTTPException(status_code=422, detail=f"{filename} is empty.")
        file_payloads.append((filename, workbook_bytes))
        filenames.append(filename)

    try:
        merged_bytes, summary = merge_excel_workbooks(file_payloads)
        output_filename = build_excel_merge_output_name(filenames[0] if filenames else "workbook.xlsx")
        output_path = session_dir / output_filename
        output_path.write_bytes(merged_bytes)
        EXCEL_MERGER_SESSIONS[session_id] = {
            "output_path": output_path,
            "output_filename": output_filename,
            "source_filenames": filenames,
            "summary": summary,
        }
        return JSONResponse(
            {
                "session_id": session_id,
                "output_filename": output_filename,
                "summary": summary,
                "message": f"Merged {summary.get('files', 0)} workbook(s) into {summary.get('sheets', 0)} sheet(s).",
            }
        )
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.get("/api/excel-merger/session/{session_id}/download.xlsx")
async def download_excel_merged_workbook(session_id: str):
    session = EXCEL_MERGER_SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Excel Merger session not found.")

    output_path = session.get("output_path")
    if not isinstance(output_path, Path) or not output_path.exists():
        raise HTTPException(status_code=404, detail="Merged Excel workbook not found.")

    filename = str(session.get("output_filename") or output_path.name)
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return Response(
        content=output_path.read_bytes(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )
