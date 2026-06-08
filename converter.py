import os
import shutil
import subprocess
import hashlib
from pathlib import Path


WORD_EXPORT_FORMAT_PDF = 17
WORD_EXPORT_OPTIMIZE_FOR_PRINT = 0
WORD_EXPORT_ALL_DOCUMENT = 0
WORD_EXPORT_DOCUMENT_CONTENT = 0
WORD_EXPORT_CREATE_NO_BOOKMARKS = 0


def _is_reusable_pdf(source: Path, pdf_path: Path) -> bool:
    if not pdf_path.exists() or pdf_path.stat().st_size <= 0:
        return False
    try:
        return pdf_path.stat().st_mtime >= source.stat().st_mtime
    except OSError:
        return True


def _file_sha1(path: Path) -> str:
    digest = hashlib.sha1()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _find_temp_cache_root() -> Path:
    return Path(__file__).resolve().parent.parent / "temp"


def _find_cached_pdf_for_docx(source: Path, pdf_name: str) -> Path | None:
    temp_root = _find_temp_cache_root()
    if not temp_root.exists():
        return None

    source_size = source.stat().st_size
    source_hash = None
    candidates = sorted(
        temp_root.rglob(pdf_name),
        key=lambda candidate: candidate.stat().st_mtime if candidate.exists() else 0.0,
        reverse=True,
    )

    for cached_pdf in candidates:
        if cached_pdf.resolve() == source.with_suffix(".pdf"):
            continue
        cached_docx = cached_pdf.with_suffix(".docx")
        if not cached_docx.exists():
            continue
        try:
            if cached_docx.stat().st_size != source_size:
                continue
        except OSError:
            continue

        if source_hash is None:
            source_hash = _file_sha1(source)

        try:
            if _file_sha1(cached_docx) != source_hash:
                continue
        except OSError:
            continue

        if _is_reusable_pdf(cached_docx, cached_pdf):
            return cached_pdf

    return None


def _format_subprocess_error(prefix: str, exc: Exception) -> str:
    if not isinstance(exc, subprocess.CalledProcessError):
        return f"{prefix}: {exc}"

    parts = [f"{prefix}: exit code {exc.returncode}"]
    stdout = (exc.stdout or "").strip()
    stderr = (exc.stderr or "").strip()
    if stderr:
        parts.append(f"stderr: {stderr}")
    if stdout:
        parts.append(f"stdout: {stdout}")
    return " ".join(parts)


def _convert_with_word(docx_path: Path, pdf_path: Path) -> None:
    # Use Microsoft Word's renderer on Windows so pagination/layout match the source DOCX.
    # Prefer ExportAsFixedFormat3 for better image-quality control, then fall back.
    command = rf"""
$ErrorActionPreference = 'Stop'
$docx = [System.IO.Path]::GetFullPath('{str(docx_path).replace("'", "''")}')
$pdf = [System.IO.Path]::GetFullPath('{str(pdf_path).replace("'", "''")}')
$word = $null
$document = $null
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($docx, $false, $true)
    if ($document -and $document.PSObject.Methods.Name -contains 'ExportAsFixedFormat3') {{
        $document.ExportAsFixedFormat3(
            $pdf,
            {WORD_EXPORT_FORMAT_PDF},
            $false,
            {WORD_EXPORT_OPTIMIZE_FOR_PRINT},
            {WORD_EXPORT_ALL_DOCUMENT},
            1,
            1,
            {WORD_EXPORT_DOCUMENT_CONTENT},
            $true,
            $true,
            {WORD_EXPORT_CREATE_NO_BOOKMARKS},
            $true,
            $false,
            $true,
            $true
        )
    }}
    else {{
        $document.ExportAsFixedFormat(
            $pdf,
            {WORD_EXPORT_FORMAT_PDF},
            $false,
            {WORD_EXPORT_OPTIMIZE_FOR_PRINT},
            {WORD_EXPORT_ALL_DOCUMENT},
            1,
            1,
            {WORD_EXPORT_DOCUMENT_CONTENT},
            $true,
            $true,
            {WORD_EXPORT_CREATE_NO_BOOKMARKS},
            $true,
            $false
        )
    }}
}}
finally {{
    if ($document -ne $null) {{
        $document.Close($false)
    }}
    if ($word -ne $null) {{
        $word.Quit()
    }}
}}
"""

    try:
        subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", command],
            check=True,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except subprocess.CalledProcessError as exc:
        raise RuntimeError(_format_subprocess_error("Microsoft Word export failed", exc)) from exc


def _convert_with_libreoffice(docx_path: Path, output_dir: Path) -> None:
    soffice_path = Path(r"C:\Program Files\LibreOffice\program\soffice.exe")
    if not soffice_path.exists():
        raise FileNotFoundError("LibreOffice was not found at the expected path.")

    try:
        subprocess.run(
            [
                str(soffice_path),
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(output_dir),
                str(docx_path),
            ],
            check=True,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except subprocess.CalledProcessError as exc:
        raise RuntimeError(_format_subprocess_error("LibreOffice export failed", exc)) from exc


def docx_to_pdf(docx_path: str, output_pdf_path: str | None = None) -> str:
    """
    Convert DOCX to PDF while preserving the original layout as strictly as possible.

    Default behavior:
    - Use Microsoft Word automation for native DOCX rendering.
    - Refuse to silently fall back to LibreOffice unless explicitly enabled.
    """
    source = Path(docx_path).resolve()
    pdf_path = Path(output_pdf_path).resolve() if output_pdf_path else source.with_suffix(".pdf")
    output_dir = pdf_path.parent
    reuse_existing = os.getenv("MATCHA_REUSE_EXISTING_PDF", "1").strip() != "0"
    reuse_cached = os.getenv("MATCHA_REUSE_TEMP_CACHE_PDF", "1").strip() != "0"

    if reuse_existing and _is_reusable_pdf(source, pdf_path):
        return str(pdf_path)

    if reuse_cached:
        cached_pdf = _find_cached_pdf_for_docx(source, pdf_path.name)
        if cached_pdf is not None:
            if cached_pdf.resolve() != pdf_path.resolve():
                pdf_path.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(cached_pdf, pdf_path)
            if _is_reusable_pdf(source, pdf_path):
                return str(pdf_path)

    renderer = os.getenv("MATCHA_DOCX_RENDERER", "word").strip().lower()
    allow_fallback = os.getenv("MATCHA_ALLOW_LIBREOFFICE_FALLBACK", "0").strip() == "1"

    errors = []

    if renderer in {"word", "auto"}:
        try:
            _convert_with_word(source, pdf_path)
        except Exception as exc:
            errors.append(str(exc))
            if reuse_existing and _is_reusable_pdf(source, pdf_path):
                return str(pdf_path)
        else:
            if pdf_path.exists():
                return str(pdf_path)

    if renderer in {"libreoffice", "auto"} or allow_fallback:
        try:
            _convert_with_libreoffice(source, output_dir)
        except Exception as exc:
            errors.append(str(exc))
            if reuse_existing and _is_reusable_pdf(source, pdf_path):
                return str(pdf_path)
        else:
            if pdf_path.exists():
                return str(pdf_path)

    if renderer == "word" and not allow_fallback:
        raise RuntimeError(
            "Strict DOCX rendering requires Microsoft Word automation. "
            "LibreOffice fallback is disabled because it can change layout. "
            + (" ".join(errors) if errors else "")
        )

    raise RuntimeError(
        "PDF conversion failed. "
        + (" ".join(errors) if errors else "No renderer produced an output PDF.")
    )
