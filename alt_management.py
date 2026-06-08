from __future__ import annotations

import base64
import io
import json
import os
import re
import shutil
from functools import lru_cache
import textwrap
from pathlib import Path
import subprocess
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

import fitz
import requests
from PIL import Image, ImageChops

from Altomizer.style_inspector import extract_doc_structure


WORD_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "o": "urn:schemas-microsoft-com:office:office",
}


def is_image_only_mode() -> bool:
    raw = (os.getenv("MATCHA_ALT_IMAGE_ONLY_MODE") or "").strip().lower()
    if not raw:
        return False
    return raw in {"1", "true", "yes", "on"}


def build_alt_inventory(docx_path: Path, pdf_path: Path | None = None) -> dict:
    include_equations = not is_image_only_mode()
    direct_inventory = build_alt_inventory_from_docx(docx_path, include_equations=include_equations)

    if direct_inventory.get("available"):
        return direct_inventory

    return {
        "available": False,
        "message": direct_inventory.get("message", "DOCX ALT inventory could not be generated."),
        "rows": [],
        "summary": {
            "total_items": 0,
            "images": 0,
            "equations": 0,
            "with_alt_text": 0,
            "without_alt_text": 0,
        },
    }


def build_alt_inventory_from_docx(
    docx_path: Path,
    include_equations: bool = False,
) -> dict:
    try:
        elements = extract_doc_structure(str(docx_path))
    except Exception as exc:
        return {
            "available": False,
            "message": f"ALT inventory could not be generated: {exc}",
            "rows": [],
            "summary": {
                "total_items": 0,
                "images": 0,
                "equations": 0,
                "with_alt_text": 0,
                "without_alt_text": 0,
            },
        }

    rows = []
    for element in elements:
        role = str(getattr(element, "role", "") or "").strip().lower()
        if role not in {"image", "equation"}:
            continue
        if role == "equation" and not include_equations:
            continue

        metadata = getattr(element, "metadata", None) or {}
        label_text = normalize_alt_text(getattr(element, "text", "") or "")

        entry = {
            "role": role,
            "text": label_text,
            "metadata": metadata,
            "has_alt_text": getattr(element, "has_alt_text", False),
        }
        if not is_real_visual_inventory_item(entry):
            continue

        has_alt_text = bool(getattr(element, "has_alt_text", False)) and not is_effectively_empty_alt_text(label_text)
        alt_text = label_text if has_alt_text else ""
        media_target = metadata.get("image_target") or metadata.get("target") or metadata.get("ole_target")
        if isinstance(media_target, str):
            media_target = media_target.strip() or None
        ole_target = metadata.get("ole_target")
        if isinstance(ole_target, str):
            ole_target = ole_target.strip() or None
        source_part = getattr(element, "source_part", "body") or "body"

        rows.append(
            {
                "id": len(rows),
                "type": "Image" if role == "image" else "Equation",
                "role": role,
                "source_part": source_part,
                "page": None,
                "preview_page": None,
                "preview_bbox": None,
                "preview_text": "",
                "has_alt_text": has_alt_text,
                "alt_text": alt_text,
                "existing_alt_text": alt_text,
                "generated_alt_text": "",
                "alt_source": "existing" if has_alt_text else "missing",
                "label": label_text or ("Image" if role == "image" else "Equation"),
                "media_target": media_target,
                "ole_target": ole_target,
                "rel_id": metadata.get("rel_id") or metadata.get("image_rid") or metadata.get("ole_rid"),
                "docpr_id": metadata.get("docpr_id") or metadata.get("shape_id"),
                "docpr_name": metadata.get("docpr_name", ""),
                "display_width_pt": metadata.get("width_pt"),
                "display_height_pt": metadata.get("height_pt"),
                "viewport_crop": metadata.get("crop"),
                "source_file": docx_path.name,
            }
        )

    return {
        "available": True,
        "message": "",
        "rows": rows,
        "summary": summarize_alt_rows(rows),
    }


def is_real_visual_inventory_item(entry: dict) -> bool:
    role = str(entry.get("role", "")).strip().lower()
    metadata = entry.get("metadata") or {}
    label = normalize_alt_text(entry.get("text", ""))
    has_alt_text = bool(entry.get("has_alt_text")) and not is_effectively_empty_alt_text(label)

    if looks_like_nonvisual_heading(label):
        return False

    def is_visual_target(target: object) -> bool:
        if not isinstance(target, str):
            return False
        normalized = target.strip().replace("\\", "/").lstrip("/").lower()
        return normalized.startswith("word/media/") or normalized.startswith("word/embeddings/")

    def has_generic_visual_signal() -> bool:
        target = metadata.get("target")
        rel_id = metadata.get("rel_id")
        docpr_name = normalize_alt_text(metadata.get("docpr_name", ""))
        width_pt = metadata.get("width_pt")
        height_pt = metadata.get("height_pt")

        if isinstance(target, str) and target.strip():
            normalized = target.strip().replace("\\", "/").lstrip("/").lower()
            if normalized.startswith("word/"):
                return True
        if isinstance(rel_id, str) and rel_id.strip():
            return True
        if docpr_name:
            return True
        if isinstance(width_pt, (int, float)) and width_pt > 0:
            return True
        if isinstance(height_pt, (int, float)) and height_pt > 0:
            return True
        return False

    def target_extension(target: object) -> str:
        if not isinstance(target, str):
            return ""
        return Path(target).suffix.lower().lstrip(".")

    def looks_like_empty_or_placeholder_label() -> bool:
        lowered = label.lower()
        if is_effectively_empty_alt_text(label):
            return True
        return lowered in {"image", "equation"}

    def is_probable_fallback_artifact() -> bool:
        if role != "image":
            return False
        target = metadata.get("target") or metadata.get("image_target")
        ext = target_extension(target)
        docpr_name = normalize_alt_text(metadata.get("docpr_name", ""))
        width_pt = metadata.get("width_pt")
        height_pt = metadata.get("height_pt")

        if has_alt_text:
            return False
        if ext == "emf" and not docpr_name:
            return True
        if looks_like_empty_or_placeholder_label() and not docpr_name:
            width_known = isinstance(width_pt, (int, float)) and width_pt > 0
            height_known = isinstance(height_pt, (int, float)) and height_pt > 0
            if not width_known or not height_known:
                return True
            if width_pt <= 2 and height_pt <= 2:
                return True
        return False

    if is_probable_fallback_artifact():
        return False

    if role == "image":
        if looks_like_empty_or_placeholder_label() and not has_alt_text and not has_generic_visual_signal():
            return False
        return (
            is_visual_target(metadata.get("target"))
            or is_visual_target(metadata.get("image_target"))
            or is_visual_target(metadata.get("ole_target"))
            or has_generic_visual_signal()
        )

    if role == "equation":
        if metadata.get("math_source") == "omml":
            return False
        if metadata.get("prog_id"):
            return True
        return (
            is_visual_target(metadata.get("image_target"))
            or is_visual_target(metadata.get("ole_target"))
            or has_generic_visual_signal()
        )

    return False


def looks_like_nonvisual_heading(text: str) -> bool:
    normalized = normalize_alt_text(text).lower()
    if not normalized:
        return False
    if re.match(r"^\d+(?:[-.]\d+)+(?:\s|$)", normalized):
        if re.search(r"\b(answers?|chapter|section|exercise|problem|review|solutions?|cont\.?)\b", normalized):
            return True
    if re.search(r"\b(answers?|chapter|section|exercise|problem|review|solutions?|cont\.?)\b", normalized):
        if re.match(r"^\d+(?:[-.]\d+)?\b", normalized):
            return True
    return False


def is_effectively_empty_alt_text(text: str) -> bool:
    normalized = normalize_alt_text(text)
    if not normalized:
        return True
    stripped = normalized.replace('"', "").replace("'", "").strip()
    return not stripped


def summarize_alt_rows(rows: list[dict]) -> dict:
    total_items = len(rows)
    images = sum(1 for row in rows if str(row.get("role", "")).lower() == "image")
    equations = sum(1 for row in rows if str(row.get("role", "")).lower() == "equation")
    with_original_alt = sum(1 for row in rows if bool(row.get("has_alt_text")))
    generated = sum(1 for row in rows if bool((row.get("generated_alt_text") or "").strip()))
    with_effective_alt = sum(1 for row in rows if bool((row.get("alt_text") or "").strip()))
    without_alt_text = max(0, total_items - with_effective_alt)
    return {
        "total_items": total_items,
        "images": images,
        "equations": equations,
        "with_alt_text": with_original_alt,
        "generated_alt_text": generated,
        "with_effective_alt_text": with_effective_alt,
        "without_alt_text": without_alt_text,
    }


@lru_cache(maxsize=1)
def load_alt_style_bank() -> tuple[dict, ...]:
    workbook_path = discover_alt_style_workbook_path()
    if workbook_path is None:
        return tuple()
    try:
        return tuple(parse_alt_style_examples_from_xlsx(workbook_path))
    except Exception:
        return tuple()


def discover_alt_style_workbook_path() -> Path | None:
    env_path = (os.getenv("MATCHA_ALT_EXAMPLE_XLSX_PATH") or "").strip()
    if env_path:
        candidate = Path(env_path)
        if candidate.exists():
            return candidate

    downloads = Path.home() / "Downloads"
    if downloads.exists():
        candidates = sorted(
            (path for path in downloads.glob("*_alt_inventory*.xlsx") if not path.name.startswith("~$")),
            key=lambda item: item.stat().st_mtime if item.exists() else 0,
            reverse=True,
        )
        if candidates:
            return candidates[0]
    return None


def parse_alt_style_examples_from_xlsx(workbook_path: Path) -> list[dict]:
    if not workbook_path.exists():
        return []

    ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    examples: list[dict] = []

    with zipfile.ZipFile(workbook_path, "r") as archive:
        sheet_xml = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        shared_strings = read_shared_strings(archive)
        rows = sheet_xml.findall(".//a:sheetData/a:row", ns)
        if not rows:
            return []

        headers = workbook_row_to_cells(rows[0], ns, shared_strings)
        header_map = {normalize_alt_text(value).lower(): col_index for col_index, value in headers.items()}

        for row in rows[1:]:
            values = workbook_row_to_cells(row, ns, shared_strings)
            if not any(values.values()):
                continue

            role = value_from_headers(values, header_map, ("type",))
            if normalize_alt_text(role).lower() != "equation":
                continue

            alt_text = value_from_headers(values, header_map, ("alt text", "detected label", "label"))
            ocr_formula = value_from_headers(values, header_map, ("ocr formula",))
            generated_alt = value_from_headers(values, header_map, ("generated alt text",))
            source_text = ocr_formula or alt_text or generated_alt
            final_alt = alt_text or generated_alt or source_text

            final_alt = normalize_alt_text(final_alt)
            source_text = normalize_alt_text(source_text)
            if not final_alt:
                continue
            if not source_text:
                source_text = final_alt

            examples.append(
                {
                    "source_text": source_text,
                    "alt_text": final_alt,
                    "page": value_from_headers(values, header_map, ("page",)),
                    "source_file": value_from_headers(values, header_map, ("source file",)),
                }
            )

    deduped: list[dict] = []
    seen = set()
    for example in examples:
        key = normalize_alt_text(example.get("alt_text", "")).lower()
        if not key or key in seen:
            continue
        seen.add(key)
        deduped.append(example)
    return [example for example in deduped if is_good_alt_style_example(example.get("alt_text", ""))]


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        shared_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []

    shared_strings: list[str] = []
    for item in shared_root.findall(".//a:si", {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}):
        texts = [node.text or "" for node in item.findall(".//a:t", {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"})]
        shared_strings.append("".join(texts))
    return shared_strings


def workbook_row_to_cells(row: ET.Element, ns: dict, shared_strings: list[str]) -> dict[int, str]:
    values: dict[int, str] = {}
    for cell in row.findall("a:c", ns):
        ref = cell.attrib.get("r", "")
        col_index = excel_column_index(ref)
        if col_index < 0:
            continue
        values[col_index] = read_workbook_cell_value(cell, ns, shared_strings)
    return values


def read_workbook_cell_value(cell: ET.Element, ns: dict, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        texts = [node.text or "" for node in cell.findall(".//a:t", ns)]
        return "".join(texts)
    value = cell.find("a:v", ns)
    if value is None or value.text is None:
        return ""
    if cell_type == "s":
        try:
            return shared_strings[int(value.text)]
        except (ValueError, IndexError):
            return ""
    return value.text


def value_from_headers(values: dict[int, str], header_map: dict[str, int], candidates: tuple[str, ...]) -> str:
    for candidate in candidates:
        index = header_map.get(candidate)
        if index is None:
            continue
        if index in values:
            return values[index]
    return ""


def excel_column_index(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref or "", flags=re.IGNORECASE)
    if not match:
        return -1
    column = match.group(1).upper()
    total = 0
    for char in column:
        if not char.isalpha():
            return -1
        total = total * 26 + (ord(char) - ord("A") + 1)
    return total - 1


def select_alt_style_examples(candidate_text: str, source_filename: str | None = None, limit: int = 3) -> list[dict]:
    bank = list(load_alt_style_bank())
    if not bank:
        return []

    normalized_source_filename = normalize_alt_text(source_filename or "").lower()
    if not normalized_source_filename:
        return []

    bank = [
        example
        for example in bank
        if normalize_alt_text(example.get("source_file", "")).lower() == normalized_source_filename
    ]
    if not bank:
        return []

    candidate_tokens = equation_similarity_tokens(candidate_text)
    scored: list[tuple[float, dict]] = []
    for example in bank:
        example_tokens = equation_similarity_tokens(example.get("source_text") or example.get("alt_text") or "")
        overlap = token_jaccard(candidate_tokens, example_tokens)
        length_bonus = min(len(example_tokens), 20) / 100.0
        scored.append((overlap + length_bonus, example))

    scored.sort(key=lambda item: item[0], reverse=True)
    return [item[1] for item in scored[:limit]]


def equation_similarity_tokens(text: str) -> set[str]:
    normalized = normalize_alt_text(text).lower()
    if not normalized:
        return set()
    tokens = re.findall(r"[a-z0-9]+", normalized)
    return set(tokens)


def token_jaccard(left: set[str], right: set[str]) -> float:
    if not left or not right:
        return 0.0
    return len(left & right) / max(1, len(left | right))


def is_good_alt_style_example(text: str) -> bool:
    cleaned = normalize_alt_text(text).lower()
    if not cleaned:
        return False
    if len(cleaned.split()) < 2 or len(cleaned.split()) > 18:
        return False
    if any(marker in cleaned for marker in ("bizchat", "feedback", "page", "provide your feedback", "equation showing", "expression showing")):
        return False
    if any(marker in cleaned for marker in ("start fraction", "end fraction", "numerator", "denominator")):
        return False
    if any(marker in cleaned for marker in ("equals", "over", "plus", "minus", "open parenthesis", "close parenthesis", "squared", "cubed")):
        return True
    return False


def get_math_ocr_config() -> dict | None:
    provider = (os.getenv("MATCHA_ALT_MATH_OCR_PROVIDER") or "azure").strip().lower()
    if provider not in {"azure", "auto"}:
        return None

    endpoint = (
        os.getenv("MATCHA_AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT")
        or os.getenv("AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT")
        or ""
    ).strip().rstrip("/")
    key = (
        os.getenv("MATCHA_AZURE_DOCUMENT_INTELLIGENCE_KEY")
        or os.getenv("AZURE_DOCUMENT_INTELLIGENCE_KEY")
        or ""
    ).strip()
    api_version = (os.getenv("MATCHA_AZURE_DOCUMENT_INTELLIGENCE_API_VERSION") or "2024-11-30").strip()
    model_id = (os.getenv("MATCHA_AZURE_DOCUMENT_INTELLIGENCE_MODEL") or "prebuilt-layout").strip()

    if not endpoint or not key:
        return None

    return {
        "provider": "azure",
        "endpoint": endpoint,
        "key": key,
        "api_version": api_version,
        "model_id": model_id or "prebuilt-layout",
        "timeout_seconds": int(os.getenv("MATCHA_AZURE_DOCUMENT_INTELLIGENCE_TIMEOUT", "180") or "180"),
        "poll_interval_seconds": float(os.getenv("MATCHA_AZURE_DOCUMENT_INTELLIGENCE_POLL_INTERVAL", "1.5") or "1.5"),
        "max_polls": int(os.getenv("MATCHA_AZURE_DOCUMENT_INTELLIGENCE_MAX_POLLS", "120") or "120"),
    }


def collect_cloud_formula_items(source_path: Path) -> dict:
    config = get_math_ocr_config()
    if not config:
        return {
            "available": False,
            "provider": "azure",
            "message": (
                "Cloud formula OCR is not configured. Set AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT and "
                "AZURE_DOCUMENT_INTELLIGENCE_KEY to enable math-specific OCR."
            ),
            "items": [],
        }

    try:
        items = fetch_azure_formula_items(source_path, config)
    except Exception as exc:
        return {
            "available": False,
            "provider": "azure",
            "message": f"Cloud formula OCR failed: {summarize_automation_error(str(exc))}",
            "items": [],
        }

    return {
        "available": True,
        "provider": "azure",
        "message": f"Cloud formula OCR detected {len(items)} formula item(s).",
        "items": items,
    }


def fetch_azure_formula_items(source_path: Path, config: dict) -> list[dict]:
    endpoint = str(config.get("endpoint", "")).rstrip("/")
    key = str(config.get("key", "")).strip()
    model_id = str(config.get("model_id", "prebuilt-layout")).strip() or "prebuilt-layout"
    api_version = str(config.get("api_version", "2024-11-30")).strip() or "2024-11-30"
    timeout_seconds = int(config.get("timeout_seconds", 180))
    poll_interval_seconds = float(config.get("poll_interval_seconds", 1.5))
    max_polls = int(config.get("max_polls", 120))

    analyze_url = (
        f"{endpoint}/documentintelligence/documentModels/{model_id}:analyze"
        f"?api-version={api_version}&features=formulas"
    )

    content_type = docx_content_type(source_path.suffix.lower())
    headers = {
        "Ocp-Apim-Subscription-Key": key,
        "Content-Type": content_type,
    }

    with source_path.open("rb") as handle:
        response = requests.post(
            analyze_url,
            headers=headers,
            data=handle,
            timeout=timeout_seconds,
        )

    if response.status_code not in {202, 200}:
        raise RuntimeError(_azure_error_message(response))

    operation_location = response.headers.get("Operation-Location") or response.headers.get("operation-location")
    if operation_location:
        result_payload = poll_azure_operation(operation_location, key, timeout_seconds, poll_interval_seconds, max_polls)
    else:
        body = _safe_json(response)
        if isinstance(body, dict) and body.get("analyzeResult"):
            result_payload = body
        else:
            if isinstance(body, dict):
                raise RuntimeError(
                    body.get("error", {}).get("message", "Azure Document Intelligence did not return an operation location.")
                )
            raise RuntimeError("Azure Document Intelligence did not return an operation location.")

    analyze_result = result_payload.get("analyzeResult") or {}
    pages = analyze_result.get("pages") or []
    formulas = []

    for page in pages:
        try:
            page_number = int(page.get("pageNumber"))
        except (TypeError, ValueError):
            continue
        width = float(page.get("width") or 0.0)
        height = float(page.get("height") or 0.0)
        for formula in page.get("formulas") or []:
            latex = normalize_alt_text(formula.get("value", ""))
            if not latex:
                continue
            bbox_norm = polygon_to_bbox_norm(formula.get("polygon"), width, height)
            formulas.append(
                {
                    "page": page_number - 1,
                    "page_number": page_number,
                    "kind": str(formula.get("kind") or "display").strip().lower(),
                    "confidence": formula.get("confidence"),
                    "latex": latex,
                    "bbox_norm": bbox_norm,
                    "center": bbox_center(bbox_norm),
                    "raw": formula,
                }
            )

    return formulas


def poll_azure_operation(operation_location: str, key: str, timeout_seconds: int, poll_interval_seconds: float, max_polls: int) -> dict:
    headers = {"Ocp-Apim-Subscription-Key": key}
    deadline = time.time() + max(5, timeout_seconds)
    last_payload: dict = {}

    for _ in range(max(1, max_polls)):
        if time.time() > deadline:
            raise RuntimeError("Azure Document Intelligence request timed out while waiting for formula OCR results.")

        response = requests.get(operation_location, headers=headers, timeout=max(10, timeout_seconds))
        if response.status_code != 200:
            raise RuntimeError(_azure_error_message(response))

        payload = _safe_json(response)
        if not isinstance(payload, dict):
            raise RuntimeError("Azure Document Intelligence returned an invalid result payload.")

        last_payload = payload
        status = str(payload.get("status") or "").lower()
        if status in {"succeeded", "completed"}:
            return payload
        if status in {"failed", "canceled"}:
            error = payload.get("error") or {}
            message = error.get("message") if isinstance(error, dict) else None
            raise RuntimeError(message or "Azure Document Intelligence formula OCR failed.")

        time.sleep(max(0.5, poll_interval_seconds))

    raise RuntimeError(
        last_payload.get("error", {}).get("message", "Azure Document Intelligence formula OCR polling exceeded the allowed attempts.")
        if isinstance(last_payload.get("error"), dict)
        else "Azure Document Intelligence formula OCR polling exceeded the allowed attempts."
    )


def _safe_json(response: requests.Response) -> dict | list | None:
    try:
        return response.json()
    except ValueError:
        return None


def _azure_error_message(response: requests.Response) -> str:
    payload = _safe_json(response)
    if isinstance(payload, dict):
        error = payload.get("error") or {}
        if isinstance(error, dict):
            message = error.get("message") or error.get("code")
            if message:
                return str(message)
    return f"Azure Document Intelligence request failed with status {response.status_code}."


def docx_content_type(suffix: str) -> str:
    suffix = suffix.lower()
    if suffix == ".docx":
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if suffix == ".pdf":
        return "application/pdf"
    if suffix in {".png", ".bmp", ".jpg", ".jpeg", ".tif", ".tiff", ".heic"}:
        return "application/octet-stream"
    return "application/octet-stream"


def polygon_to_bbox_norm(polygon: object, page_width: float, page_height: float) -> dict | None:
    if not isinstance(polygon, (list, tuple)) or len(polygon) < 4:
        return None

    try:
        coords = [float(value) for value in polygon]
    except (TypeError, ValueError):
        return None

    xs = coords[0::2]
    ys = coords[1::2]
    if not xs or not ys:
        return None

    x0 = min(xs)
    y0 = min(ys)
    x1 = max(xs)
    y1 = max(ys)
    if page_width > 0:
        x0 /= page_width
        x1 /= page_width
    if page_height > 0:
        y0 /= page_height
        y1 /= page_height

    return {
        "x0": max(0.0, min(1.0, x0)),
        "y0": max(0.0, min(1.0, y0)),
        "x1": max(0.0, min(1.0, x1)),
        "y1": max(0.0, min(1.0, y1)),
    }


def bbox_center(bbox: dict | None) -> tuple[float, float] | None:
    if not bbox:
        return None
    try:
        x0 = float(bbox.get("x0", 0.0))
        y0 = float(bbox.get("y0", 0.0))
        x1 = float(bbox.get("x1", x0))
        y1 = float(bbox.get("y1", y0))
    except (TypeError, ValueError, AttributeError):
        return None
    return ((x0 + x1) / 2.0, (y0 + y1) / 2.0)


def match_formula_items_to_rows(rows: list[dict], formula_items: list[dict] | None) -> dict[int, dict]:
    formula_items = [item for item in (formula_items or []) if isinstance(item, dict)]
    if not formula_items:
        return {}

    page_map: dict[int, list[dict]] = {}
    for item in formula_items:
        page = item.get("page")
        if not isinstance(page, int):
            continue
        page_map.setdefault(page, []).append(item)

    matched: dict[int, dict] = {}
    for row_index, row in enumerate(rows):
        if str(row.get("role", "")).lower() != "equation":
            continue

        row_page = row.get("preview_page")
        row_bbox = row.get("preview_bbox")
        row_center = bbox_center(row_bbox)
        if not isinstance(row_page, int):
            continue

        candidates = page_map.get(row_page, [])
        if not candidates:
            candidates = [item for item in formula_items if isinstance(item.get("page"), int) and abs(int(item["page"]) - row_page) <= 1]
        if not candidates:
            continue

        best_item = None
        best_score = None
        for candidate in candidates:
            candidate_center = candidate.get("center") or bbox_center(candidate.get("bbox_norm"))
            if not candidate_center:
                continue
            dx = abs(candidate_center[0] - (row_center[0] if row_center else candidate_center[0]))
            dy = abs(candidate_center[1] - (row_center[1] if row_center else candidate_center[1]))
            page_penalty = 0.0 if candidate.get("page") == row_page else 0.4
            confidence = candidate.get("confidence")
            confidence_penalty = 0.0
            if isinstance(confidence, (int, float)):
                confidence_penalty = max(0.0, 1.0 - float(confidence)) * 0.15
            score = page_penalty + dx + dy + confidence_penalty
            if best_score is None or score < best_score:
                best_score = score
                best_item = candidate

        if best_item:
            matched[row_index] = best_item

    return matched


def latex_formula_to_alt_text(latex: str, kind: str | None = None) -> str:
    text = normalize_alt_text(latex)
    if not text:
        return ""

    text = text.replace("\u2212", "-")
    text = text.replace("\\left", "").replace("\\right", "")
    text = text.replace("\\,", " ").replace("\\;", " ").replace("\\:", " ").replace("\\!", "")
    text = text.replace("\\qquad", " ").replace("\\quad", " ")
    text = text.replace("$", "")

    replacements = {
        r"\\times": " times ",
        r"\\cdot": " times ",
        r"\\div": " divided by ",
        r"\\pm": " plus or minus ",
        r"\\mp": " minus or plus ",
        r"\\approx": " approximately equal to ",
        r"\\leq": " less than or equal to ",
        r"\\le": " less than or equal to ",
        r"\\geq": " greater than or equal to ",
        r"\\ge": " greater than or equal to ",
        r"\\neq": " not equal to ",
        r"\\infty": " infinity ",
        r"\\pi": " pi ",
        r"\\theta": " theta ",
        r"\\alpha": " alpha ",
        r"\\beta": " beta ",
        r"\\gamma": " gamma ",
        r"\\delta": " delta ",
        r"\\lambda": " lambda ",
        r"\\mu": " mu ",
        r"\\sigma": " sigma ",
        r"\\phi": " phi ",
        r"\\omega": " omega ",
        r"\\sum": " sum ",
        r"\\prod": " product ",
        r"\\int": " integral ",
        r"\\sin": " sine ",
        r"\\cos": " cosine ",
        r"\\tan": " tangent ",
        r"\\log": " log ",
        r"\\ln": " natural log ",
        r"\\exp": " exponential ",
        r"\\deg": " degrees ",
        r"\\cdots": " ellipsis ",
    }
    for pattern, replacement in replacements.items():
        text = re.sub(pattern, replacement, text)

    text = _expand_superscripts(text)
    text = _expand_subscripts(text)
    text = _expand_frac_expressions(text)
    text = _expand_sqrt_expressions(text)

    text = text.replace("{", " ").replace("}", " ")
    text = text.replace("\\", " ")
    text = text.replace("^", " ")
    text = text.replace("_", " ")
    text = text.replace("(", " ")
    text = text.replace(")", " ")
    text = text.replace("[", " ")
    text = text.replace("]", " ")
    text = text.replace(",", ", ")
    text = text.replace("=", " equals ")
    text = text.replace("+", " plus ")
    text = text.replace("-", " minus ")
    text = text.replace("/", " over ")

    text = " ".join(text.split())
    text = _apply_math_phrase_cleanup(text)
    if not text:
        return ""

    if kind == "display" and not text.lower().startswith("equation"):
        text = f"Equation showing {text}"
    elif kind == "inline" and not text.lower().startswith("expression"):
        text = f"Expression showing {text}"
    elif not text.lower().startswith(("equation", "expression")):
        text = f"Equation showing {text}"

    return text


def _expand_frac_expressions(text: str) -> str:
    # Normalize the most common spaced forms first, then repeatedly reduce simple nested fractions.
    text = text.replace("\\frac ", "\\frac")
    text = text.replace("\\frac{", "\\frac{")
    for _ in range(8):
        new_text = re.sub(
            r"\\frac\{([^{}]+)\}\{([^{}]+)\}",
            lambda match: f"({match.group(1)} over {match.group(2)})",
            text,
        )
        if new_text == text:
            break
        text = new_text
    return text


def _expand_sqrt_expressions(text: str) -> str:
    for _ in range(8):
        new_text = re.sub(
            r"\\sqrt\{([^{}]+)\}",
            lambda match: f"square root of {match.group(1)}",
            text,
        )
        if new_text == text:
            break
        text = new_text
    text = text.replace("\\sqrt", "square root of")
    return text


def _expand_superscripts(text: str) -> str:
    superscript_pattern = re.compile(r"([A-Za-z0-9\)\]])\^\{?([A-Za-z0-9+\-]+)\}?")
    for _ in range(8):
        new_text = superscript_pattern.sub(_superscript_replacement, text)
        if new_text == text:
            break
        text = new_text
    return text


def _expand_subscripts(text: str) -> str:
    subscript_pattern = re.compile(r"([A-Za-z0-9\)\]])_\{?([A-Za-z0-9+\-]+)\}?")
    for _ in range(8):
        new_text = subscript_pattern.sub(_subscript_replacement, text)
        if new_text == text:
            break
        text = new_text
    return text


def _superscript_replacement(match: re.Match[str]) -> str:
    base = match.group(1).strip()
    power = match.group(2).strip()
    if power == "2":
        return f"{base} squared"
    if power == "3":
        return f"{base} cubed"
    if power == "1":
        return base
    return f"{base} to the {power}"


def _subscript_replacement(match: re.Match[str]) -> str:
    base = match.group(1).strip()
    subscript = match.group(2).strip()
    return f"{base} sub {subscript}"


def _apply_math_phrase_cleanup(text: str) -> str:
    replacements = [
        ("Equation showing equation showing", "Equation showing"),
        ("Expression showing expression showing", "Expression showing"),
        ("over over", "over"),
        ("times times", "times"),
        ("plus minus", "plus minus"),
        ("minus plus", "minus plus"),
        ("  ", " "),
    ]
    cleaned = text
    for old, new in replacements:
        cleaned = cleaned.replace(old, new)
    cleaned = cleaned.replace(" over )", " over")
    cleaned = cleaned.replace("( ", "(").replace(" )", ")")
    cleaned = " ".join(cleaned.split())
    return cleaned.strip(" ,;")


def inspect_docx_equation_storage(docx_path: Path) -> dict:
    legacy_equations = 0
    omml_equations = 0
    equation_prog_ids = set()

    try:
        with zipfile.ZipFile(docx_path, "r") as archive:
            part_names = [
                name
                for name in archive.namelist()
                if name.startswith("word/")
                and name.endswith(".xml")
                and "/_rels/" not in name
            ]
            for part_name in part_names:
                try:
                    root = ET.fromstring(archive.read(part_name))
                except Exception:
                    continue

                math_count = len(root.findall(".//m:oMath", WORD_NS))
                if math_count == 0:
                    math_count = len(root.findall(".//m:oMathPara", WORD_NS))
                omml_equations += math_count

                for ole in root.findall(".//o:OLEObject", WORD_NS):
                    prog_id = (ole.attrib.get("ProgID", "") or "").strip()
                    prog_id_lower = prog_id.lower()
                    if not prog_id_lower:
                        continue
                    if any(keyword in prog_id_lower for keyword in ("equation", "mathtype", "math")):
                        legacy_equations += 1
                        equation_prog_ids.add(prog_id)
    except (FileNotFoundError, zipfile.BadZipFile):
        return {"legacy_equations": 0, "omml_equations": 0, "prog_ids": []}

    return {
        "legacy_equations": legacy_equations,
        "omml_equations": omml_equations,
        "prog_ids": sorted(equation_prog_ids),
    }


def auto_convert_legacy_equations(docx_path: Path, output_dir: Path | None = None) -> dict:
    before = inspect_docx_equation_storage(docx_path)
    legacy_before = int(before.get("legacy_equations", 0))
    omml_before = int(before.get("omml_equations", 0))

    result = {
        "source_docx_path": docx_path,
        "working_docx_path": docx_path,
        "attempted": False,
        "converted": False,
        "legacy_before": legacy_before,
        "legacy_after": legacy_before,
        "omml_before": omml_before,
        "omml_after": omml_before,
        "automation": {},
        "message": "",
    }

    if legacy_before <= 0:
        result["message"] = "No legacy equation objects detected."
        return result

    work_root = output_dir if isinstance(output_dir, Path) else docx_path.parent
    work_root.mkdir(parents=True, exist_ok=True)
    working_docx_path = work_root / f"{docx_path.stem}_office_math{docx_path.suffix}"
    try:
        working_docx_path.write_bytes(docx_path.read_bytes())
    except Exception as exc:
        result["message"] = f"Could not prepare a working DOCX copy: {exc}"
        return result

    automation = run_word_office_math_conversion(working_docx_path)
    result["automation"] = automation
    result["attempted"] = bool(automation.get("attempted"))

    after = inspect_docx_equation_storage(working_docx_path)
    legacy_after = int(after.get("legacy_equations", legacy_before))
    omml_after = int(after.get("omml_equations", omml_before))

    result["legacy_after"] = legacy_after
    result["omml_after"] = omml_after

    converted = (legacy_after < legacy_before) or (omml_after > omml_before)
    result["converted"] = converted
    if converted:
        result["working_docx_path"] = working_docx_path
        result["message"] = (
            f"Auto-converted equations: legacy {legacy_before} -> {legacy_after}, "
            f"OMML {omml_before} -> {omml_after}."
        )
    else:
        result["working_docx_path"] = docx_path
        if working_docx_path.exists():
            try:
                working_docx_path.unlink()
            except Exception:
                pass
        automation_error = (automation.get("error") or "").strip()
        if automation_error:
            result["message"] = (
                f"Auto-conversion attempted but did not complete: {summarize_automation_error(automation_error)}"
            )
        else:
            result["message"] = (
                "Auto-conversion attempted but no equation conversion changes were detected."
            )

    return result


def summarize_automation_error(error: str) -> str:
    text = (error or "").strip()
    lowered = text.lower()
    if "com class factory" in lowered or "word.application" in lowered:
        return "Microsoft Word automation is unavailable in the current runtime session."
    if "timeout" in lowered:
        return "Word automation timed out while converting equation objects."
    return text


def run_word_office_math_conversion(docx_path: Path) -> dict:
    script_body = r"""
param(
  [Parameter(Mandatory = $true)]
  [string]$TargetPath
)

$ErrorActionPreference = 'Stop'
$result = [ordered]@{
  attempted = $false
  opened = $false
  legacy_objects = 0
  trigger_attempts = 0
  trigger_successes = 0
  command_candidates = @(
    "ConvertToOfficeMath",
    "EquationConvertToOfficeMath",
    "EquationConvert",
    "ConvertMath",
    "ConvertEquation"
  )
  commands_used = @()
  error = ""
}

$word = $null
$doc = $null
$used = New-Object System.Collections.Generic.HashSet[string]

try {
  $result.attempted = $true
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.ScreenUpdating = $false
  $word.DisplayAlerts = 0

  $doc = $word.Documents.Open($TargetPath, $false, $false, $false)
  $result.opened = $true

  try { $doc.Convert() } catch {}

  foreach ($shape in @($doc.Shapes)) {
    try { $shape.ConvertToInlineShape() | Out-Null } catch {}
  }

  foreach ($inlineShape in @($doc.InlineShapes)) {
    $shapeType = 0
    try { $shapeType = [int]$inlineShape.Type } catch {}
    if ($shapeType -ne 1 -and $shapeType -ne 2) { continue }

    $progId = ""
    try { $progId = [string]$inlineShape.OLEFormat.ProgID } catch {}
    if ([string]::IsNullOrWhiteSpace($progId)) { continue }
    if ($progId -notmatch '(?i)(equation|mathtype|math)') { continue }

    $result.legacy_objects = [int]$result.legacy_objects + 1
    try { $inlineShape.Range.Select() | Out-Null } catch {}

    foreach ($idMso in $result.command_candidates) {
      $result.trigger_attempts = [int]$result.trigger_attempts + 1
      $enabled = $false
      try { $enabled = $word.CommandBars.GetEnabledMso($idMso) } catch {}
      if (-not $enabled) { continue }
      try {
        $word.CommandBars.ExecuteMso($idMso)
        $used.Add($idMso) | Out-Null
        $result.trigger_successes = [int]$result.trigger_successes + 1
      } catch {}
    }

    $result.trigger_attempts = [int]$result.trigger_attempts + 1
    try {
      $inlineShape.OLEFormat.DoVerb(0)
      $result.trigger_successes = [int]$result.trigger_successes + 1
    } catch {}
  }

  try {
    if ($doc.OMaths.Count -gt 0) {
      $doc.OMaths.BuildUp()
      $doc.OMaths.Linearize()
    }
  } catch {}

  try {
    $doc.SaveAs2($TargetPath, 16)
  } catch {
    $doc.Save()
  }
}
catch {
  $result.error = $_.Exception.Message
}
finally {
  $result.commands_used = @($used)
  if ($doc -ne $null) {
    try { $doc.Close(0) | Out-Null } catch {}
  }
  if ($word -ne $null) {
    try { $word.Quit() | Out-Null } catch {}
  }
}

Write-Output ("MATCHA_ALT_JSON::" + ($result | ConvertTo-Json -Depth 5 -Compress))
"""

    script_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8") as handle:
            handle.write(script_body)
            script_path = Path(handle.name)

        completed = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
                "-TargetPath",
                str(docx_path),
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=240,
            check=False,
        )
        parsed = {
            "attempted": True,
            "error": "",
            "stdout": completed.stdout,
            "stderr": completed.stderr,
            "returncode": completed.returncode,
        }
        for line in (completed.stdout or "").splitlines():
            if not line.startswith("MATCHA_ALT_JSON::"):
                continue
            payload = line.split("MATCHA_ALT_JSON::", 1)[1].strip()
            if not payload:
                continue
            try:
                info = json.loads(payload)
                parsed.update(info if isinstance(info, dict) else {})
            except json.JSONDecodeError:
                continue
        if completed.returncode != 0 and not str(parsed.get("error", "")).strip():
            parsed["error"] = (completed.stderr or "").strip() or f"Word automation failed with exit code {completed.returncode}."
        return parsed
    except Exception as exc:
        return {"attempted": True, "error": str(exc)}
    finally:
        if script_path is not None and script_path.exists():
            try:
                script_path.unlink()
            except OSError:
                pass


def generate_missing_alt_rows(
    rows: list[dict],
    pdf_path: Path,
    docx_path: Path | None = None,
    formula_items: list[dict] | None = None,
) -> dict:
    generated_count = 0
    formula_items = formula_items or []
    omml_candidates = extract_omml_equation_candidates(docx_path) if isinstance(docx_path, Path) else []
    equation_row_indexes = [idx for idx, row in enumerate(rows) if str(row.get("role", "")).lower() == "equation"]
    equation_order_map = {row_index: order for order, row_index in enumerate(equation_row_indexes)}
    formula_matches = match_formula_items_to_rows(rows, formula_items)

    for row_index, row in enumerate(rows):
        role = str(row.get("role", "")).lower()
        formula_match = formula_matches.get(row_index)
        if formula_match:
            row["ocr_formula_latex"] = formula_match.get("latex", "")
            row["ocr_formula_kind"] = formula_match.get("kind", "")
            row["ocr_formula_confidence"] = formula_match.get("confidence")
            row["ocr_provider"] = "local"
            row["ocr_formula_page"] = formula_match.get("page_number")

        existing_text = (row.get("existing_alt_text") or "").strip()
        if existing_text:
            row["alt_text"] = existing_text
            row["generated_alt_text"] = row.get("generated_alt_text", "")
            row["alt_source"] = "existing"
            continue

        current_generated = (row.get("generated_alt_text") or "").strip()
        if current_generated:
            row["alt_text"] = current_generated
            row["alt_source"] = "generated"
            continue

        # Focus mode: skip generated ALT for equations unless explicitly enabled.
        if role == "equation" and not is_equation_generation_enabled():
            row["alt_text"] = ""
            row["alt_source"] = "missing_ocr"
            continue

        suggestion = ""
        if role == "equation" and formula_match:
            suggestion = finalize_alt_text(
                formula_latex_to_spoken_alt_text(
                    formula_match.get("latex", ""),
                    formula_match.get("kind", ""),
                ),
                role,
                row.get("page"),
            )

        if not suggestion:
            suggestion = suggest_alt_text_for_row(
                row,
                pdf_path,
                docx_path=docx_path,
                rows=rows,
                row_index=row_index,
                omml_candidates=omml_candidates,
                equation_order_map=equation_order_map,
            )

        if suggestion:
            row["generated_alt_text"] = suggestion
            row["alt_text"] = suggestion
            row["alt_source"] = "ocr_formula" if formula_match else "generated"
            generated_count += 1
        else:
            row["alt_text"] = ""
            row["alt_source"] = "missing_ocr" if role == "equation" else "missing"

    return {
        "generated_count": generated_count,
        "summary": summarize_alt_rows(rows),
        "rows": rows,
        "ocr_provider": "local",
        "ocr_message": "",
        "ocr_items": len(formula_items),
        "formula_items": formula_items,
    }


def get_groq_alt_config() -> dict | None:
    api_keys = collect_groq_api_keys()
    if not api_keys:
        return None

    base_url = (os.getenv("MATCHA_GROQ_BASE_URL") or "https://api.groq.com/openai/v1").strip().rstrip("/")
    model = (os.getenv("MATCHA_GROQ_VISION_MODEL") or "meta-llama/llama-4-scout-17b-16e-instruct").strip()
    if not base_url or not model:
        return None

    if len(api_keys) >= 4:
        ocr_api_keys = api_keys[:2]
        rewrite_api_keys = api_keys[2:4]
    elif len(api_keys) == 3:
        ocr_api_keys = api_keys[:1]
        rewrite_api_keys = api_keys[1:3]
    elif len(api_keys) == 2:
        ocr_api_keys = api_keys[:1]
        rewrite_api_keys = api_keys[1:2]
    else:
        ocr_api_keys = api_keys[:1]
        rewrite_api_keys = api_keys[:1]

    return {
        "api_key": api_keys[0],
        "api_keys": api_keys,
        "ocr_api_keys": ocr_api_keys,
        "rewrite_api_keys": rewrite_api_keys,
        "base_url": base_url,
        "model": model,
        "timeout_seconds": int(os.getenv("MATCHA_GROQ_TIMEOUT_SECONDS", "25") or "25"),
        "max_completion_tokens": int(os.getenv("MATCHA_GROQ_MAX_COMPLETION_TOKENS", "80") or "80"),
        "ocr_max_completion_tokens": int(os.getenv("MATCHA_GROQ_OCR_MAX_COMPLETION_TOKENS", "96") or "96"),
        "rewrite_max_completion_tokens": int(os.getenv("MATCHA_GROQ_REWRITE_MAX_COMPLETION_TOKENS", "160") or "160"),
        "grid_batch_size": int(os.getenv("MATCHA_GROQ_GRID_BATCH_SIZE", "8") or "8"),
        "grid_max_completion_tokens": int(os.getenv("MATCHA_GROQ_GRID_MAX_COMPLETION_TOKENS", "220") or "220"),
        "grid_ocr_max_completion_tokens": int(os.getenv("MATCHA_GROQ_GRID_OCR_MAX_COMPLETION_TOKENS", "240") or "240"),
        "grid_rewrite_max_completion_tokens": int(os.getenv("MATCHA_GROQ_GRID_REWRITE_MAX_COMPLETION_TOKENS", "480") or "480"),
        "grid_cols": int(os.getenv("MATCHA_GROQ_GRID_COLS", "4") or "4"),
        "grid_cell_width": int(os.getenv("MATCHA_GROQ_GRID_CELL_WIDTH", "180") or "180"),
        "grid_cell_height": int(os.getenv("MATCHA_GROQ_GRID_CELL_HEIGHT", "140") or "140"),
        "grid_margin": int(os.getenv("MATCHA_GROQ_GRID_MARGIN", "12") or "12"),
        "grid_image_max_side": int(os.getenv("MATCHA_GROQ_GRID_IMAGE_MAX_SIDE", "1200") or "1200"),
        "grid_image_jpeg_quality": int(os.getenv("MATCHA_GROQ_GRID_IMAGE_JPEG_QUALITY", "68") or "68"),
        "rate_limit_retries": int(os.getenv("MATCHA_GROQ_RATE_LIMIT_RETRIES", "0") or "0"),
        "retry_after_cap_seconds": float(os.getenv("MATCHA_GROQ_RETRY_AFTER_CAP_SECONDS", "3") or "3"),
    }


def collect_claude_api_keys() -> list[tuple[str, str]]:
    candidates = [
        ("MATCHA_CLAUDE_API_KEY", os.getenv("MATCHA_CLAUDE_API_KEY")),
        ("CLAUDE_API_KEY", os.getenv("CLAUDE_API_KEY")),
        ("ANTHROPIC_API_KEY", os.getenv("ANTHROPIC_API_KEY")),
    ]
    keys: list[tuple[str, str]] = []
    seen = set()
    for source_name, candidate in candidates:
        value = (candidate or "").strip()
        if not value or value in seen:
            continue
        seen.add(value)
        keys.append((source_name, value))
    return keys


def get_claude_alt_config() -> dict | None:
    api_keys = collect_claude_api_keys()
    if not api_keys:
        return None

    return {
        "api_key": api_keys[0][1],
        "api_keys": [key for _source_name, key in api_keys],
        "api_key_sources": [source_name for source_name, _key in api_keys],
        "base_url": (os.getenv("MATCHA_CLAUDE_BASE_URL") or "https://api.anthropic.com/v1").strip().rstrip("/"),
        "model": (os.getenv("MATCHA_CLAUDE_VISION_MODEL") or "claude-sonnet-4-6").strip(),
        "timeout_seconds": max(5, int(os.getenv("MATCHA_CLAUDE_TIMEOUT_SECONDS", "45") or "45")),
        "max_tokens": max(256, int(os.getenv("MATCHA_CLAUDE_MAX_TOKENS", "1024") or "1024")),
        "grid_batch_size": int(os.getenv("MATCHA_CLAUDE_GRID_BATCH_SIZE", "24") or "24"),
        "grid_cols": int(os.getenv("MATCHA_CLAUDE_GRID_COLS", "4") or "4"),
        "grid_cell_width": int(os.getenv("MATCHA_CLAUDE_GRID_CELL_WIDTH", "180") or "180"),
        "grid_cell_height": int(os.getenv("MATCHA_CLAUDE_GRID_CELL_HEIGHT", "140") or "140"),
        "grid_margin": int(os.getenv("MATCHA_CLAUDE_GRID_MARGIN", "12") or "12"),
        "grid_max_tokens": max(512, int(os.getenv("MATCHA_CLAUDE_GRID_MAX_TOKENS", "4800") or "4800")),
        "grid_image_max_side": int(os.getenv("MATCHA_CLAUDE_GRID_IMAGE_MAX_SIDE", "1568") or "1568"),
        "grid_image_jpeg_quality": int(os.getenv("MATCHA_CLAUDE_GRID_IMAGE_JPEG_QUALITY", "72") or "72"),
    }


def get_openrouter_alt_config() -> dict | None:
    api_key = (os.getenv("OPENROUTER_API_KEY") or "").strip()
    if not api_key:
        return None

    return {
        "api_key": api_key,
        "base_url": (os.getenv("MATCHA_OPENROUTER_BASE_URL") or "https://openrouter.ai/api/v1").strip().rstrip("/"),
        "model": (os.getenv("MATCHA_OPENROUTER_VISION_MODEL") or "anthropic/claude-sonnet-4").strip(),
        "timeout_seconds": max(5, int(os.getenv("MATCHA_OPENROUTER_TIMEOUT_SECONDS", "45") or "45")),
        "max_tokens": max(256, int(os.getenv("MATCHA_OPENROUTER_MAX_TOKENS", "1024") or "1024")),
        "grid_batch_size": int(os.getenv("MATCHA_OPENROUTER_GRID_BATCH_SIZE", "24") or "24"),
        "grid_cols": int(os.getenv("MATCHA_OPENROUTER_GRID_COLS", "4") or "4"),
        "grid_cell_width": int(os.getenv("MATCHA_OPENROUTER_GRID_CELL_WIDTH", "180") or "180"),
        "grid_cell_height": int(os.getenv("MATCHA_OPENROUTER_GRID_CELL_HEIGHT", "140") or "140"),
        "grid_margin": int(os.getenv("MATCHA_OPENROUTER_GRID_MARGIN", "12") or "12"),
        "grid_max_tokens": max(512, int(os.getenv("MATCHA_OPENROUTER_GRID_MAX_TOKENS", "4800") or "4800")),
        "grid_image_max_side": int(os.getenv("MATCHA_OPENROUTER_GRID_IMAGE_MAX_SIDE", "1568") or "1568"),
        "grid_image_jpeg_quality": int(os.getenv("MATCHA_OPENROUTER_GRID_IMAGE_JPEG_QUALITY", "72") or "72"),
        "http_referer": (os.getenv("MATCHA_OPENROUTER_HTTP_REFERER") or "").strip(),
        "x_title": (os.getenv("MATCHA_OPENROUTER_X_TITLE") or "Altomizer").strip(),
    }


def get_copilot_alt_config() -> dict | None:
    api_key = (os.getenv("MATCHA_COPILOT_API_KEY") or os.getenv("GITHUB_TOKEN") or os.getenv("GH_TOKEN") or "").strip()
    if not api_key:
        return None

    return {
        "api_key": api_key,
        "base_url": (os.getenv("MATCHA_COPILOT_BASE_URL") or "https://models.github.ai").strip().rstrip("/"),
        "model": (os.getenv("MATCHA_COPILOT_VISION_MODEL") or "openai/gpt-4.1").strip(),
        "api_version": (os.getenv("MATCHA_COPILOT_API_VERSION") or "2026-03-10").strip(),
        "org": (os.getenv("MATCHA_COPILOT_ORG") or "").strip(),
        "timeout_seconds": max(5, int(os.getenv("MATCHA_COPILOT_TIMEOUT_SECONDS", "45") or "45")),
        "max_tokens": max(256, int(os.getenv("MATCHA_COPILOT_MAX_TOKENS", "1024") or "1024")),
        "grid_batch_size": int(os.getenv("MATCHA_COPILOT_GRID_BATCH_SIZE", "24") or "24"),
        "grid_cols": int(os.getenv("MATCHA_COPILOT_GRID_COLS", "4") or "4"),
        "grid_cell_width": int(os.getenv("MATCHA_COPILOT_GRID_CELL_WIDTH", "180") or "180"),
        "grid_cell_height": int(os.getenv("MATCHA_COPILOT_GRID_CELL_HEIGHT", "140") or "140"),
        "grid_margin": int(os.getenv("MATCHA_COPILOT_GRID_MARGIN", "12") or "12"),
        "grid_max_tokens": max(512, int(os.getenv("MATCHA_COPILOT_GRID_MAX_TOKENS", "4800") or "4800")),
        "grid_image_max_side": int(os.getenv("MATCHA_COPILOT_GRID_IMAGE_MAX_SIDE", "1568") or "1568"),
        "grid_image_jpeg_quality": int(os.getenv("MATCHA_COPILOT_GRID_IMAGE_JPEG_QUALITY", "72") or "72"),
    }


def get_gemini_alt_config() -> dict | None:
    api_key = (os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or "").strip()
    if not api_key:
        return None

    return {
        "api_key": api_key,
        "base_url": (os.getenv("MATCHA_GEMINI_BASE_URL") or "https://generativelanguage.googleapis.com/v1beta").strip().rstrip("/"),
        "model": (os.getenv("MATCHA_GEMINI_VISION_MODEL") or "gemini-3.5-flash").strip(),
        "timeout_seconds": max(5, int(os.getenv("MATCHA_GEMINI_TIMEOUT_SECONDS", "45") or "45")),
        "max_tokens": max(256, int(os.getenv("MATCHA_GEMINI_MAX_TOKENS", "1024") or "1024")),
        "grid_batch_size": int(os.getenv("MATCHA_GEMINI_GRID_BATCH_SIZE", "24") or "24"),
        "grid_cols": int(os.getenv("MATCHA_GEMINI_GRID_COLS", "4") or "4"),
        "grid_cell_width": int(os.getenv("MATCHA_GEMINI_GRID_CELL_WIDTH", "180") or "180"),
        "grid_cell_height": int(os.getenv("MATCHA_GEMINI_GRID_CELL_HEIGHT", "140") or "140"),
        "grid_margin": int(os.getenv("MATCHA_GEMINI_GRID_MARGIN", "12") or "12"),
        "grid_max_tokens": max(512, int(os.getenv("MATCHA_GEMINI_GRID_MAX_TOKENS", "4800") or "4800")),
        "grid_image_max_side": int(os.getenv("MATCHA_GEMINI_GRID_IMAGE_MAX_SIDE", "1568") or "1568"),
        "grid_image_jpeg_quality": int(os.getenv("MATCHA_GEMINI_GRID_IMAGE_JPEG_QUALITY", "72") or "72"),
    }


def collect_groq_api_keys() -> list[str]:
    candidates = [
        os.getenv("MATCHA_GROQ_API_KEY_1"),
        os.getenv("MATCHA_GROQ_API_KEY_2"),
        os.getenv("MATCHA_GROQ_API_KEY_3"),
        os.getenv("MATCHA_GROQ_API_KEY_4"),
        os.getenv("GROQ_API_KEY_1"),
        os.getenv("GROQ_API_KEY_2"),
        os.getenv("GROQ_API_KEY_3"),
        os.getenv("GROQ_API_KEY_4"),
        os.getenv("MATCHA_GROQ_API_KEY"),
        os.getenv("GROQ_API_KEY"),
    ]
    keys: list[str] = []
    seen = set()
    for candidate in candidates:
        value = (candidate or "").strip()
        if not value or value in seen:
            continue
        seen.add(value)
        keys.append(value)
        if len(keys) >= 4:
            break
    return keys


def generate_missing_alt_rows_with_groq(
    rows: list[dict],
    preview_images: dict[int, dict] | None,
    pdf_path: Path | None,
    docx_path: Path | None = None,
) -> dict:
    config = get_groq_alt_config()
    if not config:
        return {
            "available": False,
            "message": "Groq ALT generation is not configured. Set GROQ_API_KEY or GROQ_API_KEY_1..4 to enable it.",
            "generated_count": 0,
            "rows": rows,
            "summary": summarize_alt_rows(rows),
            "provider": "groq",
        }

    preview_images = preview_images or {}
    generated_count = 0
    error_count = 0
    first_error = ""
    batch_candidates: list[dict] = []
    for row in rows:
        role = str(row.get("role", "")).lower()
        if role not in {"image", "equation"}:
            continue
        current_alt = normalize_alt_text(str(row.get("alt_text", "") or ""))
        if current_alt:
            continue

        preview_entry = None
        row_id = row.get("id")
        if isinstance(row_id, int):
            preview_entry = preview_images.get(row_id)
        if preview_entry is None:
            preview_entry = build_alt_preview_entry(row, docx_path, pdf_path)
        if not isinstance(preview_entry, dict):
            error_count += 1
            continue
        batch_candidates.append({"row": row, "preview_entry": preview_entry})

    try:
        batch_result = generate_alt_text_batches_with_groq(batch_candidates, config, pdf_path, docx_path)
        generated_count += int(batch_result.get("generated_count", 0) or 0)
        error_count += int(batch_result.get("error_count", 0) or 0)
        first_error = normalize_alt_text(str(batch_result.get("first_error", "") or ""))
    except Exception as exc:
        error_count += len(batch_candidates)
        first_error = normalize_alt_text(str(exc))

    summary = summarize_alt_rows(rows)
    message = f"Generated ALT text for {generated_count} item(s) with Groq."
    if error_count:
        message += f" Skipped {error_count} item(s)"
        if first_error:
            message += f": {first_error}"
        else:
            message += "."
    return {
        "available": True,
        "message": message,
        "generated_count": generated_count,
        "rows": rows,
        "summary": summary,
        "provider": "groq",
    }


def generate_missing_alt_rows_with_claude(
    rows: list[dict],
    preview_images: dict[int, dict] | None,
    pdf_path: Path | None,
    docx_path: Path | None = None,
) -> dict:
    config = get_claude_alt_config()
    if not config:
        return {
            "available": False,
            "message": "Claude ALT generation is not configured. Set ANTHROPIC_API_KEY to enable it.",
            "generated_count": 0,
            "rows": rows,
            "summary": summarize_alt_rows(rows),
            "provider": "claude",
        }

    preview_images = preview_images or {}
    generated_count = 0
    error_count = 0
    first_error = ""
    batch_candidates: list[dict] = []
    for row in rows:
        role = str(row.get("role", "")).lower()
        if role not in {"image", "equation"}:
            continue
        current_alt = normalize_alt_text(str(row.get("alt_text", "") or ""))
        if current_alt:
            continue

        preview_entry = None
        row_id = row.get("id")
        if isinstance(row_id, int):
            preview_entry = preview_images.get(row_id)
        if preview_entry is None:
            preview_entry = build_alt_preview_entry(row, docx_path, pdf_path)
        if not isinstance(preview_entry, dict):
            error_count += 1
            continue
        batch_candidates.append({"row": row, "preview_entry": preview_entry})

    try:
        batch_result = generate_alt_text_batches_with_claude(batch_candidates, config, pdf_path, docx_path)
        generated_count += int(batch_result.get("generated_count", 0) or 0)
        error_count += int(batch_result.get("error_count", 0) or 0)
        first_error = normalize_alt_text(str(batch_result.get("first_error", "") or ""))
    except Exception as exc:
        error_count += len(batch_candidates)
        first_error = normalize_alt_text(str(exc))

    summary = summarize_alt_rows(rows)
    message = f"Generated ALT text for {generated_count} item(s) with Claude."
    if error_count:
        message += f" Skipped {error_count} item(s)"
        if first_error:
            message += f": {first_error}"
        else:
            message += "."
    return {
        "available": True,
        "message": message,
        "generated_count": generated_count,
        "rows": rows,
        "summary": summary,
        "provider": "claude",
    }


def generate_missing_alt_rows_with_openrouter(
    rows: list[dict],
    preview_images: dict[int, dict] | None,
    pdf_path: Path | None,
    docx_path: Path | None = None,
) -> dict:
    config = get_openrouter_alt_config()
    if not config:
        return {
            "available": False,
            "message": "OpenRouter ALT generation is not configured. Set OPENROUTER_API_KEY to enable it.",
            "generated_count": 0,
            "rows": rows,
            "summary": summarize_alt_rows(rows),
            "provider": "openrouter",
        }

    preview_images = preview_images or {}
    generated_count = 0
    error_count = 0
    first_error = ""
    batch_candidates: list[dict] = []
    for row in rows:
        role = str(row.get("role", "")).lower()
        if role not in {"image", "equation"}:
            continue
        current_alt = normalize_alt_text(str(row.get("alt_text", "") or ""))
        if current_alt:
            continue

        preview_entry = None
        row_id = row.get("id")
        if isinstance(row_id, int):
            preview_entry = preview_images.get(row_id)
        if preview_entry is None:
            preview_entry = build_alt_preview_entry(row, docx_path, pdf_path)
        if not isinstance(preview_entry, dict):
            error_count += 1
            continue
        batch_candidates.append({"row": row, "preview_entry": preview_entry})

    try:
        batch_result = generate_alt_text_batches_with_openrouter(batch_candidates, config, pdf_path, docx_path)
        generated_count += int(batch_result.get("generated_count", 0) or 0)
        error_count += int(batch_result.get("error_count", 0) or 0)
        first_error = normalize_alt_text(str(batch_result.get("first_error", "") or ""))
    except Exception as exc:
        error_count += len(batch_candidates)
        first_error = normalize_alt_text(str(exc))

    summary = summarize_alt_rows(rows)
    message = f"Generated ALT text for {generated_count} item(s) with OpenRouter."
    if error_count:
        message += f" Skipped {error_count} item(s)"
        if first_error:
            message += f": {first_error}"
        else:
            message += "."
    return {
        "available": True,
        "message": message,
        "generated_count": generated_count,
        "rows": rows,
        "summary": summary,
        "provider": "openrouter",
    }


def generate_missing_alt_rows_with_copilot(
    rows: list[dict],
    preview_images: dict[int, dict] | None,
    pdf_path: Path | None,
    docx_path: Path | None = None,
) -> dict:
    config = get_copilot_alt_config()
    if not config:
        return {
            "available": False,
            "message": "Copilot ALT generation is not configured. Set MATCHA_COPILOT_API_KEY to a GitHub token with models: read to enable it.",
            "generated_count": 0,
            "rows": rows,
            "summary": summarize_alt_rows(rows),
            "provider": "copilot",
        }

    preview_images = preview_images or {}
    generated_count = 0
    error_count = 0
    first_error = ""
    batch_candidates: list[dict] = []
    for row in rows:
        role = str(row.get("role", "")).lower()
        if role not in {"image", "equation"}:
            continue
        current_alt = normalize_alt_text(str(row.get("alt_text", "") or ""))
        if current_alt:
            continue

        preview_entry = None
        row_id = row.get("id")
        if isinstance(row_id, int):
            preview_entry = preview_images.get(row_id)
        if preview_entry is None:
            preview_entry = build_alt_preview_entry(row, docx_path, pdf_path)
        if not isinstance(preview_entry, dict):
            error_count += 1
            continue
        batch_candidates.append({"row": row, "preview_entry": preview_entry})

    try:
        batch_result = generate_alt_text_batches_with_copilot(batch_candidates, config, pdf_path, docx_path)
        generated_count += int(batch_result.get("generated_count", 0) or 0)
        error_count += int(batch_result.get("error_count", 0) or 0)
        first_error = normalize_alt_text(str(batch_result.get("first_error", "") or ""))
    except Exception as exc:
        error_count += len(batch_candidates)
        first_error = normalize_alt_text(str(exc))

    summary = summarize_alt_rows(rows)
    message = f"Generated ALT text for {generated_count} item(s) with Copilot."
    if error_count:
        message += f" Skipped {error_count} item(s)"
        if first_error:
            message += f": {first_error}"
        else:
            message += "."
    return {
        "available": True,
        "message": message,
        "generated_count": generated_count,
        "rows": rows,
        "summary": summary,
        "provider": "copilot",
    }


def generate_missing_alt_rows_with_gemini(
    rows: list[dict],
    preview_images: dict[int, dict] | None,
    pdf_path: Path | None,
    docx_path: Path | None = None,
) -> dict:
    config = get_gemini_alt_config()
    if not config:
        return {
            "available": False,
            "message": "Gemini ALT generation is not configured. Set GEMINI_API_KEY to enable it.",
            "generated_count": 0,
            "rows": rows,
            "summary": summarize_alt_rows(rows),
            "provider": "gemini",
        }

    preview_images = preview_images or {}
    generated_count = 0
    error_count = 0
    first_error = ""
    batch_candidates: list[dict] = []
    for row in rows:
        role = str(row.get("role", "")).lower()
        if role not in {"image", "equation"}:
            continue
        current_alt = normalize_alt_text(str(row.get("alt_text", "") or ""))
        if current_alt:
            continue

        preview_entry = None
        row_id = row.get("id")
        if isinstance(row_id, int):
            preview_entry = preview_images.get(row_id)
        if preview_entry is None:
            preview_entry = build_alt_preview_entry(row, docx_path, pdf_path)
        if not isinstance(preview_entry, dict):
            error_count += 1
            continue
        batch_candidates.append({"row": row, "preview_entry": preview_entry})

    try:
        batch_result = generate_alt_text_batches_with_gemini(batch_candidates, config, pdf_path, docx_path)
        generated_count += int(batch_result.get("generated_count", 0) or 0)
        error_count += int(batch_result.get("error_count", 0) or 0)
        first_error = normalize_alt_text(str(batch_result.get("first_error", "") or ""))
    except Exception as exc:
        error_count += len(batch_candidates)
        first_error = normalize_alt_text(str(exc))

    summary = summarize_alt_rows(rows)
    message = f"Generated ALT text for {generated_count} item(s) with Gemini."
    if error_count:
        message += f" Skipped {error_count} item(s)"
        if first_error:
            message += f": {first_error}"
        else:
            message += "."
    return {
        "available": True,
        "message": message,
        "generated_count": generated_count,
        "rows": rows,
        "summary": summary,
        "provider": "gemini",
    }


def generate_alt_rows_via_workbook_with_groq(
    rows: list[dict],
    preview_images: dict[int, dict] | None,
    source_filename: str,
    pdf_path: Path | None,
    docx_path: Path | None = None,
) -> dict:
    config = get_groq_alt_config()
    if not config:
        return {
            "available": False,
            "message": "Groq ALT generation is not configured. Set GROQ_API_KEY or GROQ_API_KEY_1..4 to enable it.",
            "generated_count": 0,
            "rows": rows,
            "summary": summarize_alt_rows(rows),
            "provider": "groq_workbook",
        }

    preview_images = preview_images or {}
    workbook_rows: list[dict] = []
    workbook_row_refs: list[dict] = []

    for row in rows:
        row_id = row.get("id")
        image_entry = preview_images.get(row_id) if isinstance(row_id, int) else None
        image_bytes, image_ext = parse_preview_entry(image_entry)
        if not isinstance(image_bytes, (bytes, bytearray)) or not image_bytes:
            continue
        workbook_rows.append(
            {
                "id": row.get("id"),
                "type": row.get("type", "Item"),
                "role": row.get("role", ""),
                "source_part": row.get("source_part", ""),
                "page": row.get("page"),
                "alt_text": row.get("alt_text", ""),
                "existing_alt_text": row.get("existing_alt_text", ""),
                "generated_alt_text": row.get("generated_alt_text", ""),
                "_workbook_ext": image_ext,
            }
        )
        workbook_row_refs.append(row)

    if not workbook_rows:
        return {
            "available": False,
            "message": "No workbook preview images were available for Groq ALT generation.",
            "generated_count": 0,
            "rows": rows,
            "summary": summarize_alt_rows(rows),
            "provider": "groq_workbook",
        }

    from Altomizer.docx_tools import extract_excel_images

    workbook_bytes = build_alt_excel(workbook_rows, source_filename or "document.docx", preview_images=preview_images)
    workbook_images = extract_excel_images(workbook_bytes)
    if not workbook_images:
        return {
            "available": False,
            "message": "No embedded workbook images were available for Groq ALT generation.",
            "generated_count": 0,
            "rows": rows,
            "summary": summarize_alt_rows(rows),
            "provider": "groq_workbook",
        }

    generated_count = 0
    error_count = 0
    first_error = ""
    key_cursor = 0

    for index, image_info in enumerate(workbook_images):
        if index >= len(workbook_row_refs):
            break
        row = workbook_row_refs[index]
        current_alt = normalize_alt_text(str(row.get("alt_text", "") or ""))
        if current_alt:
            continue

        image = image_info.get("image")
        if image is None:
            error_count += 1
            continue
        try:
            image_buffer = io.BytesIO()
            image.save(image_buffer, format="PNG")
            generated_alt = request_groq_ocr_then_rewrite_alt_text(
                image_buffer.getvalue(),
                "image/png",
                row,
                config,
                pdf_path,
                docx_path=docx_path,
            )
        except Exception as exc:
            error_count += 1
            if not first_error:
                first_error = normalize_alt_text(str(exc))
            continue

        if not generated_alt:
            error_count += 1
            continue

        row["generated_alt_text"] = generated_alt
        row["alt_text"] = generated_alt
        row["effective_alt_text"] = generated_alt
        row["alt_source"] = "generated_groq_workbook"
        generated_count += 1

    summary = summarize_alt_rows(rows)
    message = f"Generated ALT text for {generated_count} item(s) from the ALT workbook with Groq."
    if error_count:
        message += f" Skipped {error_count} item(s)"
        if first_error:
            message += f": {first_error}"
        else:
            message += "."
    return {
        "available": True,
        "message": message,
        "generated_count": generated_count,
        "rows": rows,
        "summary": summary,
        "provider": "groq_workbook",
    }


def generate_alt_text_batches_with_groq(
    batch_candidates: list[dict],
    config: dict,
    pdf_path: Path | None,
    docx_path: Path | None = None,
) -> dict:
    if not batch_candidates:
        return {"generated_count": 0, "error_count": 0, "first_error": ""}

    from Altomizer.docx_tools import extract_excel_images, make_grids

    batch_size = max(1, min(int(config.get("grid_batch_size", 24) or 24), 24))
    grid_cols = max(1, min(int(config.get("grid_cols", 4) or 4), batch_size))
    grid_cell_width = max(96, int(config.get("grid_cell_width", 180) or 180))
    grid_cell_height = max(96, int(config.get("grid_cell_height", 140) or 140))
    grid_margin = max(4, int(config.get("grid_margin", 12) or 12))
    synthetic_rows: list[dict] = []
    synthetic_previews: dict[int, dict] = {}
    row_refs: list[dict] = []
    preview_refs: list[dict] = []

    for index, candidate in enumerate(batch_candidates):
        row = candidate.get("row")
        preview_entry = candidate.get("preview_entry")
        if not isinstance(row, dict) or not isinstance(preview_entry, dict):
            continue
        synthetic_rows.append(
            {
                "id": index,
                "type": row.get("type", "Item"),
                "role": row.get("role", ""),
                "source_part": row.get("source_part", ""),
                "page": row.get("page"),
                "alt_text": "",
                "existing_alt_text": "",
                "generated_alt_text": "",
            }
        )
        synthetic_previews[index] = preview_entry
        row_refs.append(row)
        preview_refs.append(preview_entry)

    if not synthetic_rows:
        return {"generated_count": 0, "error_count": 0, "first_error": ""}

    equation_ratio = (
        sum(1 for row in row_refs if str(row.get("role", "")).lower() == "equation") / len(row_refs)
        if row_refs else 0.0
    )
    if equation_ratio >= 0.4:
        batch_size = min(batch_size, 6)
        grid_cols = min(grid_cols, 2)
        grid_cell_width = max(grid_cell_width, 280)
        grid_cell_height = max(grid_cell_height, 180)

    workbook_bytes = build_alt_excel(synthetic_rows, "groq_grid_batch.docx", preview_images=synthetic_previews)
    images = extract_excel_images(workbook_bytes)
    if not images:
        return {"generated_count": 0, "error_count": len(row_refs), "first_error": "No preview images were available for Groq batching."}

    grid_files = make_grids(
        images,
        max_per_grid=batch_size,
        cols=grid_cols,
        cell_w=grid_cell_width,
        cell_h=grid_cell_height,
        margin=grid_margin,
    )
    generated_count = 0
    error_count = 0
    first_error = ""

    for grid_index, (_name, grid_bytes) in enumerate(grid_files):
        start = grid_index * batch_size
        batch_rows = row_refs[start : start + batch_size]
        if not batch_rows:
            continue
        try:
            optimized_grid_bytes, optimized_mime = optimize_groq_grid_image(grid_bytes, config)
            batch_bundle = request_groq_grid_alt_texts(
                optimized_grid_bytes,
                optimized_mime,
                batch_rows,
                start + 1,
                config,
                pdf_path,
                docx_path,
            )
        except Exception as exc:
            error_count += len(batch_rows)
            if not first_error:
                first_error = normalize_alt_text(str(exc))
            continue

        batch_alt_map = dict(batch_bundle.get("alt_map") or {})
        batch_ocr_map = dict(batch_bundle.get("ocr_map") or {})

        batch_generated = 0
        for offset, row in enumerate(batch_rows):
            displayed_number = start + offset + 1
            preview_entry = preview_refs[start + offset] if start + offset < len(preview_refs) else None
            role = str(row.get("role", "")).lower()
            ocr_text = normalize_alt_text(batch_ocr_map.get(displayed_number, "") or "")
            raw_alt = batch_alt_map.get(displayed_number, "")
            if role == "image" and isinstance(preview_entry, dict):
                image_bytes = preview_entry.get("bytes")
                if isinstance(image_bytes, (bytes, bytearray)) and image_bytes:
                    image_ext = str(preview_entry.get("ext", "png") or "png").lower()
                    image_mime = "image/jpeg" if image_ext in {"jpg", "jpeg"} else f"image/{image_ext}"
                    if "/" not in image_mime:
                        image_mime = "image/png"
                    cleaned_alt = request_groq_rescue_alt_text(
                        normalize_image_to_png(bytes(image_bytes)) or bytes(image_bytes),
                        image_mime,
                        row,
                        config,
                        ocr_text=ocr_text,
                    )
                else:
                    cleaned_alt = ""
            else:
                cleaned_alt = sanitize_groq_generated_alt(raw_alt, role, row.get("page"))
                if not is_meaningful_groq_alt_rewrite(cleaned_alt, role, ocr_text):
                    cleaned_alt, _ = generate_alt_text_for_row_with_groq(
                        row,
                        preview_entry,
                        config,
                        pdf_path,
                        docx_path=docx_path,
                    )
            if not cleaned_alt:
                continue
            row["generated_alt_text"] = cleaned_alt
            row["alt_text"] = cleaned_alt
            row["effective_alt_text"] = cleaned_alt
            row["alt_source"] = "generated_groq"
            batch_generated += 1
        generated_count += batch_generated
        error_count += max(0, len(batch_rows) - batch_generated)

    return {
        "generated_count": generated_count,
        "error_count": error_count,
        "first_error": first_error,
    }


def generate_alt_text_batches_with_claude(
    batch_candidates: list[dict],
    config: dict,
    pdf_path: Path | None,
    docx_path: Path | None = None,
) -> dict:
    if not batch_candidates:
        return {"generated_count": 0, "error_count": 0, "first_error": ""}

    from Altomizer.docx_tools import extract_excel_images, make_grids

    batch_size = max(1, min(int(config.get("grid_batch_size", 24) or 24), 24))
    grid_cols = max(1, min(int(config.get("grid_cols", 4) or 4), batch_size))
    grid_cell_width = max(96, int(config.get("grid_cell_width", 180) or 180))
    grid_cell_height = max(96, int(config.get("grid_cell_height", 140) or 140))
    grid_margin = max(4, int(config.get("grid_margin", 12) or 12))
    synthetic_rows: list[dict] = []
    synthetic_previews: dict[int, dict] = {}
    row_refs: list[dict] = []
    preview_refs: list[dict] = []

    for index, candidate in enumerate(batch_candidates):
        row = candidate.get("row")
        preview_entry = candidate.get("preview_entry")
        if not isinstance(row, dict) or not isinstance(preview_entry, dict):
            continue
        synthetic_rows.append(
            {
                "id": index,
                "type": row.get("type", "Item"),
                "role": row.get("role", ""),
                "source_part": row.get("source_part", ""),
                "page": row.get("page"),
                "alt_text": "",
                "existing_alt_text": "",
                "generated_alt_text": "",
            }
        )
        synthetic_previews[index] = preview_entry
        row_refs.append(row)
        preview_refs.append(preview_entry)

    if not synthetic_rows:
        return {"generated_count": 0, "error_count": 0, "first_error": ""}

    workbook_bytes = build_alt_excel(synthetic_rows, "claude_grid_batch.docx", preview_images=synthetic_previews)
    images = extract_excel_images(workbook_bytes)
    if not images:
        return {"generated_count": 0, "error_count": len(row_refs), "first_error": "No preview images were available for Claude batching."}

    grid_files = make_grids(
        images,
        max_per_grid=batch_size,
        cols=grid_cols,
        cell_w=grid_cell_width,
        cell_h=grid_cell_height,
        margin=grid_margin,
    )
    generated_count = 0
    error_count = 0
    first_error = ""

    for grid_index, (_name, grid_bytes) in enumerate(grid_files):
        start = grid_index * batch_size
        batch_rows = row_refs[start : start + batch_size]
        if not batch_rows:
            continue
        try:
            optimized_grid_bytes, optimized_mime = optimize_groq_grid_image(grid_bytes, config)
            batch_alt_map = request_claude_grid_alt_texts(
                optimized_grid_bytes,
                optimized_mime,
                batch_rows,
                start + 1,
                config,
                pdf_path,
                docx_path,
            )
        except Exception as exc:
            error_count += len(batch_rows)
            if not first_error:
                first_error = normalize_alt_text(str(exc))
            continue

        batch_generated = 0
        for offset, row in enumerate(batch_rows):
            displayed_number = start + offset + 1
            role = str(row.get("role", "")).lower()
            preview_entry = preview_refs[start + offset] if start + offset < len(preview_refs) else None
            cleaned_alt = sanitize_groq_generated_alt(batch_alt_map.get(displayed_number, ""), role, row.get("page"))
            if not cleaned_alt:
                cleaned_alt = request_claude_rescue_alt_text(preview_entry, row, config)
            if not cleaned_alt:
                continue
            row["generated_alt_text"] = cleaned_alt
            row["alt_text"] = cleaned_alt
            row["effective_alt_text"] = cleaned_alt
            row["alt_source"] = "generated_claude"
            batch_generated += 1
        generated_count += batch_generated
        error_count += max(0, len(batch_rows) - batch_generated)

    return {
        "generated_count": generated_count,
        "error_count": error_count,
        "first_error": first_error,
    }


def generate_alt_text_batches_with_openrouter(
    batch_candidates: list[dict],
    config: dict,
    pdf_path: Path | None,
    docx_path: Path | None = None,
) -> dict:
    if not batch_candidates:
        return {"generated_count": 0, "error_count": 0, "first_error": ""}

    from Altomizer.docx_tools import extract_excel_images, make_grids

    batch_size = max(1, min(int(config.get("grid_batch_size", 24) or 24), 24))
    grid_cols = max(1, min(int(config.get("grid_cols", 4) or 4), batch_size))
    grid_cell_width = max(96, int(config.get("grid_cell_width", 180) or 180))
    grid_cell_height = max(96, int(config.get("grid_cell_height", 140) or 140))
    grid_margin = max(4, int(config.get("grid_margin", 12) or 12))
    synthetic_rows: list[dict] = []
    synthetic_previews: dict[int, dict] = {}
    row_refs: list[dict] = []
    preview_refs: list[dict] = []

    for index, candidate in enumerate(batch_candidates):
        row = candidate.get("row")
        preview_entry = candidate.get("preview_entry")
        if not isinstance(row, dict) or not isinstance(preview_entry, dict):
            continue
        synthetic_rows.append(
            {
                "id": index,
                "type": row.get("type", "Item"),
                "role": row.get("role", ""),
                "source_part": row.get("source_part", ""),
                "page": row.get("page"),
                "alt_text": "",
                "existing_alt_text": "",
                "generated_alt_text": "",
            }
        )
        synthetic_previews[index] = preview_entry
        row_refs.append(row)
        preview_refs.append(preview_entry)

    if not synthetic_rows:
        return {"generated_count": 0, "error_count": 0, "first_error": ""}

    workbook_bytes = build_alt_excel(synthetic_rows, "openrouter_grid_batch.docx", preview_images=synthetic_previews)
    images = extract_excel_images(workbook_bytes)
    if not images:
        return {"generated_count": 0, "error_count": len(row_refs), "first_error": "No preview images were available for OpenRouter batching."}

    grid_files = make_grids(
        images,
        max_per_grid=batch_size,
        cols=grid_cols,
        cell_w=grid_cell_width,
        cell_h=grid_cell_height,
        margin=grid_margin,
    )
    generated_count = 0
    error_count = 0
    first_error = ""

    for grid_index, (_name, grid_bytes) in enumerate(grid_files):
        start = grid_index * batch_size
        batch_rows = row_refs[start : start + batch_size]
        if not batch_rows:
            continue
        try:
            optimized_grid_bytes, optimized_mime = optimize_groq_grid_image(grid_bytes, config)
            batch_alt_map = request_openrouter_grid_alt_texts(
                optimized_grid_bytes,
                optimized_mime,
                batch_rows,
                start + 1,
                config,
                pdf_path,
                docx_path,
            )
        except Exception as exc:
            error_count += len(batch_rows)
            if not first_error:
                first_error = normalize_alt_text(str(exc))
            continue

        batch_generated = 0
        for offset, row in enumerate(batch_rows):
            displayed_number = start + offset + 1
            role = str(row.get("role", "")).lower()
            preview_entry = preview_refs[start + offset] if start + offset < len(preview_refs) else None
            cleaned_alt = sanitize_groq_generated_alt(batch_alt_map.get(displayed_number, ""), role, row.get("page"))
            if not cleaned_alt:
                cleaned_alt = request_openrouter_rescue_alt_text(preview_entry, row, config)
            if not cleaned_alt:
                continue
            row["generated_alt_text"] = cleaned_alt
            row["alt_text"] = cleaned_alt
            row["effective_alt_text"] = cleaned_alt
            row["alt_source"] = "generated_openrouter"
            batch_generated += 1
        generated_count += batch_generated
        error_count += max(0, len(batch_rows) - batch_generated)

    return {
        "generated_count": generated_count,
        "error_count": error_count,
        "first_error": first_error,
    }


def generate_alt_text_batches_with_copilot(
    batch_candidates: list[dict],
    config: dict,
    pdf_path: Path | None,
    docx_path: Path | None = None,
) -> dict:
    if not batch_candidates:
        return {"generated_count": 0, "error_count": 0, "first_error": ""}

    from Altomizer.docx_tools import extract_excel_images, make_grids

    batch_size = max(1, min(int(config.get("grid_batch_size", 24) or 24), 24))
    grid_cols = max(1, min(int(config.get("grid_cols", 4) or 4), batch_size))
    grid_cell_width = max(96, int(config.get("grid_cell_width", 180) or 180))
    grid_cell_height = max(96, int(config.get("grid_cell_height", 140) or 140))
    grid_margin = max(4, int(config.get("grid_margin", 12) or 12))
    synthetic_rows: list[dict] = []
    synthetic_previews: dict[int, dict] = {}
    row_refs: list[dict] = []
    preview_refs: list[dict] = []

    for index, candidate in enumerate(batch_candidates):
        row = candidate.get("row")
        preview_entry = candidate.get("preview_entry")
        if not isinstance(row, dict) or not isinstance(preview_entry, dict):
            continue
        synthetic_rows.append(
            {
                "id": index,
                "type": row.get("type", "Item"),
                "role": row.get("role", ""),
                "source_part": row.get("source_part", ""),
                "page": row.get("page"),
                "alt_text": "",
                "existing_alt_text": "",
                "generated_alt_text": "",
            }
        )
        synthetic_previews[index] = preview_entry
        row_refs.append(row)
        preview_refs.append(preview_entry)

    if not synthetic_rows:
        return {"generated_count": 0, "error_count": 0, "first_error": ""}

    workbook_bytes = build_alt_excel(synthetic_rows, "copilot_grid_batch.docx", preview_images=synthetic_previews)
    images = extract_excel_images(workbook_bytes)
    if not images:
        return {"generated_count": 0, "error_count": len(row_refs), "first_error": "No preview images were available for Copilot batching."}

    grid_files = make_grids(
        images,
        max_per_grid=batch_size,
        cols=grid_cols,
        cell_w=grid_cell_width,
        cell_h=grid_cell_height,
        margin=grid_margin,
    )
    generated_count = 0
    error_count = 0
    first_error = ""

    for grid_index, (_name, grid_bytes) in enumerate(grid_files):
        start = grid_index * batch_size
        batch_rows = row_refs[start : start + batch_size]
        if not batch_rows:
            continue
        try:
            optimized_grid_bytes, optimized_mime = optimize_groq_grid_image(grid_bytes, config)
            batch_alt_map = request_copilot_grid_alt_texts(
                optimized_grid_bytes,
                optimized_mime,
                batch_rows,
                start + 1,
                config,
                pdf_path,
                docx_path,
            )
        except Exception as exc:
            error_count += len(batch_rows)
            if not first_error:
                first_error = normalize_alt_text(str(exc))
            continue

        batch_generated = 0
        for offset, row in enumerate(batch_rows):
            displayed_number = start + offset + 1
            role = str(row.get("role", "")).lower()
            preview_entry = preview_refs[start + offset] if start + offset < len(preview_refs) else None
            cleaned_alt = sanitize_groq_generated_alt(batch_alt_map.get(displayed_number, ""), role, row.get("page"))
            if not cleaned_alt:
                cleaned_alt = request_copilot_rescue_alt_text(preview_entry, row, config)
            if not cleaned_alt:
                continue
            row["generated_alt_text"] = cleaned_alt
            row["alt_text"] = cleaned_alt
            row["effective_alt_text"] = cleaned_alt
            row["alt_source"] = "generated_copilot"
            batch_generated += 1
        generated_count += batch_generated
        error_count += max(0, len(batch_rows) - batch_generated)

    return {
        "generated_count": generated_count,
        "error_count": error_count,
        "first_error": first_error,
    }


def generate_alt_text_batches_with_gemini(
    batch_candidates: list[dict],
    config: dict,
    pdf_path: Path | None,
    docx_path: Path | None = None,
) -> dict:
    if not batch_candidates:
        return {"generated_count": 0, "error_count": 0, "first_error": ""}

    from Altomizer.docx_tools import extract_excel_images, make_grids

    batch_size = max(1, min(int(config.get("grid_batch_size", 24) or 24), 24))
    grid_cols = max(1, min(int(config.get("grid_cols", 4) or 4), batch_size))
    grid_cell_width = max(96, int(config.get("grid_cell_width", 180) or 180))
    grid_cell_height = max(96, int(config.get("grid_cell_height", 140) or 140))
    grid_margin = max(4, int(config.get("grid_margin", 12) or 12))
    synthetic_rows: list[dict] = []
    synthetic_previews: dict[int, dict] = {}
    row_refs: list[dict] = []
    preview_refs: list[dict] = []

    for index, candidate in enumerate(batch_candidates):
        row = candidate.get("row")
        preview_entry = candidate.get("preview_entry")
        if not isinstance(row, dict) or not isinstance(preview_entry, dict):
            continue
        synthetic_rows.append(
            {
                "id": index,
                "type": row.get("type", "Item"),
                "role": row.get("role", ""),
                "source_part": row.get("source_part", ""),
                "page": row.get("page"),
                "alt_text": "",
                "existing_alt_text": "",
                "generated_alt_text": "",
            }
        )
        synthetic_previews[index] = preview_entry
        row_refs.append(row)
        preview_refs.append(preview_entry)

    if not synthetic_rows:
        return {"generated_count": 0, "error_count": 0, "first_error": ""}

    workbook_bytes = build_alt_excel(synthetic_rows, "gemini_grid_batch.docx", preview_images=synthetic_previews)
    images = extract_excel_images(workbook_bytes)
    if not images:
        return {"generated_count": 0, "error_count": len(row_refs), "first_error": "No preview images were available for Gemini batching."}

    grid_files = make_grids(
        images,
        max_per_grid=batch_size,
        cols=grid_cols,
        cell_w=grid_cell_width,
        cell_h=grid_cell_height,
        margin=grid_margin,
    )
    generated_count = 0
    error_count = 0
    first_error = ""

    for grid_index, (_name, grid_bytes) in enumerate(grid_files):
        start = grid_index * batch_size
        batch_rows = row_refs[start : start + batch_size]
        if not batch_rows:
            continue
        try:
            optimized_grid_bytes, optimized_mime = optimize_groq_grid_image(grid_bytes, config)
            batch_alt_map = request_gemini_grid_alt_texts(
                optimized_grid_bytes,
                optimized_mime,
                batch_rows,
                start + 1,
                config,
                pdf_path,
                docx_path,
            )
        except Exception as exc:
            error_count += len(batch_rows)
            if not first_error:
                first_error = normalize_alt_text(str(exc))
            continue

        batch_generated = 0
        for offset, row in enumerate(batch_rows):
            displayed_number = start + offset + 1
            role = str(row.get("role", "")).lower()
            preview_entry = preview_refs[start + offset] if start + offset < len(preview_refs) else None
            cleaned_alt = sanitize_groq_generated_alt(batch_alt_map.get(displayed_number, ""), role, row.get("page"))
            if not cleaned_alt:
                cleaned_alt = request_gemini_rescue_alt_text(preview_entry, row, config)
            if not cleaned_alt:
                continue
            row["generated_alt_text"] = cleaned_alt
            row["alt_text"] = cleaned_alt
            row["effective_alt_text"] = cleaned_alt
            row["alt_source"] = "generated_gemini"
            batch_generated += 1
        generated_count += batch_generated
        error_count += max(0, len(batch_rows) - batch_generated)

    return {
        "generated_count": generated_count,
        "error_count": error_count,
        "first_error": first_error,
    }


def generate_alt_text_for_row_with_groq(
    row: dict,
    preview_entry: dict | None,
    config: dict,
    pdf_path: Path | None,
    docx_path: Path | None = None,
    *,
    key_cursor: int = 0,
) -> tuple[str, int]:
    if not isinstance(preview_entry, dict):
        return ("", key_cursor)

    image_bytes = preview_entry.get("bytes")
    if not isinstance(image_bytes, (bytes, bytearray)) or not image_bytes:
        return ("", key_cursor)
    normalized_image_bytes = normalize_image_to_png(bytes(image_bytes)) or bytes(image_bytes)

    image_ext = "png" if normalized_image_bytes != bytes(image_bytes) else str(preview_entry.get("ext", "png") or "png").lower()
    image_mime = "image/jpeg" if image_ext in {"jpg", "jpeg"} else f"image/{image_ext}"
    if "/" not in image_mime:
        image_mime = "image/png"

    generated_alt = request_groq_ocr_then_rewrite_alt_text(
        normalized_image_bytes,
        image_mime,
        row,
        config,
        pdf_path,
        docx_path=docx_path,
    )
    return (generated_alt, get_groq_stage_cursor(config, "rewrite"))


def build_groq_alt_prompt(role: str, hint_text: str = "") -> str:
    cleaned_hint = normalize_alt_text(hint_text)
    semantic_hint = groq_rewrite_semantic_hint(role, cleaned_hint)
    if role == "equation":
        prompt = """
        You write the final accessibility alt text for an educational equation or mathematical expression.
        The alt text must be easy to understand.
        All numericals must be written as words.
        If a decimal number is visible, speak the decimal point as "point". Example: 6.5 becomes "six point five".
        Never start with "Equation shows", "Equation showing", "Image shows", or "Image showing".
        Write the exact final alt text only, as one detailed screen-reader-ready sentence in natural spoken math language.
        Do not read punctuation names such as comma, brace, bracket, or parenthesis aloud unless that punctuation itself is the key meaning.
        If the content is a set or sequence like {1, 2, 3, ...}, describe it as a set or sequence and end with "and so on".
        If the content is a relation like 2 >= 2, write it as "two is greater than or equal to two."
        If the expression is partially unclear, stay conservative and describe only what is visible.
        If the OCR transcription contains blanks written as ___, say "blank" for each blank and preserve the visible operator order exactly.
        Do not add, remove, or reorder operators.
        """
    else:
        prompt = """
        You write the final accessibility alt text for an educational image.
        The alt text must be easy to understand.
        All numericals must be written as words.
        If a decimal number is visible, speak the decimal point as "point". Example: 57.5 becomes "fifty seven point five".
        Never start with "Equation shows", "Equation showing", "Image shows", or "Image showing".
        Do not say "Here is the alt text", "Screen-reader-friendly alt text", "ALT text:", or any similar introduction.
        Write only the exact final alt text.
        Use four to twelve detailed sentences when the image is a table, graph, chart, or diagram.
        Use one or two sentences for simpler images.
        Focus on the key visual content, relationship, labels, layout, and trend when relevant.
        Make the description specific enough that a blind reader can understand what the image is and what matters in it.
        For graphs, charts, and diagrams, be highly detailed. Name the chart type, title, axes, scales, tick marks, units, legend entries, and visible plotted values or labeled intervals.
        If the image is a histogram, bar chart, line graph, scatter plot, or frequency chart, list each visible interval, category, bar, bin, or series value when it can be read reliably.
        Do not mention surrounding page instructions, questions, answers, or chapter text.
        """

    prompt = textwrap.dedent(prompt).strip()
    if semantic_hint:
        prompt += f"\nSemantic hint: {semantic_hint[:260]}"
    if cleaned_hint:
        prompt += f"\nOCR transcription: {cleaned_hint[:260]}"
    return prompt


def build_claude_grid_alt_prompt(rows: list[dict], start_number: int) -> str:
    number_list = ", ".join(str(start_number + offset) for offset in range(len(rows)))
    role_lines = []
    for offset, row in enumerate(rows):
        item_number = start_number + offset
        role = str(row.get("role", "") or "image").lower()
        page = row.get("page")
        if page:
            role_lines.append(f"{item_number}: role={role} page={page}")
        else:
            role_lines.append(f"{item_number}: role={role}")

    strict_prompt = """
    Detect EVERY individual bordered sub-image.
    Match each detected sub-image to the correct numbered row in the manifest grid.
    Fill ONLY the Alt Text column.
    Preserve ALL existing values exactly as they are:
    Item IDs
    Source Labels
    Page values
    Image File values
    Status values
    Never delete, reorder, merge, or skip rows.

    Reading order must always be:
    Top-to-bottom
    Left-to-right

    Border handling rules:
    Watch borders extremely carefully.
    Never combine two neighboring images into one description.
    Never let text from adjacent boxes leak into another image’s alt text.
    Some bordered images may be very small — inspect carefully before reading.

    Alt text generation rules:
    Write ONLY what is visually present in the image.
    Do NOT add assumptions.
    Do NOT expand abbreviations unless visibly written.
    Do NOT insert punctuation, commas, parentheses, brackets, symbols, or words unless they actually appear in the image.
    If a decimal number is visible, say the decimal point as "point". Example: 57.5 becomes "fifty seven point five".
    If parentheses are visually present, write:
    open parenthesis
    close parenthesis
    If parentheses are NOT visually present, NEVER mention parenthesis.
    Do NOT add extra words for readability.
    Do NOT write introductory phrases such as:
    "the image shows"
    "the equation shows"
    "this graph contains"
    "symbol representing"

    Equation and formula rules:
    Transcribe exactly in linear spoken-text format.
    Use only visually present mathematical structure.
    Use terms like:
    plus
    minus
    times
    divided by
    equals
    superscript
    subscript
    open parenthesis
    close parenthesis
    Preserve exact ordering and grouping from the image.
    Never invent brackets or implied grouping.

    Character accuracy is critical:
    Distinguish carefully between:
    x vs X
    z vs Z
    z vs 2
    E vs e
    mu vs nu
    O vs 0
    l vs 1
    Never substitute similar-looking characters.
    Read symbols carefully before transcribing.

    Examples:
    Correct:
    mu subscript x equals 1 times open parenthesis 0.42 close parenthesis plus 2 times open parenthesis 0.18 close parenthesis

    Incorrect:
    mu subscript x equals 1 times 0.42 plus 2 times 0.18
    (parentheses were removed)

    Incorrect:
    The equation shows mu subscript x equals ...
    (introductory wording added)

    Tables/spreadsheets:
    Include ALL visible:
    headers
    row labels
    values
    symbols
    percentages
    notes
    When the tile is mainly a table or spreadsheet, explain it in a bit more detail than other images.
    State the table structure first, including the column headers and row group labels when visible.
    Summarize the key values, comparisons, totals, or category differences that matter most.
    If the table is dense, do not read every cell mechanically; give the structure plus the most important visible entries and patterns.
    If a total, sum, average, or final row is visible, mention it explicitly.

    Graphs/diagrams:
    Include:
    title
    axis labels
    legends
    arrows
    plotted labels
    annotations
    visible text
    When the tile is a graph, chart, or diagram, explain it in detail.
    Name the chart type first, then describe axes, units, legend entries, and the overall pattern or comparison.
    Mention peaks, lows, intersections, ordering, direction of change, and any standout labeled values that are visible.
    Include the chart title, axis scales, and tick-mark intervals when visible.
    If it is a histogram, bar chart, line graph, scatter plot, or other frequency/data chart, list every visible bin, interval, category, bar, or plotted series value that can be read reliably.
    If exact frequencies or values are visible, include them explicitly.

    Icons/symbols:
    Describe concisely.
    Example:
    Plus sign icon
    Warning triangle icon

    Special cases:
    If the image box visibly contains the word:
    none
    then write exactly:
    none
    Never skip boxes containing "none".
    If image is decorative only:
    Decorative image.
    If unreadable:
    [Unclear image - needs manual review]

    Final quality rules:
    Accuracy is more important than speed.
    Read every image carefully before writing alt text.
    Never hallucinate missing content.
    Never autocorrect symbols or equations.
    Never infer hidden mathematical structure.
    Output must match the image exactly.
    """.strip()

    response_format = """
    Return JSON only with this exact shape:
    {{"items":[{{"number":1,"alt_text":"..."}}]}}

    Additional output rules:
    - Return one item for every tile number in this grid: {numbers}
    - The numbered tiles already represent the manifest row order.
    - Use each tile number exactly once.
    - Do not omit any tile. If a tile is hard to read, still return it and use [Unclear image - needs manual review].
    - Do not include any intro such as "Here is the alt text" or "Screen-reader-friendly alt text".
    - Fill only alt_text values.
    - No markdown.
    - No code fences.
    - No explanation outside the JSON.
    """.strip().format(numbers=number_list)

    return textwrap.dedent(
        f"""
        You are filling Alt Text for a numbered grid image created from manifest preview rows.
        Each numbered tile corresponds to one manifest row in order.
        Tile numbers in this grid: {number_list}
        Row map:
        {chr(10).join(role_lines)}

        {strict_prompt}

        {response_format}
        """
    ).strip()


def request_claude_grid_alt_texts(
    grid_bytes: bytes,
    grid_mime: str,
    rows: list[dict],
    start_number: int,
    config: dict,
    pdf_path: Path | None,
    docx_path: Path | None = None,
) -> dict[int, str]:
    del pdf_path, docx_path
    prompt = build_claude_grid_alt_prompt(rows, start_number)
    raw = request_claude_vision_text(
        grid_bytes,
        grid_mime,
        prompt,
        config,
        max_tokens=int(config.get("grid_max_tokens", config.get("max_tokens", 1024)) or 1024),
    )
    return parse_groq_grid_alt_response(raw)


def request_openrouter_grid_alt_texts(
    grid_bytes: bytes,
    grid_mime: str,
    rows: list[dict],
    start_number: int,
    config: dict,
    pdf_path: Path | None,
    docx_path: Path | None = None,
) -> dict[int, str]:
    del pdf_path, docx_path
    prompt = build_claude_grid_alt_prompt(rows, start_number)
    raw = request_openrouter_vision_text(
        grid_bytes,
        grid_mime,
        prompt,
        config,
        max_completion_tokens=int(config.get("grid_max_tokens", config.get("max_tokens", 1024)) or 1024),
    )
    return parse_groq_grid_alt_response(raw)


def request_copilot_grid_alt_texts(
    grid_bytes: bytes,
    grid_mime: str,
    rows: list[dict],
    start_number: int,
    config: dict,
    pdf_path: Path | None,
    docx_path: Path | None = None,
) -> dict[int, str]:
    del pdf_path, docx_path
    prompt = build_claude_grid_alt_prompt(rows, start_number)
    raw = request_copilot_vision_text(
        grid_bytes,
        grid_mime,
        prompt,
        config,
        max_tokens=int(config.get("grid_max_tokens", config.get("max_tokens", 1024)) or 1024),
    )
    return parse_groq_grid_alt_response(raw)


def request_gemini_grid_alt_texts(
    grid_bytes: bytes,
    grid_mime: str,
    rows: list[dict],
    start_number: int,
    config: dict,
    pdf_path: Path | None,
    docx_path: Path | None = None,
) -> dict[int, str]:
    del pdf_path, docx_path
    number_list = ", ".join(str(start_number + offset) for offset in range(len(rows)))
    equation_numbers = ", ".join(
        str(start_number + offset)
        for offset, row in enumerate(rows)
        if str(row.get("role", "")).lower() == "equation"
    ) or "none"

    ocr_map = request_gemini_grid_ocr_texts(
        grid_bytes,
        grid_mime,
        rows,
        start_number,
        number_list,
        equation_numbers,
        config,
    )
    if not ocr_map:
        return {}

    prompt = build_groq_grid_rewrite_prompt(rows, start_number, ocr_map)
    raw = request_gemini_generate_content(
        prompt,
        config,
        max_output_tokens=max(
            240,
            min(
                int(config.get("grid_rewrite_max_tokens", config.get("grid_max_tokens", config.get("max_tokens", 1024))) or 1024),
                72 * max(1, len(rows)),
            ),
        ),
    )
    parsed_alt_map = parse_groq_grid_alt_response(raw)
    filtered_alt_map: dict[int, str] = {}
    for offset, row in enumerate(rows):
        number = start_number + offset
        candidate = normalize_alt_text(parsed_alt_map.get(number, "") or "")
        if not candidate or looks_like_filename_metadata_text(candidate):
            continue
        role = str(row.get("role", "")).lower() or "image"
        ocr_text = normalize_alt_text(ocr_map.get(number, "") or "")
        if ocr_text and not is_meaningful_groq_alt_rewrite(candidate, role, ocr_text):
            continue
        filtered_alt_map[number] = candidate
    return filtered_alt_map


def request_gemini_grid_ocr_texts(
    grid_bytes: bytes,
    grid_mime: str,
    rows: list[dict],
    start_number: int,
    number_list: str,
    equation_numbers: str,
    config: dict,
) -> dict[int, str]:
    prompt = textwrap.dedent(
        f"""
        Read the numbered grid image and transcribe the visible content of each tile.
        Tile numbers present in this grid: {number_list}
        Equation tile numbers: {equation_numbers}
        Return JSON only with this exact shape:
        {{"items":[{{"number":1,"ocr_text":"..."}}]}}
        Rules:
        - Return one item for every listed tile number.
        - Use the tile number exactly as shown.
        - For equation tiles, transcribe the visible math as exactly as possible and keep the math symbols.
        - If a fill-in blank is visible, transcribe it as exactly three underscores: ___
        - Preserve the visible operator order exactly and do not invent missing symbols.
        - For non-equation tiles, transcribe the core visible content, labels, chart text, or subject exactly as seen.
        - Ignore workbook labels, file names, row-column markers, and surrounding interface text.
        - If a tile only appears to show workbook metadata or a file name, return an empty string for that tile.
        - Do not write accessibility alt text yet.
        - No markdown, no code fences, and no explanation outside the JSON.
        """
    ).strip()

    raw = request_gemini_vision_text(
        grid_bytes,
        grid_mime,
        prompt,
        config,
        max_output_tokens=max(
            160,
            min(
                int(config.get("grid_ocr_max_tokens", config.get("grid_max_tokens", config.get("max_tokens", 1024))) or 1024),
                32 * max(1, len(rows)),
            ),
        ),
    )
    parsed_ocr_map = parse_groq_grid_ocr_response(raw)
    sanitized_ocr_map: dict[int, str] = {}
    for offset, row in enumerate(rows):
        number = start_number + offset
        raw_text = normalize_alt_text(parsed_ocr_map.get(number, "") or "")
        if looks_like_filename_metadata_text(raw_text):
            continue
        cleaned_text = sanitize_groq_ocr_text(raw_text, str(row.get("role", "")).lower() or "image")
        if not cleaned_text or looks_like_filename_metadata_text(cleaned_text):
            continue
        sanitized_ocr_map[number] = cleaned_text
    return sanitized_ocr_map


def request_groq_grid_alt_texts(
    grid_bytes: bytes,
    grid_mime: str,
    rows: list[dict],
    start_number: int,
    config: dict,
    pdf_path: Path | None,
    docx_path: Path | None = None,
) -> dict[str, dict[int, str]]:
    number_list = ", ".join(str(start_number + offset) for offset in range(len(rows)))
    equation_numbers = ", ".join(
        str(start_number + offset)
        for offset, row in enumerate(rows)
        if str(row.get("role", "")).lower() == "equation"
    ) or "none"

    ocr_map = request_groq_grid_ocr_texts(
        grid_bytes,
        grid_mime,
        rows,
        start_number,
        number_list,
        equation_numbers,
        config,
    )
    if not ocr_map:
        return {"alt_map": {}, "ocr_map": {}}

    prompt = build_groq_grid_rewrite_prompt(rows, start_number, ocr_map)

    max_tokens = max(
        240,
        min(
            int(config.get("grid_rewrite_max_completion_tokens", config.get("grid_max_completion_tokens", 420)) or 420),
            72 * max(1, len(rows)),
        ),
    )
    rewrite_config = build_groq_stage_config(config, "rewrite")
    rewrite_cursor = get_groq_stage_cursor(config, "rewrite")
    raw = ""
    try:
        raw, next_cursor = request_groq_text_completion(
            prompt,
            rewrite_config,
            max_completion_tokens=max_tokens,
            response_format=build_groq_grid_json_schema(),
            key_cursor=rewrite_cursor,
        )
    except RuntimeError:
        raw, next_cursor = request_groq_text_completion(
            prompt,
            rewrite_config,
            max_completion_tokens=max_tokens,
            response_format={"type": "json_object"},
            key_cursor=rewrite_cursor,
        )
    set_groq_stage_cursor(config, "rewrite", next_cursor)
    return {
        "alt_map": parse_groq_grid_alt_response(raw),
        "ocr_map": ocr_map,
    }


def request_groq_grid_ocr_texts(
    grid_bytes: bytes,
    grid_mime: str,
    rows: list[dict],
    start_number: int,
    number_list: str,
    equation_numbers: str,
    config: dict,
) -> dict[int, str]:
    prompt = textwrap.dedent(
        f"""
        Read the numbered grid image and transcribe the visible content of each tile.
        Tile numbers present in this grid: {number_list}
        Equation tile numbers: {equation_numbers}
        Return JSON only with this exact shape:
        {{"items":[{{"number":1,"ocr_text":"..."}}]}}
        Rules:
        - Return one item for every listed tile number.
        - Use the tile number exactly as shown.
        - For equation tiles, transcribe the visible math as exactly as possible and keep the math symbols.
        - If a fill-in blank is visible, transcribe it as exactly three underscores: ___
        - Preserve the visible operator order exactly and do not invent missing symbols.
        - For non-equation tiles, transcribe the core visible content, labels, chart text, or subject exactly as seen.
        - Do not write accessibility alt text yet.
        - Ignore workbook labels, row-column markers, and surrounding interface text.
        - No markdown, no code fences, and no explanation outside the JSON.
        """
    ).strip()

    max_tokens = max(
        160,
        min(
            int(config.get("grid_ocr_max_completion_tokens", config.get("grid_max_completion_tokens", 420)) or 420),
            32 * max(1, len(rows)),
        ),
    )
    ocr_config = build_groq_stage_config(config, "ocr")
    ocr_cursor = get_groq_stage_cursor(config, "ocr")
    raw = ""
    try:
        raw, next_cursor = request_groq_vision_alt_text(
            grid_bytes,
            grid_mime,
            prompt,
            ocr_config,
            max_completion_tokens=max_tokens,
            response_format=build_groq_grid_ocr_json_schema(),
            key_cursor=ocr_cursor,
        )
    except RuntimeError:
        raw, next_cursor = request_groq_vision_alt_text(
            grid_bytes,
            grid_mime,
            prompt,
            ocr_config,
            max_completion_tokens=max_tokens,
            response_format={"type": "json_object"},
            key_cursor=ocr_cursor,
        )
    set_groq_stage_cursor(config, "ocr", next_cursor)
    return parse_groq_grid_ocr_response(raw)


def normalize_rewrite_source_text(text: str) -> str:
    working = compact_ocr_text(text)
    if not working:
        return ""
    replacements = (
        (r"\bcomma\b", ","),
        (r"\bellipsis\b", "..."),
        (r"\bopen brace\b", "{"),
        (r"\bclose brace\b", "}"),
        (r"\bopen bracket\b", "["),
        (r"\bclose bracket\b", "]"),
        (r"\bopen parenthesis\b", "("),
        (r"\bclose parenthesis\b", ")"),
    )
    for pattern, replacement in replacements:
        working = re.sub(pattern, replacement, working, flags=re.IGNORECASE)
    working = re.sub(r"\s+", " ", working).strip()
    return working


def groq_rewrite_semantic_hint(role: str, ocr_text: str) -> str:
    normalized = normalize_rewrite_source_text(ocr_text)
    if not normalized:
        return ""
    if role != "equation":
        return "Describe the subject, labels, structure, and the key visual takeaway."

    has_ellipsis = "..." in normalized or "…" in normalized
    has_braces = any(marker in normalized for marker in ("{", "}", "[", "]"))
    if has_ellipsis and has_braces:
        return "Likely set notation with an ellipsis. Describe it as a set or collection, not by reading commas or braces aloud."
    if has_ellipsis and "," in normalized:
        return "Likely a sequence with an ellipsis. Describe it as a sequence ending with 'and so on', not by reading commas aloud."
    if any(marker in normalized for marker in (">=", "<=", "!=", "<>", "=", "+", "-", "/", "^", "≤", "≥")):
        return "Describe the mathematical relationship in natural spoken math, not as raw symbols."
    return "Explain the mathematical object in clear spoken language."


def build_groq_grid_json_schema() -> dict:
    return {
        "type": "json_schema",
        "json_schema": {
            "name": "grid_alt_batch",
            "strict": False,
            "schema": {
                "type": "object",
                "properties": {
                    "items": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "number": {"type": "integer"},
                                "alt_text": {"type": "string"},
                            },
                            "required": ["number", "alt_text"],
                            "additionalProperties": False,
                        },
                    }
                },
                "required": ["items"],
                "additionalProperties": False,
            },
        },
    }


def build_groq_grid_ocr_json_schema() -> dict:
    return {
        "type": "json_schema",
        "json_schema": {
            "name": "grid_ocr_batch",
            "strict": False,
            "schema": {
                "type": "object",
                "properties": {
                    "items": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "number": {"type": "integer"},
                                "ocr_text": {"type": "string"},
                            },
                            "required": ["number", "ocr_text"],
                            "additionalProperties": False,
                        },
                    }
                },
                "required": ["items"],
                "additionalProperties": False,
            },
        },
    }


def build_groq_grid_rewrite_prompt(rows: list[dict], start_number: int, ocr_map: dict[int, str]) -> str:
    item_lines = []
    for offset, row in enumerate(rows):
        number = start_number + offset
        role = str(row.get("role", "")).lower() or "image"
        ocr_text = normalize_alt_text(ocr_map.get(number, "") or "")
        semantic_hint = groq_rewrite_semantic_hint(role, ocr_text)
        item_lines.append(
            f"{number} | role={role} | ocr_text={json.dumps(ocr_text, ensure_ascii=True)} | "
            f"semantic_hint={json.dumps(semantic_hint, ensure_ascii=True)}"
        )
    items_block = "\n".join(item_lines)

    return textwrap.dedent(
        f"""
        You are writing the final screen-reader alt text from OCR transcriptions.
        The given alt text for the equation and image must be easy to understand.
        All numericals must be written as words in the alt.
        If a decimal number is visible, speak the decimal point as "point". Example: 57.5 becomes "fifty seven point five".
        Example: 2+2=5 becomes "two plus two equals five".
        The alt must never start with "Equation shows", "Images shows", "Equation showing", "Image shows", or "Image showing".
        Do not read punctuation names such as comma, brace, bracket, or parenthesis aloud unless that punctuation itself is the key meaning.
        If an OCR item is set notation like {{1, 2, 3, ...}}, write "Set containing one, two, three, and so on."
        If an OCR item is a sequence like 1, 2, 3, ..., write "Sequence one, two, three, and so on."
        If an OCR item is an image, describe what it is in enough detail for a screen reader user to understand the content and what matters.
        Write the exact alt to be written.
        Return JSON only with this exact shape:
        {{"items":[{{"number":1,"alt_text":"..."}}]}}
        Rules:
        - Return one item for every listed number.
        - Keep each alt text concise for simple images, but allow four to twelve detailed sentences for tables, graphs, charts, and diagrams.
        - For equations, convert symbols into clear spoken math and describe the mathematical object or pattern, not just the literal characters.
        - For ordinary images, write a specific content description with no filler intro.
        - For histograms, bar charts, line graphs, scatter plots, and tables, include the title, axes or headers, intervals or categories, and each visible frequency or value when readable.
        - Ignore workbook labels, row-column markers, and interface chrome.
        - No markdown, no code fences, and no explanation outside the JSON.
        Items:
        {items_block}
        """
    ).strip()


def groq_row_hint_text(row: dict, pdf_path: Path | None, docx_path: Path | None = None) -> str:
    role = str(row.get("role", "")).lower()
    if role == "equation":
        page_number = row.get("preview_page")
        bbox = row.get("preview_bbox")
        if isinstance(pdf_path, Path) and isinstance(page_number, int) and isinstance(bbox, dict):
            hint_text = extract_best_equation_text(
                pdf_path,
                page_number,
                bbox,
                docx_path=docx_path,
                media_target=row.get("media_target"),
                ole_target=row.get("ole_target"),
                viewport_crop=row.get("viewport_crop"),
                display_width_pt=row.get("display_width_pt"),
                display_height_pt=row.get("display_height_pt"),
            )
            if hint_text:
                return clean_equation_candidate(hint_text)
        return clean_equation_candidate(row.get("label", "") or "")

    preview_entry = build_alt_preview_entry(row, docx_path, pdf_path)
    if not isinstance(preview_entry, dict):
        return ""
    image_bytes = preview_entry.get("bytes")
    if not isinstance(image_bytes, (bytes, bytearray)) or not image_bytes:
        return ""
    normalized_image_bytes = normalize_image_to_png(bytes(image_bytes)) or bytes(image_bytes)
    return clean_image_ocr_candidate(run_tesseract_on_png_bytes(normalized_image_bytes, psm_modes=("11", "6")))


def parse_groq_grid_items_response(text: str, value_field: str, error_message: str) -> dict[int, str]:
    payload = extract_json_payload_from_text(text)
    if payload is None:
        fallback = parse_loose_numbered_item_lines(text, value_field)
        if fallback:
            return fallback
        raise RuntimeError(error_message)

    items = []
    if isinstance(payload, dict):
        raw_items = payload.get("items")
        if isinstance(raw_items, list):
            items = raw_items
        elif all(str(key).strip().isdigit() for key in payload.keys()):
            for key, value in payload.items():
                items.append({"number": key, value_field: value})
        else:
            numbered_map = payload.get("alts") or payload.get("results") or payload.get("data") or payload.get("ocr")
            if isinstance(numbered_map, dict):
                for key, value in numbered_map.items():
                    items.append({"number": key, value_field: value})
    elif isinstance(payload, list):
        items = payload

    result: dict[int, str] = {}
    for item in items:
        if not isinstance(item, dict):
            continue
        number = item.get("number")
        item_text = normalize_alt_text(str(item.get(value_field, "") or ""))
        if not isinstance(number, int):
            try:
                number = int(number)
            except (TypeError, ValueError):
                continue
        if number <= 0:
            continue
        result[number] = item_text
    if result:
        return result

    fallback = parse_loose_numbered_item_lines(text, value_field)
    if fallback:
        return fallback
    return result


def parse_groq_grid_alt_response(text: str) -> dict[int, str]:
    return parse_groq_grid_items_response(
        text,
        "alt_text",
        "Groq did not return valid JSON for the grid ALT batch.",
    )


def parse_groq_grid_ocr_response(text: str) -> dict[int, str]:
    return parse_groq_grid_items_response(
        text,
        "ocr_text",
        "Groq did not return valid JSON for the grid OCR batch.",
    )


def extract_json_payload_from_text(text: str) -> dict | list | None:
    cleaned = str(text or "").strip()
    if not cleaned:
        return None
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*```$", "", cleaned)

    candidates = [cleaned]
    object_match = re.search(r"\{.*\}", cleaned, flags=re.DOTALL)
    if object_match:
        candidates.append(object_match.group(0))
    array_match = re.search(r"\[.*\]", cleaned, flags=re.DOTALL)
    if array_match:
        candidates.append(array_match.group(0))

    seen = set()
    for candidate in candidates:
        candidate = candidate.strip()
        if not candidate or candidate in seen:
            continue
        seen.add(candidate)
        try:
            return json.loads(candidate)
        except ValueError:
            continue
    return None


def parse_loose_numbered_item_lines(text: str, value_field: str) -> dict[int, str]:
    cleaned = str(text or "").strip()
    if not cleaned:
        return {}
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*```$", "", cleaned)
    pattern = re.compile(
        r"(?ms)^\s*(?:number\s*)?(\d{1,3})\s*(?:[.)\-:|]|=>)\s*(.+?)(?=^\s*(?:number\s*)?\d{1,3}\s*(?:[.)\-:|]|=>)|\Z)"
    )
    result: dict[int, str] = {}
    for match in pattern.finditer(cleaned):
        try:
            number = int(match.group(1))
        except ValueError:
            continue
        item_text = normalize_alt_text(match.group(2))
        prefixes = [
            rf"^(?:{re.escape(value_field)}|description)\s*[:=\-]\s*",
        ]
        if value_field == "alt_text":
            prefixes.append(r"^(?:alt(?:_?text)?)\s*[:=\-]\s*")
        elif value_field == "ocr_text":
            prefixes.append(r"^(?:ocr(?:_?text)?|transcription)\s*[:=\-]\s*")
        for prefix in prefixes:
            item_text = re.sub(prefix, "", item_text, flags=re.IGNORECASE)
        if number > 0 and item_text:
            result[number] = item_text
    return result


def optimize_groq_grid_image(grid_bytes: bytes, config: dict) -> tuple[bytes, str]:
    if not grid_bytes:
        return (b"", "image/png")

    max_side = max(512, int(config.get("grid_image_max_side", 1200) or 1200))
    jpeg_quality = max(35, min(int(config.get("grid_image_jpeg_quality", 68) or 68), 90))
    try:
        with Image.open(io.BytesIO(grid_bytes)) as image:
            image = image.convert("RGB")
            width, height = image.size
            scale = min(max_side / max(1, width), max_side / max(1, height), 1.0)
            if scale < 1.0:
                resized = image.resize(
                    (max(1, int(width * scale)), max(1, int(height * scale))),
                    Image.Resampling.LANCZOS,
                )
            else:
                resized = image
            output = io.BytesIO()
            resized.save(output, format="JPEG", quality=jpeg_quality, optimize=True)
            return (output.getvalue(), "image/jpeg")
    except Exception:
        return (grid_bytes, "image/png")


def groq_stage_api_keys(config: dict, stage: str) -> list[str]:
    if stage == "ocr":
        keys = list(config.get("ocr_api_keys") or [])
    elif stage == "rewrite":
        keys = list(config.get("rewrite_api_keys") or [])
    else:
        keys = list(config.get("api_keys") or [])
    if keys:
        return keys
    fallback_key = str(config.get("api_key") or "").strip()
    return [fallback_key] if fallback_key else []


def build_groq_stage_config(config: dict, stage: str) -> dict:
    stage_keys = groq_stage_api_keys(config, stage)
    stage_config = dict(config)
    stage_config["api_keys"] = stage_keys
    stage_config["api_key"] = stage_keys[0] if stage_keys else ""
    return stage_config


def get_groq_stage_cursor(config: dict, stage: str) -> int:
    return max(0, int(config.get(f"_{stage}_key_cursor", 0) or 0))


def set_groq_stage_cursor(config: dict, stage: str, value: int) -> None:
    config[f"_{stage}_key_cursor"] = max(0, int(value or 0))


def request_claude_message(
    prompt: str,
    config: dict,
    *,
    content_blocks: list[dict] | None = None,
    max_tokens: int | None = None,
) -> str:
    api_keys = [str(key or "").strip() for key in config.get("api_keys") or [] if str(key or "").strip()]
    if not api_keys:
        api_key = str(config.get("api_key") or "").strip()
        if api_key:
            api_keys = [api_key]
    if not api_keys:
        raise RuntimeError("No Claude API key is configured.")
    source_names = [str(source or "").strip() for source in config.get("api_key_sources") or []]

    content = list(content_blocks or [])
    content.append({"type": "text", "text": prompt})
    payload = {
        "model": str(config.get("model") or "claude-sonnet-4-6"),
        "max_tokens": int(max_tokens or config.get("max_tokens", 1024)),
        "temperature": 0,
        "messages": [
            {
                "role": "user",
                "content": content,
            }
        ],
    }

    response = None
    auth_errors: list[str] = []
    for key_index, api_key in enumerate(api_keys):
        response = requests.post(
            f"{str(config.get('base_url')).rstrip('/')}/messages",
            headers={
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json=payload,
            timeout=int(config.get("timeout_seconds", 45)),
        )
        if response.status_code == 200:
            break

        error_message = _claude_error_message(response)
        source_name = source_names[key_index] if key_index < len(source_names) and source_names[key_index] else f"configured key #{key_index + 1}"
        if is_claude_auth_error(response, error_message) and key_index + 1 < len(api_keys):
            auth_errors.append(f"{source_name}: {error_message}")
            continue
        if auth_errors:
            tried_sources = ", ".join(source_names[: len(api_keys)] or [f"{len(api_keys)} configured key(s)"])
            raise RuntimeError(f"{error_message} Tried Claude key source(s): {tried_sources}.")
        raise RuntimeError(error_message)

    if response is None or response.status_code != 200:
        tried_sources = ", ".join(source_names[: len(api_keys)] or [f"{len(api_keys)} configured key(s)"])
        raise RuntimeError(f"Claude request failed. Tried Claude key source(s): {tried_sources}.")

    response_payload = _safe_json(response)
    if not isinstance(response_payload, dict):
        return ""

    blocks = response_payload.get("content")
    if not isinstance(blocks, list):
        return ""

    parts: list[str] = []
    for block in blocks:
        if isinstance(block, dict) and str(block.get("type") or "") == "text":
            text = block.get("text")
            if isinstance(text, str) and text.strip():
                parts.append(text)
    return "\n".join(parts).strip()


def extract_chat_message_text_content(content: object) -> str:
    if isinstance(content, str):
        return content.strip()
    if not isinstance(content, list):
        return ""

    parts: list[str] = []
    for block in content:
        if not isinstance(block, dict):
            continue
        if str(block.get("type") or "") != "text":
            continue
        text = block.get("text")
        if isinstance(text, str) and text.strip():
            parts.append(text.strip())
    return "\n".join(parts).strip()


def request_claude_vision_text(
    image_bytes: bytes,
    image_mime: str,
    prompt: str,
    config: dict,
    *,
    max_tokens: int | None = None,
) -> str:
    if not image_bytes:
        return ""

    encoded_image = base64.b64encode(image_bytes).decode("ascii")
    return request_claude_message(
        prompt,
        config,
        content_blocks=[
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": image_mime,
                    "data": encoded_image,
                },
            }
        ],
        max_tokens=max_tokens,
    )


def request_openrouter_chat_completion(
    prompt: str,
    config: dict,
    *,
    content_blocks: list[dict] | None = None,
    max_completion_tokens: int | None = None,
) -> str:
    api_key = str(config.get("api_key") or "").strip()
    if not api_key:
        raise RuntimeError("No OpenRouter API key is configured.")

    content = [{"type": "text", "text": prompt}]
    if isinstance(content_blocks, list):
        content.extend(block for block in content_blocks if isinstance(block, dict))

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    http_referer = str(config.get("http_referer") or "").strip()
    x_title = str(config.get("x_title") or "").strip()
    if http_referer:
        headers["HTTP-Referer"] = http_referer
    if x_title:
        headers["X-Title"] = x_title

    payload = {
        "model": str(config.get("model") or "anthropic/claude-sonnet-4"),
        "messages": [
            {
                "role": "user",
                "content": content,
            }
        ],
        "temperature": 0,
        "max_completion_tokens": int(max_completion_tokens or config.get("max_tokens", 1024)),
    }
    response = requests.post(
        f"{str(config.get('base_url')).rstrip('/')}/chat/completions",
        headers=headers,
        json=payload,
        timeout=int(config.get("timeout_seconds", 45)),
    )
    if response.status_code != 200:
        raise RuntimeError(_openrouter_error_message(response))

    response_payload = _safe_json(response)
    if not isinstance(response_payload, dict):
        return ""

    choices = response_payload.get("choices")
    if not isinstance(choices, list) or not choices:
        return ""
    message = choices[0].get("message") if isinstance(choices[0], dict) else {}
    if not isinstance(message, dict):
        return ""
    content = message.get("content")
    if isinstance(content, str):
        return content.strip()
    extracted = extract_chat_message_text_content(content)
    if extracted:
        return extracted
    return str(content or "").strip()


def request_openrouter_vision_text(
    image_bytes: bytes,
    image_mime: str,
    prompt: str,
    config: dict,
    *,
    max_completion_tokens: int | None = None,
) -> str:
    if not image_bytes:
        return ""

    encoded_image = base64.b64encode(image_bytes).decode("ascii")
    return request_openrouter_chat_completion(
        prompt,
        config,
        content_blocks=[
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:{image_mime};base64,{encoded_image}",
                },
            }
        ],
        max_completion_tokens=max_completion_tokens,
    )


def build_copilot_inference_url(config: dict) -> str:
    base_url = str(config.get("base_url") or "https://models.github.ai").rstrip("/")
    org = str(config.get("org") or "").strip()
    if org:
        return f"{base_url}/orgs/{org}/inference/chat/completions"
    return f"{base_url}/inference/chat/completions"


def request_copilot_chat_completion(
    prompt: str,
    config: dict,
    *,
    content_blocks: list[dict] | None = None,
    max_tokens: int | None = None,
) -> str:
    api_key = str(config.get("api_key") or "").strip()
    if not api_key:
        raise RuntimeError("No Copilot API key is configured.")

    content = [{"type": "text", "text": prompt}]
    if isinstance(content_blocks, list):
        content.extend(block for block in content_blocks if isinstance(block, dict))

    payload = {
        "model": str(config.get("model") or "openai/gpt-4.1"),
        "messages": [
            {
                "role": "user",
                "content": content,
            }
        ],
        "temperature": 0,
        "max_tokens": int(max_tokens or config.get("max_tokens", 1024)),
    }
    response = requests.post(
        build_copilot_inference_url(config),
        headers={
            "Accept": "application/vnd.github+json",
            "Authorization": f"Bearer {api_key}",
            "X-GitHub-Api-Version": str(config.get("api_version") or "2026-03-10"),
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=int(config.get("timeout_seconds", 45)),
    )
    if response.status_code != 200:
        raise RuntimeError(_copilot_error_message(response))

    response_payload = _safe_json(response)
    if not isinstance(response_payload, dict):
        return ""

    choices = response_payload.get("choices")
    if not isinstance(choices, list) or not choices:
        return ""
    message = choices[0].get("message") if isinstance(choices[0], dict) else {}
    if not isinstance(message, dict):
        return ""
    content = message.get("content")
    if isinstance(content, str):
        return content.strip()
    extracted = extract_chat_message_text_content(content)
    if extracted:
        return extracted
    return str(content or "").strip()


def request_copilot_vision_text(
    image_bytes: bytes,
    image_mime: str,
    prompt: str,
    config: dict,
    *,
    max_tokens: int | None = None,
) -> str:
    if not image_bytes:
        return ""

    encoded_image = base64.b64encode(image_bytes).decode("ascii")
    return request_copilot_chat_completion(
        prompt,
        config,
        content_blocks=[
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:{image_mime};base64,{encoded_image}",
                },
            }
        ],
        max_tokens=max_tokens,
    )


def request_gemini_generate_content(
    prompt: str,
    config: dict,
    *,
    content_parts: list[dict] | None = None,
    max_output_tokens: int | None = None,
) -> str:
    api_key = str(config.get("api_key") or "").strip()
    if not api_key:
        raise RuntimeError("No Gemini API key is configured.")

    parts = list(content_parts or [])
    parts.append({"text": prompt})
    payload = {
        "contents": [
            {
                "role": "user",
                "parts": parts,
            }
        ],
        "generationConfig": {
            "temperature": 0,
            "maxOutputTokens": int(max_output_tokens or config.get("max_tokens", 1024)),
        },
    }
    response = requests.post(
        f"{str(config.get('base_url')).rstrip('/')}/models/{str(config.get('model') or 'gemini-3.5-flash')}:generateContent",
        headers={
            "x-goog-api-key": api_key,
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=int(config.get("timeout_seconds", 45)),
    )
    if response.status_code != 200:
        raise RuntimeError(_gemini_error_message(response))

    response_payload = _safe_json(response)
    if not isinstance(response_payload, dict):
        return ""

    candidates = response_payload.get("candidates")
    if not isinstance(candidates, list) or not candidates:
        return ""
    candidate = candidates[0] if isinstance(candidates[0], dict) else {}
    if not isinstance(candidate, dict):
        return ""
    content = candidate.get("content")
    if not isinstance(content, dict):
        return ""
    parts_payload = content.get("parts")
    if not isinstance(parts_payload, list):
        return ""

    output_parts: list[str] = []
    for part in parts_payload:
        if not isinstance(part, dict):
            continue
        text = part.get("text")
        if isinstance(text, str) and text.strip():
            output_parts.append(text.strip())
    return "\n".join(output_parts).strip()


def request_gemini_vision_text(
    image_bytes: bytes,
    image_mime: str,
    prompt: str,
    config: dict,
    *,
    max_output_tokens: int | None = None,
) -> str:
    if not image_bytes:
        return ""

    encoded_image = base64.b64encode(image_bytes).decode("ascii")
    return request_gemini_generate_content(
        prompt,
        config,
        content_parts=[
            {
                "inline_data": {
                    "mime_type": image_mime,
                    "data": encoded_image,
                }
            }
        ],
        max_output_tokens=max_output_tokens,
    )


def request_claude_rescue_alt_text(preview_entry: dict | None, row: dict, config: dict) -> str:
    if not isinstance(preview_entry, dict):
        return ""

    image_bytes = preview_entry.get("bytes")
    if not isinstance(image_bytes, (bytes, bytearray)) or not image_bytes:
        return ""

    normalized_image_bytes = normalize_image_to_png(bytes(image_bytes)) or bytes(image_bytes)
    image_ext = "png" if normalized_image_bytes != bytes(image_bytes) else str(preview_entry.get("ext", "png") or "png").lower()
    image_mime = "image/jpeg" if image_ext in {"jpg", "jpeg"} else f"image/{image_ext}"
    if "/" not in image_mime:
        image_mime = "image/png"

    rescue_prompt = build_groq_rescue_vision_prompt(str(row.get("role", "")).lower())
    rescue_raw = request_claude_vision_text(
        normalized_image_bytes,
        image_mime,
        rescue_prompt,
        config,
        max_tokens=max(int(config.get("max_tokens", 1024) or 1024), 512),
    )
    return sanitize_groq_generated_alt(rescue_raw, str(row.get("role", "")).lower(), row.get("page"))


def request_openrouter_rescue_alt_text(preview_entry: dict | None, row: dict, config: dict) -> str:
    if not isinstance(preview_entry, dict):
        return ""

    image_bytes = preview_entry.get("bytes")
    if not isinstance(image_bytes, (bytes, bytearray)) or not image_bytes:
        return ""

    normalized_image_bytes = normalize_image_to_png(bytes(image_bytes)) or bytes(image_bytes)
    image_ext = "png" if normalized_image_bytes != bytes(image_bytes) else str(preview_entry.get("ext", "png") or "png").lower()
    image_mime = "image/jpeg" if image_ext in {"jpg", "jpeg"} else f"image/{image_ext}"
    if "/" not in image_mime:
        image_mime = "image/png"

    rescue_prompt = build_groq_rescue_vision_prompt(str(row.get("role", "")).lower())
    rescue_raw = request_openrouter_vision_text(
        normalized_image_bytes,
        image_mime,
        rescue_prompt,
        config,
        max_completion_tokens=max(int(config.get("max_tokens", 1024) or 1024), 512),
    )
    return sanitize_groq_generated_alt(rescue_raw, str(row.get("role", "")).lower(), row.get("page"))


def request_copilot_rescue_alt_text(preview_entry: dict | None, row: dict, config: dict) -> str:
    if not isinstance(preview_entry, dict):
        return ""

    image_bytes = preview_entry.get("bytes")
    if not isinstance(image_bytes, (bytes, bytearray)) or not image_bytes:
        return ""

    normalized_image_bytes = normalize_image_to_png(bytes(image_bytes)) or bytes(image_bytes)
    image_ext = "png" if normalized_image_bytes != bytes(image_bytes) else str(preview_entry.get("ext", "png") or "png").lower()
    image_mime = "image/jpeg" if image_ext in {"jpg", "jpeg"} else f"image/{image_ext}"
    if "/" not in image_mime:
        image_mime = "image/png"

    rescue_prompt = build_groq_rescue_vision_prompt(str(row.get("role", "")).lower())
    rescue_raw = request_copilot_vision_text(
        normalized_image_bytes,
        image_mime,
        rescue_prompt,
        config,
        max_tokens=max(int(config.get("max_tokens", 1024) or 1024), 512),
    )
    return sanitize_groq_generated_alt(rescue_raw, str(row.get("role", "")).lower(), row.get("page"))


def request_gemini_rescue_alt_text(preview_entry: dict | None, row: dict, config: dict) -> str:
    if not isinstance(preview_entry, dict):
        return ""

    image_bytes = preview_entry.get("bytes")
    if not isinstance(image_bytes, (bytes, bytearray)) or not image_bytes:
        return ""

    normalized_image_bytes = normalize_image_to_png(bytes(image_bytes)) or bytes(image_bytes)
    image_ext = "png" if normalized_image_bytes != bytes(image_bytes) else str(preview_entry.get("ext", "png") or "png").lower()
    image_mime = "image/jpeg" if image_ext in {"jpg", "jpeg"} else f"image/{image_ext}"
    if "/" not in image_mime:
        image_mime = "image/png"

    rescue_prompt = build_groq_rescue_vision_prompt(str(row.get("role", "")).lower())
    rescue_raw = request_gemini_vision_text(
        normalized_image_bytes,
        image_mime,
        rescue_prompt,
        config,
        max_output_tokens=max(int(config.get("max_tokens", 1024) or 1024), 512),
    )
    return sanitize_groq_generated_alt(rescue_raw, str(row.get("role", "")).lower(), row.get("page"))


def request_groq_chat_completion(
    prompt: str,
    config: dict,
    *,
    content_blocks: list[dict] | None = None,
    max_completion_tokens: int | None = None,
    response_format: dict | None = None,
    key_cursor: int = 0,
) -> tuple[str, int]:
    api_keys = list(config.get("api_keys") or [])
    if not api_keys:
        fallback_key = str(config.get("api_key") or "").strip()
        if fallback_key:
            api_keys = [fallback_key]
    if not api_keys:
        raise RuntimeError("No Groq API keys are configured.")

    content = [{"type": "text", "text": prompt}]
    if isinstance(content_blocks, list):
        content.extend(block for block in content_blocks if isinstance(block, dict))

    last_error = "Groq request failed."
    total_keys = len(api_keys)
    retries_remaining = max(0, int(config.get("rate_limit_retries", 6) or 6))
    next_key_cursor = key_cursor
    while True:
        for offset in range(total_keys):
            key_index = (next_key_cursor + offset) % total_keys
            payload = {
                "model": str(config.get("model") or "meta-llama/llama-4-scout-17b-16e-instruct"),
                "messages": [
                    {
                        "role": "user",
                        "content": content,
                    }
                ],
                "temperature": 0.2,
                "max_completion_tokens": int(max_completion_tokens or config.get("max_completion_tokens", 120)),
                "top_p": 0.9,
                "stream": False,
            }
            if isinstance(response_format, dict):
                payload["response_format"] = response_format
            response = requests.post(
                f"{str(config.get('base_url')).rstrip('/')}/chat/completions",
                headers={
                    "Authorization": f"Bearer {api_keys[key_index]}",
                    "Content-Type": "application/json",
                },
                json=payload,
                timeout=int(config.get("timeout_seconds", 60)),
            )
            if response.status_code == 200:
                response_payload = _safe_json(response)
                if not isinstance(response_payload, dict):
                    return ("", (key_index + 1) % total_keys)

                choices = response_payload.get("choices")
                if not isinstance(choices, list) or not choices:
                    return ("", (key_index + 1) % total_keys)
                message = choices[0].get("message") if isinstance(choices[0], dict) else {}
                if not isinstance(message, dict):
                    return ("", (key_index + 1) % total_keys)
                return (str(message.get("content") or ""), (key_index + 1) % total_keys)

            last_error = _groq_error_message(response)
            if not is_groq_retryable_status(response.status_code, last_error):
                raise RuntimeError(last_error)

            wait_seconds = groq_retry_after_seconds(response, last_error, config)
            next_key_cursor = (key_index + 1) % total_keys
            if retries_remaining <= 0:
                raise RuntimeError(last_error)
            retries_remaining -= 1
            time.sleep(wait_seconds)
            break
        else:
            break

    raise RuntimeError(last_error)


def request_groq_text_completion(
    prompt: str,
    config: dict,
    *,
    max_completion_tokens: int | None = None,
    response_format: dict | None = None,
    key_cursor: int = 0,
) -> tuple[str, int]:
    return request_groq_chat_completion(
        prompt,
        config,
        max_completion_tokens=max_completion_tokens,
        response_format=response_format,
        key_cursor=key_cursor,
    )


def request_groq_vision_alt_text(
    image_bytes: bytes,
    image_mime: str,
    prompt: str,
    config: dict,
    *,
    max_completion_tokens: int | None = None,
    response_format: dict | None = None,
    key_cursor: int = 0,
) -> tuple[str, int]:
    if not image_bytes:
        return ("", key_cursor)

    encoded_image = base64.b64encode(image_bytes).decode("ascii")
    return request_groq_chat_completion(
        prompt,
        config,
        content_blocks=[
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:{image_mime};base64,{encoded_image}",
                },
            }
        ],
        max_completion_tokens=max_completion_tokens,
        response_format=response_format,
        key_cursor=key_cursor,
    )


def build_groq_single_ocr_prompt(role: str, hint_text: str = "") -> str:
    cleaned_hint = normalize_alt_text(hint_text)
    if role == "equation":
        prompt = """
        Read the equation in the image and transcribe only the exact visible math.
        Keep the math symbols exactly when they are visible.
        If a fill-in blank line is visible, transcribe that blank as exactly three underscores: ___
        Preserve the operator order exactly and never add or remove operators.
        Do not explain, interpret, or write alt text.
        Return only the transcription text.
        """
    else:
        prompt = """
        Read the image and transcribe the core visible content, labels, or chart text.
        Do not write accessibility alt text yet.
        Return only the transcription text.
        """
    if cleaned_hint:
        prompt = f"{textwrap.dedent(prompt).strip()}\nUseful OCR hint: {cleaned_hint[:240]}"
    else:
        prompt = textwrap.dedent(prompt).strip()
    return prompt


def request_groq_ocr_then_rewrite_alt_text(
    image_bytes: bytes,
    image_mime: str,
    row: dict,
    config: dict,
    pdf_path: Path | None,
    docx_path: Path | None = None,
) -> str:
    role = str(row.get("role", "")).lower()
    hint_text = groq_row_hint_text(row, pdf_path, docx_path=docx_path)
    ocr_prompt = build_groq_single_ocr_prompt(role, hint_text)
    ocr_config = build_groq_stage_config(config, "ocr")
    ocr_cursor = get_groq_stage_cursor(config, "ocr")
    raw_ocr, next_ocr_cursor = request_groq_vision_alt_text(
        image_bytes,
        image_mime,
        ocr_prompt,
        ocr_config,
        max_completion_tokens=int(config.get("ocr_max_completion_tokens", config.get("max_completion_tokens", 120)) or 120),
        key_cursor=ocr_cursor,
    )
    set_groq_stage_cursor(config, "ocr", next_ocr_cursor)
    ocr_text = sanitize_groq_ocr_text(raw_ocr, role)
    if not ocr_text and hint_text:
        ocr_text = sanitize_groq_ocr_text(hint_text, role)
    if not ocr_text:
        return ""

    if role == "image":
        rescued_alt = request_groq_rescue_alt_text(
            image_bytes,
            image_mime,
            row,
            config,
            ocr_text=ocr_text,
        )
        if rescued_alt:
            return rescued_alt
        return ""

    rewrite_prompt = build_groq_alt_prompt(role, ocr_text)
    rewrite_config = build_groq_stage_config(config, "rewrite")
    rewrite_cursor = get_groq_stage_cursor(config, "rewrite")
    raw_alt, next_rewrite_cursor = request_groq_text_completion(
        rewrite_prompt,
        rewrite_config,
        max_completion_tokens=int(config.get("rewrite_max_completion_tokens", config.get("max_completion_tokens", 120)) or 120),
        key_cursor=rewrite_cursor,
    )
    set_groq_stage_cursor(config, "rewrite", next_rewrite_cursor)
    cleaned_alt = sanitize_groq_generated_alt(raw_alt, role, row.get("page"))
    if is_meaningful_groq_alt_rewrite(cleaned_alt, role, ocr_text):
        return cleaned_alt
    rescued_alt = request_groq_rescue_alt_text(
        image_bytes,
        image_mime,
        row,
        config,
        ocr_text=ocr_text,
    )
    if rescued_alt:
        return rescued_alt
    return ""


def sanitize_groq_ocr_text(text: str, role: str) -> str:
    cleaned = strip_llm_wrappers(normalize_alt_text(text))
    cleaned = re.sub(r"^(?:ocr|ocr text|transcription|transcribed text)\s*[:=\-]\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.strip().strip('"').strip("'")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    if not cleaned:
        return ""
    if role == "equation":
        return strict_equation_transcription(cleaned)
    return clean_image_ocr_candidate(cleaned)


def sanitize_groq_generated_alt(text: str, role: str, page: int | None) -> str:
    cleaned = strip_llm_wrappers(normalize_alt_text(text))
    cleaned = cleaned.strip().strip('"').strip("'")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    if not cleaned:
        return ""
    cleaned = strip_forbidden_alt_prefixes(cleaned)
    if looks_like_filename_metadata_text(cleaned):
        return ""

    if role == "image":
        cleaned = sanitize_generated_image_alt(cleaned)
        cleaned = numbers_to_spoken_words(cleaned)
    else:
        cleaned = re.split(r"(?<=[.!?])\s+", cleaned, maxsplit=1)[0].strip()
        cleaned = equation_symbols_to_spoken_words(cleaned)
        cleaned = numbers_to_spoken_words(cleaned)

    if not cleaned:
        return ""
    return finalize_alt_text(cleaned, role, page)


def build_groq_rescue_vision_prompt(role: str, ocr_text: str = "") -> str:
    cleaned_ocr = normalize_alt_text(ocr_text)
    if role == "equation":
        prompt = """
        Write the final accessibility alt text for this single mathematical image.
        Focus only on the main equation or mathematical expression.
        Ignore worksheet labels, workbook text, row or column references, cell references, borders, and interface text.
        The alt text must be easy to understand, detailed, and screen-reader worthy.
        All numericals must be written as words.
        If a decimal number is visible, speak the decimal point as "point". Example: 6.5 becomes "six point five".
        Never start with "Equation shows", "Equation showing", "Image shows", or "Image showing".
        Do not read punctuation names like comma, brace, or bracket aloud unless they are the key meaning.
        If this is a set or sequence, describe it as a set or sequence in natural language.
        Return only the exact final alt text.
        """
    else:
        prompt = """
        Write the final accessibility alt text for this single educational image.
        Focus only on the actual image content.
        Ignore worksheet labels, workbook text, row or column references, cell references, borders, and interface text.
        The alt text must be easy to understand, detailed, and screen-reader worthy.
        All numericals must be written as words.
        If a decimal number is visible, speak the decimal point as "point". Example: 57.5 becomes "fifty seven point five".
        Never start with "Equation shows", "Equation showing", "Image shows", or "Image showing".
        Do not say "Here is the alt text", "Screen-reader-friendly alt text", "ALT text:", or any similar introduction.
        Write only the exact final alt text.
        Describe the main subject, labels, structure, and the key visual takeaway.
        If it is a graph, chart, or diagram, use four to twelve detailed sentences: name the chart type, title, axes, scales, units, labels, legend, line or shape, and the main relationship or trend, including standout highs, lows, intersections, or changes when visible.
        If it is a histogram, bar chart, line graph, scatter plot, or frequency chart, list each visible interval, category, bar, bin, or series value when it can be read reliably.
        If it is a table or spreadsheet, use four to twelve detailed sentences: name the table structure, mention the visible column headers and row labels, and summarize the most important values, comparisons, totals, or patterns.
        If it is a photo or illustration, describe the subject, count, appearance, and setting when visible.
        Do not answer with instructions, requests, placeholders, or meta commentary.
        Return only the exact final alt text.
        """
    prompt = textwrap.dedent(prompt).strip()
    if cleaned_ocr:
        prompt += f"\nOCR transcription hint: {cleaned_ocr[:260]}"
    return prompt


def build_groq_rescue_text_prompt(role: str, ocr_text: str) -> str:
    cleaned_ocr = normalize_alt_text(ocr_text)
    if role == "equation":
        prompt = f"""
        Rewrite this OCR transcription into the final accessibility alt text for a mathematical image:
        {cleaned_ocr}
        Rules:
        - Make it easy to understand, detailed, and screen-reader worthy.
        - All numericals must be written as words.
        - If a decimal number is visible, speak the decimal point as "point". Example: 57.5 becomes "fifty seven point five".
        - Never start with "Equation shows", "Equation showing", "Image shows", or "Image showing".
        - Do not read punctuation names like comma, brace, or bracket aloud unless they are the key meaning.
        - If it is set notation or a sequence, describe it as a set or sequence in natural language.
        - Return only the exact final alt text.
        """
    else:
        prompt = f"""
        Rewrite this OCR transcription into the final accessibility alt text for an educational image:
        {cleaned_ocr}
        Rules:
        - Make it easy to understand, detailed, and screen-reader worthy.
        - All numericals must be written as words.
        - If a decimal number is visible, speak the decimal point as "point". Example: 57.5 becomes "fifty seven point five".
        - Never start with "Equation shows", "Equation showing", "Image shows", or "Image showing".
        - Do not include any intro such as "Here is the alt text" or "Screen-reader-friendly alt text".
        - Describe the actual image content, not worksheet or interface text.
        - If the image is a table or spreadsheet, explain the structure, visible headers, row labels, and the most important values, totals, or comparisons in more detail.
        - If the image is a graph, chart, or diagram, explain the chart type, title, axes, scale, legend, labels, and main trend or comparison in more detail.
        - If the image is a histogram, bar chart, line graph, scatter plot, or frequency chart, list each visible interval, category, bar, bin, or series value when it can be read reliably.
        - Return only the exact final alt text.
        """
    return textwrap.dedent(prompt).strip()


def request_groq_rescue_alt_text(
    image_bytes: bytes,
    image_mime: str,
    row: dict,
    config: dict,
    *,
    ocr_text: str = "",
) -> str:
    role = str(row.get("role", "")).lower()
    rewrite_config = build_groq_stage_config(config, "rewrite")
    rewrite_cursor = get_groq_stage_cursor(config, "rewrite")

    rescue_prompt = build_groq_rescue_vision_prompt(role, ocr_text)
    rescue_raw, next_rewrite_cursor = request_groq_vision_alt_text(
        image_bytes,
        image_mime,
        rescue_prompt,
        rewrite_config,
        max_completion_tokens=max(
            int(config.get("rewrite_max_completion_tokens", config.get("max_completion_tokens", 120)) or 120),
            180,
        ),
        key_cursor=rewrite_cursor,
    )
    set_groq_stage_cursor(config, "rewrite", next_rewrite_cursor)
    cleaned_alt = sanitize_groq_generated_alt(rescue_raw, role, row.get("page"))
    if is_meaningful_groq_alt_rewrite(cleaned_alt, role, ocr_text):
        return cleaned_alt

    cleaned_ocr = sanitize_groq_ocr_text(ocr_text, role)
    if cleaned_ocr:
        rewrite_cursor = get_groq_stage_cursor(config, "rewrite")
        rescue_text_prompt = build_groq_rescue_text_prompt(role, cleaned_ocr)
        rescue_text_raw, next_rewrite_cursor = request_groq_text_completion(
            rescue_text_prompt,
            rewrite_config,
            max_completion_tokens=max(
                int(config.get("rewrite_max_completion_tokens", config.get("max_completion_tokens", 120)) or 120),
                180,
            ),
            key_cursor=rewrite_cursor,
        )
        set_groq_stage_cursor(config, "rewrite", next_rewrite_cursor)
        cleaned_alt = sanitize_groq_generated_alt(rescue_text_raw, role, row.get("page"))
        if is_meaningful_groq_alt_rewrite(cleaned_alt, role, cleaned_ocr):
            return cleaned_alt

    return ""


def is_meaningful_groq_alt_rewrite(text: str, role: str, ocr_text: str) -> bool:
    cleaned = normalize_alt_text(text)
    if not cleaned:
        return False
    lowered = cleaned.lower()
    source = normalize_rewrite_source_text(ocr_text)
    if looks_like_groq_ui_leak(cleaned):
        return False
    if role == "equation":
        if lowered.count(" comma ") >= 2:
            return False
        if any(marker in lowered for marker in ("open brace", "close brace", "open bracket", "close bracket")):
            return False
        if "ellipsis" in lowered and "and so on" not in lowered:
            return False
        if ("..." in source or "…" in source) and "," in source:
            if "comma" in lowered:
                return False
            if not any(marker in lowered for marker in ("set ", "set containing", "sequence", "and so on")):
                return False
        if len(cleaned.split()) < 4:
            return False
        return True
    if len(cleaned.split()) < 5:
        return False
    if any(
        marker in lowered
        for marker in (
            "please provide",
            "i will create",
            "i can create",
            "alt text",
            "description, and i will",
        )
    ):
        return False
    if any(marker in lowered for marker in ("graph", "chart", "diagram")):
        if not any(
            marker in lowered
            for marker in (
                "x axis",
                "y axis",
                "labeled",
                "label",
                "line",
                "curve",
                "intersect",
                "origin",
                "rises",
                "slopes",
                "v-shaped",
                "v shaped",
            )
        ):
            return False
    return True


def looks_like_groq_ui_leak(text: str) -> bool:
    lowered = normalize_alt_text(text).lower()
    if not lowered:
        return False
    if re.search(r"\br\d+c\d+\b", lowered):
        return True
    strong_markers = (
        "alt management",
        "worksheet",
        "workbook",
        "cell reference",
        "row reference",
        "column reference",
        "interface text",
    )
    if any(marker in lowered for marker in strong_markers):
        return True
    combo_markers = (
        ("structured", "grid"),
        ("two by two", "grid"),
        ("tile", "grid"),
        ("row", "column"),
        ("preview", "grid"),
    )
    return any(all(marker in lowered for marker in combo) for combo in combo_markers)


def looks_like_filename_metadata_text(text: str) -> bool:
    lowered = compact_ocr_text(text).strip().strip('"').strip("'").lower()
    if not lowered:
        return False
    if lowered in {
        "file name",
        "image file",
        "source label",
        "status",
        "alt text",
        "preview unavailable",
    }:
        return True
    if any(
        marker in lowered
        for marker in (
            "alt_text_manifest.xlsx",
            "grid_batch.docx",
            "_alt_inventory",
            "_pdf_alt_inventory",
        )
    ):
        return True
    return bool(
        re.fullmatch(
            r"[a-z0-9][a-z0-9 _().-]{0,220}\.(?:docx|doc|pdf|xlsx|xls|png|jpg|jpeg|gif|bmp|webp|wmf|emf)",
            lowered,
        )
    )


def normalize_spoken_list_text(text: str) -> str:
    working = compact_ocr_text(text)
    if not working:
        return ""
    working = re.sub(r"\s*,\s*", ", ", working)
    working = re.sub(r"\s+", " ", working).strip(" ,")
    return working


def build_accessible_alt_from_ocr(ocr_text: str, role: str, page: int | None) -> str:
    cleaned_ocr = normalize_rewrite_source_text(ocr_text)
    if not cleaned_ocr:
        return ""
    if role != "equation":
        return sanitize_groq_generated_alt(cleaned_ocr, role, page)

    semantic_equation_alt = build_semantic_equation_alt_from_ocr(cleaned_ocr)
    if semantic_equation_alt:
        return finalize_alt_text(semantic_equation_alt, role, page)
    return sanitize_groq_generated_alt(cleaned_ocr, role, page)


def build_semantic_equation_alt_from_ocr(ocr_text: str) -> str:
    source = normalize_rewrite_source_text(ocr_text)
    if not source:
        return ""

    has_ellipsis = "..." in source or "…" in source
    has_braces = any(marker in source for marker in ("{", "}", "[", "]"))
    sequence_body = source.replace("...", "").replace("…", "")
    sequence_body = sequence_body.replace("{", " ").replace("}", " ")
    sequence_body = sequence_body.replace("[", " ").replace("]", " ")
    sequence_body = sequence_body.replace("(", " ").replace(")", " ")
    sequence_body = normalize_spoken_list_text(sequence_body)

    if has_ellipsis and "," in source and sequence_body:
        spoken_body = equation_symbols_to_spoken_words(sequence_body)
        spoken_body = numbers_to_spoken_words(spoken_body)
        spoken_body = normalize_spoken_list_text(spoken_body)
        if spoken_body:
            if has_braces:
                return f"Set containing {spoken_body}, and so on"
            return f"Sequence {spoken_body}, and so on"

    if has_braces and "," in source and sequence_body:
        spoken_body = equation_symbols_to_spoken_words(sequence_body)
        spoken_body = numbers_to_spoken_words(spoken_body)
        spoken_body = normalize_spoken_list_text(spoken_body)
        if spoken_body:
            return f"Set containing {spoken_body}"

    if any(marker in source for marker in (">=", "<=", "!=", "<>", "=", "+", "-", "/", "^", "≤", "≥")):
        spoken = equation_symbols_to_spoken_words(source)
        spoken = numbers_to_spoken_words(spoken)
        return normalize_spoken_list_text(spoken)

    return ""


def strip_forbidden_alt_prefixes(text: str) -> str:
    cleaned = compact_ocr_text(text)
    if not cleaned:
        return ""
    patterns = (
        r"^(?:this|the)\s+(?:equation|image)\s+(?:show(?:s|ing)?|contains?)\s+",
        r"^(?:equation|equations|image|images|picture|photo)\s+(?:show(?:s|ing)?|of)\s+",
        r"^(?:equation|image)\s*[:\-]\s*",
        r"^(?:there is|there are)\s+",
        r'^(?:here is|here\'s)\s+(?:a\s+|the\s+)?(?:screen[\s-]?reader(?:[\s-]?friendly)?\s+)?(?:alt\s+text|description)(?:\s+(?:for|of)\s+(?:the\s+)?)?(?:image|figure|chart|graph|table|diagram)?\s*[:\-]\s*',
        r"^(?:screen[\s-]?reader(?:[\s-]?friendly)?\s+)?(?:alt\s+text|description)\s*(?:for|of)?\s*(?:the\s+)?(?:image|figure|chart|graph|table|diagram)?\s*[:\-]\s*",
    )
    for pattern in patterns:
        cleaned = re.sub(pattern, "", cleaned, flags=re.IGNORECASE)
    return cleaned.strip().strip('"').strip("'")


def is_data_dense_image_alt(text: str) -> bool:
    lowered = normalize_alt_text(text).lower()
    if not lowered:
        return False
    markers = (
        "table",
        "spreadsheet",
        "graph",
        "chart",
        "plot",
        "diagram",
        "axis",
        "axes",
        "legend",
        "column",
        "columns",
        "row",
        "rows",
        "header",
        "headers",
        "frequency",
        "tally",
    )
    return any(marker in lowered for marker in markers)


def numbers_to_spoken_words(text: str) -> str:
    working = compact_ocr_text(text)
    if not working:
        return ""

    working = working.replace("%", " percent ")

    def replace_number(match: re.Match) -> str:
        value = str(match.group(0) or "").replace(",", "")
        if not value:
            return ""
        negative = value.startswith("-")
        core = value[1:] if negative else value
        if "." in core:
            whole, fraction = core.split(".", 1)
            whole_words = integer_to_words(whole or "0")
            fraction_words = " ".join(integer_to_words(digit) for digit in fraction if digit.isdigit())
            spoken = f"{whole_words} point {fraction_words}".strip()
        else:
            spoken = integer_to_words(core)
        if negative:
            spoken = f"minus {spoken}"
        return spoken

    working = re.sub(r"(?<![A-Za-z0-9])-?\d[\d,]*(?:\.\d+)?(?![A-Za-z0-9])", replace_number, working)
    working = re.sub(r"\s+", " ", working).strip()
    return working


def is_groq_retryable_status(status_code: int, error_message: str) -> bool:
    if status_code in {408, 409, 425, 429, 500, 502, 503, 504}:
        return True
    lowered = normalize_alt_text(error_message).lower()
    return any(marker in lowered for marker in ("rate limit", "too many requests", "temporarily unavailable", "timeout"))


def groq_retry_after_seconds(response: requests.Response, error_message: str, config: dict) -> float:
    retry_after = normalize_alt_text(response.headers.get("retry-after", ""))
    wait_seconds = parse_duration_seconds(retry_after)
    if wait_seconds <= 0:
        wait_seconds = parse_duration_seconds(normalize_alt_text(response.headers.get("x-ratelimit-reset-tokens", "")))
    if wait_seconds <= 0:
        match = re.search(r"try again in\s+([0-9]+(?:\.[0-9]+)?)s", error_message, flags=re.IGNORECASE)
        if match:
            try:
                wait_seconds = float(match.group(1))
            except ValueError:
                wait_seconds = 0.0
    if wait_seconds <= 0:
        wait_seconds = 2.5

    cap_seconds = float(config.get("retry_after_cap_seconds", 20) or 20)
    return max(0.5, min(wait_seconds + 0.35, max(0.5, cap_seconds)))


def parse_duration_seconds(value: str) -> float:
    cleaned = normalize_alt_text(value)
    if not cleaned:
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        pass

    total = 0.0
    matched = False
    for amount, unit in re.findall(r"([0-9]+(?:\.[0-9]+)?)\s*([smh])", cleaned, flags=re.IGNORECASE):
        try:
            numeric = float(amount)
        except ValueError:
            continue
        matched = True
        unit = unit.lower()
        if unit == "h":
            total += numeric * 3600.0
        elif unit == "m":
            total += numeric * 60.0
        else:
            total += numeric
    return total if matched else 0.0


def _groq_error_message(response: requests.Response) -> str:
    payload = _safe_json(response)
    if isinstance(payload, dict):
        error = payload.get("error")
        if isinstance(error, dict):
            message = normalize_alt_text(error.get("message", ""))
            if message:
                return message
        message = normalize_alt_text(payload.get("message", ""))
        if message:
            return message
    text = normalize_alt_text(response.text)
    if text:
        return text[:240]
    return f"Groq request failed with status {response.status_code}."


def _claude_error_message(response: requests.Response) -> str:
    payload = _safe_json(response)
    if isinstance(payload, dict):
        error = payload.get("error")
        if isinstance(error, dict):
            message = normalize_alt_text(error.get("message", "")) or normalize_alt_text(error.get("type", ""))
            if message:
                return message
        message = normalize_alt_text(payload.get("message", ""))
        if message:
            return message
    text = normalize_alt_text(response.text)
    if text:
        return text[:240]
    return f"Claude request failed with status {response.status_code}."


def is_claude_auth_error(response: requests.Response, message: str = "") -> bool:
    lowered = normalize_alt_text(message).lower()
    return response.status_code in {401, 403} or any(
        marker in lowered
        for marker in (
            "invalid x-api-key",
            "api key",
            "x-api-key",
            "authentication",
            "unauthorized",
            "permission",
        )
    )


def _openrouter_error_message(response: requests.Response) -> str:
    payload = _safe_json(response)
    if isinstance(payload, dict):
        error = payload.get("error")
        if isinstance(error, dict):
            message = normalize_alt_text(error.get("message", "")) or normalize_alt_text(error.get("code", ""))
            if message:
                return message
        message = normalize_alt_text(payload.get("message", ""))
        if message:
            return message
    text = normalize_alt_text(response.text)
    if text:
        return text[:240]
    return f"OpenRouter request failed with status {response.status_code}."


def _copilot_error_message(response: requests.Response) -> str:
    payload = _safe_json(response)
    if isinstance(payload, dict):
        error = payload.get("error")
        if isinstance(error, dict):
            message = (
                normalize_alt_text(error.get("message", ""))
                or normalize_alt_text(error.get("code", ""))
                or normalize_alt_text(error.get("status", ""))
            )
            if message:
                return message
        message = normalize_alt_text(payload.get("message", ""))
        if message:
            return message
    text = normalize_alt_text(response.text)
    if text:
        return text[:240]
    return f"Copilot request failed with status {response.status_code}."


def _gemini_error_message(response: requests.Response) -> str:
    payload = _safe_json(response)
    if isinstance(payload, dict):
        error = payload.get("error")
        if isinstance(error, dict):
            message = normalize_alt_text(error.get("message", "")) or normalize_alt_text(error.get("status", ""))
            if message:
                return message
        message = normalize_alt_text(payload.get("message", ""))
        if message:
            return message
    text = normalize_alt_text(response.text)
    if text:
        return text[:240]
    return f"Gemini request failed with status {response.status_code}."


def is_equation_generation_enabled() -> bool:
    return (os.getenv("MATCHA_ALT_ENABLE_EQUATION_GENERATION") or "").strip().lower() in {"1", "true", "yes", "on"}


PDF_FORMULA_PREVIEW_PADDING = {
    "min_pad_x": 0.15,
    "min_pad_y": 0.75,
    "relative_pad_x": 0.004,
    "relative_pad_y": 0.035,
}


PDF_FIGURE_PREVIEW_PADDING = {
    "min_pad_x": 0.75,
    "min_pad_y": 0.75,
    "relative_pad_x": 0.0,
    "relative_pad_y": 0.0,
}

PDF_PREVIEW_RENDER_SCALE = 4.0
PDF_VISUAL_REFINEMENT_EXPAND_PT = 6.0
PDF_VISUAL_REFINEMENT_WHITE_THRESHOLD = 246
PDF_VISUAL_REFINEMENT_PADDING_PX = 4
PDF_VISUAL_REFINEMENT_ANALYSIS_MAX_EDGE = 960


def tighten_pdf_preview_png(image_bytes: bytes) -> bytes | None:
    if not image_bytes:
        return None
    try:
        with Image.open(io.BytesIO(image_bytes)) as image:
            trimmed = crop_image_whitespace(image, threshold=12)
            output = io.BytesIO()
            trimmed.save(output, format="PNG")
            return output.getvalue()
    except Exception:
        return None


def expand_preview_rect(rect: fitz.Rect, page_rect: fitz.Rect, *, x_pad: float, y_pad: float) -> fitz.Rect:
    return fitz.Rect(
        max(page_rect.x0, float(rect.x0) - max(0.0, float(x_pad))),
        max(page_rect.y0, float(rect.y0) - max(0.0, float(y_pad))),
        min(page_rect.x1, float(rect.x1) + max(0.0, float(x_pad))),
        min(page_rect.y1, float(rect.y1) + max(0.0, float(y_pad))),
    )


def pixel_box_overlap_area(
    left: tuple[int, int, int, int],
    right: tuple[int, int, int, int],
) -> int:
    overlap_x = max(0, min(int(left[2]), int(right[2])) - max(int(left[0]), int(right[0])))
    overlap_y = max(0, min(int(left[3]), int(right[3])) - max(int(left[1]), int(right[1])))
    return overlap_x * overlap_y


def pixel_box_gap(
    left: tuple[int, int, int, int],
    right: tuple[int, int, int, int],
) -> tuple[int, int]:
    gap_x = max(0, max(int(right[0]) - int(left[2]), int(left[0]) - int(right[2])))
    gap_y = max(0, max(int(right[1]) - int(left[3]), int(left[1]) - int(right[3])))
    return gap_x, gap_y


def pixel_box_union(boxes: list[tuple[int, int, int, int]]) -> tuple[int, int, int, int] | None:
    if not boxes:
        return None
    return (
        min(box[0] for box in boxes),
        min(box[1] for box in boxes),
        max(box[2] for box in boxes),
        max(box[3] for box in boxes),
    )


def pdf_visible_content_mask(
    image: Image.Image,
    *,
    white_threshold: int = PDF_VISUAL_REFINEMENT_WHITE_THRESHOLD,
) -> Image.Image:
    rgb = image.convert("RGB")
    diff = ImageChops.difference(rgb, Image.new("RGB", rgb.size, (255, 255, 255))).convert("L")
    delta_threshold = max(1, 255 - int(white_threshold))
    return diff.point(lambda value: 255 if value > delta_threshold else 0)


def analyze_pdf_mask_size(image_size: tuple[int, int]) -> tuple[int, int]:
    width, height = image_size
    max_edge = max(width, height)
    if max_edge <= PDF_VISUAL_REFINEMENT_ANALYSIS_MAX_EDGE:
        return width, height
    scale = float(PDF_VISUAL_REFINEMENT_ANALYSIS_MAX_EDGE) / float(max_edge)
    return (
        max(1, int(round(width * scale))),
        max(1, int(round(height * scale))),
    )


def connected_mask_components(mask: Image.Image) -> list[dict]:
    width, height = mask.size
    pixels = mask.load()
    visited = bytearray(width * height)
    components: list[dict] = []

    for y in range(height):
        row_offset = y * width
        for x in range(width):
            index = row_offset + x
            if visited[index] or pixels[x, y] == 0:
                continue
            visited[index] = 1
            stack = [(x, y)]
            min_x = max_x = x
            min_y = max_y = y
            area = 0
            while stack:
                px, py = stack.pop()
                area += 1
                if px < min_x:
                    min_x = px
                if py < min_y:
                    min_y = py
                if px > max_x:
                    max_x = px
                if py > max_y:
                    max_y = py
                for ny in range(max(0, py - 1), min(height, py + 2)):
                    next_row_offset = ny * width
                    for nx in range(max(0, px - 1), min(width, px + 2)):
                        next_index = next_row_offset + nx
                        if visited[next_index] or pixels[nx, ny] == 0:
                            continue
                        visited[next_index] = 1
                        stack.append((nx, ny))
            components.append({"bbox": (min_x, min_y, max_x + 1, max_y + 1), "area": area})
    return components


def scale_pixel_box(
    box: tuple[int, int, int, int],
    source_size: tuple[int, int],
    target_size: tuple[int, int],
) -> tuple[int, int, int, int]:
    source_width, source_height = source_size
    target_width, target_height = target_size
    if source_width <= 0 or source_height <= 0:
        return box
    scale_x = float(target_width) / float(source_width)
    scale_y = float(target_height) / float(source_height)
    return (
        max(0, int(round(box[0] * scale_x))),
        max(0, int(round(box[1] * scale_y))),
        min(target_width, int(round(box[2] * scale_x))),
        min(target_height, int(round(box[3] * scale_y))),
    )


def shrink_pixel_box(
    box: tuple[int, int, int, int],
    *,
    min_inset: int = 4,
    inset_ratio: float = 0.18,
) -> tuple[int, int, int, int]:
    width = max(1, int(box[2]) - int(box[0]))
    height = max(1, int(box[3]) - int(box[1]))
    inset_x = min(max(0, width // 3), max(0, int(round(width * inset_ratio)), min_inset))
    inset_y = min(max(0, height // 3), max(0, int(round(height * inset_ratio)), min_inset))
    left = int(box[0]) + inset_x
    top = int(box[1]) + inset_y
    right = int(box[2]) - inset_x
    bottom = int(box[3]) - inset_y
    if right <= left or bottom <= top:
        return box
    return (left, top, right, bottom)


def component_center_distance(component_box: tuple[int, int, int, int], anchor_box: tuple[int, int, int, int]) -> float:
    component_cx = (int(component_box[0]) + int(component_box[2])) / 2.0
    component_cy = (int(component_box[1]) + int(component_box[3])) / 2.0
    anchor_cx = (int(anchor_box[0]) + int(anchor_box[2])) / 2.0
    anchor_cy = (int(anchor_box[1]) + int(anchor_box[3])) / 2.0
    return abs(component_cx - anchor_cx) + abs(component_cy - anchor_cy)


def component_center_in_padded_box(
    component_box: tuple[int, int, int, int],
    anchor_box: tuple[int, int, int, int],
    *,
    pad_x: int,
    pad_y: int,
) -> bool:
    center_x = (int(component_box[0]) + int(component_box[2])) / 2.0
    center_y = (int(component_box[1]) + int(component_box[3])) / 2.0
    return (
        center_x >= int(anchor_box[0]) - int(pad_x)
        and center_x <= int(anchor_box[2]) + int(pad_x)
        and center_y >= int(anchor_box[1]) - int(pad_y)
        and center_y <= int(anchor_box[3]) + int(pad_y)
    )


def select_main_pdf_component(
    components: list[dict],
    anchor_box: tuple[int, int, int, int],
    *,
    role: str,
) -> dict | None:
    if not components:
        return None
    focus_box = shrink_pixel_box(anchor_box)
    best_component = None
    best_score: tuple[float, float, float] | None = None
    for component in components:
        component_box = component["bbox"]
        focus_overlap = pixel_box_overlap_area(component_box, focus_box)
        anchor_overlap = pixel_box_overlap_area(component_box, anchor_box)
        area = int(component.get("area", 0))
        distance = component_center_distance(component_box, focus_box)
        if role == "equation":
            score = (
                float(focus_overlap * 5 + anchor_overlap * 3 + area),
                -distance,
                -float(component_box[1]),
            )
        else:
            score = (
                float(focus_overlap * 4 + anchor_overlap * 4 + area * 2),
                -distance,
                -float(component_box[1]),
            )
        if best_score is None or score > best_score:
            best_score = score
            best_component = component
    return best_component


def should_attach_formula_component(
    union_box: tuple[int, int, int, int],
    component_box: tuple[int, int, int, int],
    anchor_box: tuple[int, int, int, int],
) -> bool:
    overlap_anchor = pixel_box_overlap_area(component_box, anchor_box)
    center_in_anchor_zone = component_center_in_padded_box(component_box, anchor_box, pad_x=12, pad_y=10)
    gap_x, gap_y = pixel_box_gap(union_box, component_box)
    union_height = max(1, int(union_box[3]) - int(union_box[1]))
    component_height = max(1, int(component_box[3]) - int(component_box[1]))
    union_center_y = (int(union_box[1]) + int(union_box[3])) / 2.0
    component_center_y = (int(component_box[1]) + int(component_box[3])) / 2.0
    baseline_delta = abs(component_center_y - union_center_y)

    if overlap_anchor > 0 and gap_y <= max(12, union_height):
        return True
    if center_in_anchor_zone and gap_x <= max(10, union_height) and gap_y <= max(10, union_height):
        return True
    if center_in_anchor_zone and baseline_delta <= max(12, union_height * 0.8) and gap_x <= max(12, component_height * 2):
        return True
    return False


def should_attach_figure_component(
    union_box: tuple[int, int, int, int],
    component_box: tuple[int, int, int, int],
    anchor_box: tuple[int, int, int, int],
) -> bool:
    if pixel_box_overlap_area(component_box, anchor_box) > 0:
        return True
    if not component_center_in_padded_box(component_box, anchor_box, pad_x=18, pad_y=16):
        return False
    gap_x, gap_y = pixel_box_gap(union_box, component_box)
    union_width = max(1, int(union_box[2]) - int(union_box[0]))
    union_height = max(1, int(union_box[3]) - int(union_box[1]))
    return gap_x <= max(16, int(round(union_width * 0.12))) and gap_y <= max(16, int(round(union_height * 0.12)))


def refine_pdf_visible_content_bbox(
    image: Image.Image,
    anchor_box: tuple[int, int, int, int],
    *,
    role: str,
    white_threshold: int = PDF_VISUAL_REFINEMENT_WHITE_THRESHOLD,
    padding: int = PDF_VISUAL_REFINEMENT_PADDING_PX,
) -> tuple[int, int, int, int] | None:
    if image.width <= 0 or image.height <= 0:
        return None

    analysis_size = analyze_pdf_mask_size(image.size)
    analysis_image = image if analysis_size == image.size else image.resize(analysis_size, Image.Resampling.BILINEAR)
    analysis_mask = pdf_visible_content_mask(analysis_image, white_threshold=white_threshold)
    components = [component for component in connected_mask_components(analysis_mask) if int(component.get("area", 0)) >= 2]
    if not components:
        bbox = analysis_mask.getbbox()
        if not bbox:
            return None
        components = [{"bbox": bbox, "area": (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])}]

    analysis_anchor = scale_pixel_box(anchor_box, image.size, analysis_size)
    main_component = select_main_pdf_component(components, analysis_anchor, role=role)
    if main_component is None:
        return None

    selected_boxes = [tuple(int(value) for value in main_component["bbox"])]
    current_union = pixel_box_union(selected_boxes) or selected_boxes[0]
    changed = True
    while changed:
        changed = False
        for component in components:
            component_box = tuple(int(value) for value in component["bbox"])
            if component_box in selected_boxes:
                continue
            if role == "equation":
                attach = should_attach_formula_component(current_union, component_box, analysis_anchor)
            else:
                attach = should_attach_figure_component(current_union, component_box, analysis_anchor)
            if not attach:
                continue
            selected_boxes.append(component_box)
            current_union = pixel_box_union(selected_boxes) or current_union
            changed = True

    analysis_bbox = pixel_box_union(selected_boxes)
    if analysis_bbox is None:
        return None
    full_bbox = scale_pixel_box(analysis_bbox, analysis_size, image.size)
    left = max(0, int(full_bbox[0]) - int(padding))
    top = max(0, int(full_bbox[1]) - int(padding))
    right = min(image.width, int(full_bbox[2]) + int(padding))
    bottom = min(image.height, int(full_bbox[3]) + int(padding))
    if right <= left or bottom <= top:
        return None
    return (left, top, right, bottom)


def crop_pdf_preview_to_visual_content(
    image: Image.Image,
    anchor_box: tuple[int, int, int, int],
    *,
    role: str,
    white_threshold: int = PDF_VISUAL_REFINEMENT_WHITE_THRESHOLD,
    padding: int = PDF_VISUAL_REFINEMENT_PADDING_PX,
) -> Image.Image:
    refined_bbox = refine_pdf_visible_content_bbox(
        image,
        anchor_box,
        role=role,
        white_threshold=white_threshold,
        padding=padding,
    )
    if refined_bbox is None:
        return crop_image_whitespace(image)

    cropped = image.crop(refined_bbox)
    final_mask = pdf_visible_content_mask(cropped, white_threshold=white_threshold)
    final_bbox = final_mask.getbbox()
    if not final_bbox:
        return cropped
    left = max(0, final_bbox[0] - padding)
    top = max(0, final_bbox[1] - padding)
    right = min(cropped.width, final_bbox[2] + padding)
    bottom = min(cropped.height, final_bbox[3] + padding)
    return cropped.crop((left, top, right, bottom))


def render_alt_preview(
    pdf_path: Path,
    page_number: int,
    bbox_norm: dict | None,
    scale: float = 2.0,
    *,
    role: str = "image",
    min_pad_x: float = 10.0,
    min_pad_y: float = 10.0,
    relative_pad_x: float = 0.2,
    relative_pad_y: float = 0.2,
) -> bytes | None:
    if bbox_norm is None:
        return None

    with fitz.open(str(pdf_path)) as document:
        if page_number < 0 or page_number >= len(document):
            return None

        page = document.load_page(page_number)
        clip = preview_clip_rect(
            page.rect,
            bbox_norm,
            min_pad_x=min_pad_x,
            min_pad_y=min_pad_y,
            relative_pad_x=relative_pad_x,
            relative_pad_y=relative_pad_y,
        )
        if clip is None:
            return None

        pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip, alpha=False)
        preview_bytes = pixmap.tobytes("png")
        return tighten_pdf_preview_png(preview_bytes) or preview_bytes


def build_alt_preview_images(
    rows: list[dict],
    source_path: Path | None,
    pdf_path: Path | None,
    roles: set[str] | None = None,
) -> dict[int, dict]:
    preview_images: dict[int, dict] = {}
    normalized_roles = {str(role).lower() for role in roles} if roles is not None else None
    for row in rows:
        if normalized_roles is not None and str(row.get("role", "")).lower() not in normalized_roles:
            continue
        row_id = row.get("id")
        if not isinstance(row_id, int):
            continue
        preview_entry = build_alt_preview_entry(row, source_path, pdf_path)
        if preview_entry is not None:
            preview_images[row_id] = preview_entry
    return preview_images


def normalize_visual_identity_value(value: object) -> str:
    if not isinstance(value, str):
        return ""
    return value.strip().replace("\\", "/").lstrip("/").lower()


def row_visual_identity(row: dict) -> set[str]:
    values = {
        normalize_visual_identity_value(row.get("media_target")),
        normalize_visual_identity_value(row.get("ole_target")),
        normalize_visual_identity_value(row.get("rel_id")),
        normalize_visual_identity_value(row.get("docpr_id")),
    }
    return {value for value in values if value}


def style_item_visual_identity(item: dict) -> set[str]:
    metadata = item.get("metadata") or {}
    values = {
        normalize_visual_identity_value(metadata.get("target")),
        normalize_visual_identity_value(metadata.get("image_target")),
        normalize_visual_identity_value(metadata.get("ole_target")),
        normalize_visual_identity_value(metadata.get("rel_id")),
        normalize_visual_identity_value(metadata.get("image_rid")),
        normalize_visual_identity_value(metadata.get("ole_rid")),
        normalize_visual_identity_value(metadata.get("docpr_id")),
        normalize_visual_identity_value(metadata.get("shape_id")),
    }
    return {value for value in values if value}


def apply_style_map_previews_to_rows(rows: list[dict], style_map: dict | None) -> None:
    items = list((style_map or {}).get("items") or [])
    visual_items = [item for item in items if str(item.get("role", "")).lower() in {"image", "equation"}]
    remaining = list(visual_items)

    def pop_match(predicate) -> dict | None:
        for index, item in enumerate(remaining):
            if predicate(item):
                return remaining.pop(index)
        return None

    for row in rows:
        role = str(row.get("role", "")).lower()
        if role not in {"image", "equation"}:
            continue

        row_identity = row_visual_identity(row)
        match = None
        if row_identity:
            match = pop_match(
                lambda item: str(item.get("role", "")).lower() == role and bool(style_item_visual_identity(item) & row_identity)
            )
        if match is None and row_identity:
            match = pop_match(lambda item: bool(style_item_visual_identity(item) & row_identity))
        if match is None:
            match = pop_match(lambda item: str(item.get("role", "")).lower() == role)

        if not isinstance(match, dict):
            continue

        preview = match.get("preview") or {}
        page = preview.get("page")
        bbox = preview.get("bbox_norm")
        if isinstance(page, int) and isinstance(bbox, dict):
            row["preview_page"] = page
            row["preview_bbox"] = bbox
            row["preview_text"] = preview.get("text", "") or row.get("preview_text", "")
            row["page"] = page + 1


def prepare_docx_preview_context(docx_path: Path, rows: list[dict]) -> Path | None:
    try:
        from Altomizer.converter import docx_to_pdf
        from Altomizer.style_inspector import build_style_map

        pdf_path = Path(docx_to_pdf(str(docx_path)))
        style_map = build_style_map(docx_path, pdf_path)
        if style_map.get("available"):
            apply_style_map_previews_to_rows(rows, style_map)
        return pdf_path
    except Exception:
        return None


def row_has_pdf_preview_anchor(row: dict) -> bool:
    return isinstance(row.get("preview_page"), int) and isinstance(row.get("preview_bbox"), dict)


def normalize_viewport_crop(crop: object) -> dict | None:
    if not isinstance(crop, dict):
        return None

    normalized: dict[str, float] = {}
    for key in ("left", "top", "right", "bottom"):
        raw = crop.get(key)
        try:
            value = float(raw)
        except (TypeError, ValueError):
            value = 0.0
        normalized[key] = max(0.0, min(value, 0.99))

    if all(value <= 0 for value in normalized.values()):
        return None
    if normalized["left"] + normalized["right"] >= 0.98:
        return None
    if normalized["top"] + normalized["bottom"] >= 0.98:
        return None
    return normalized


def apply_viewport_crop_to_png(image_bytes: bytes | None, crop: object) -> bytes | None:
    normalized_crop = normalize_viewport_crop(crop)
    if not image_bytes or normalized_crop is None:
        return image_bytes

    try:
        with Image.open(io.BytesIO(image_bytes)) as image:
            image.load()
            width, height = image.size
            left = int(round(width * normalized_crop["left"]))
            top = int(round(height * normalized_crop["top"]))
            right = int(round(width * (1.0 - normalized_crop["right"])))
            bottom = int(round(height * (1.0 - normalized_crop["bottom"])))
            if right <= left or bottom <= top:
                return image_bytes

            cropped = image.crop((left, top, right, bottom))
            output = io.BytesIO()
            cropped.save(output, format="PNG")
            return output.getvalue()
    except Exception:
        return image_bytes


def apply_display_aspect_crop_to_png(
    image_bytes: bytes | None,
    display_width_pt: object,
    display_height_pt: object,
    *,
    tolerance: float = 0.08,
) -> bytes | None:
    if not image_bytes:
        return image_bytes

    try:
        target_width = float(display_width_pt)
        target_height = float(display_height_pt)
    except (TypeError, ValueError):
        return image_bytes

    if target_width <= 0 or target_height <= 0:
        return image_bytes

    target_aspect = target_width / target_height
    if target_aspect <= 0:
        return image_bytes

    try:
        with Image.open(io.BytesIO(image_bytes)) as image:
            image.load()
            width, height = image.size
            if width <= 0 or height <= 0:
                return image_bytes

            current_aspect = width / height
            if abs(current_aspect - target_aspect) / max(target_aspect, 0.01) <= tolerance:
                return image_bytes

            if current_aspect > target_aspect:
                target_width_px = int(round(height * target_aspect))
                if target_width_px <= 0 or target_width_px >= width:
                    return image_bytes
                left = max(0, (width - target_width_px) // 2)
                cropped = image.crop((left, 0, left + target_width_px, height))
            else:
                target_height_px = int(round(width / target_aspect))
                if target_height_px <= 0 or target_height_px >= height:
                    return image_bytes
                top = max(0, (height - target_height_px) // 2)
                cropped = image.crop((0, top, width, top + target_height_px))

            output = io.BytesIO()
            cropped.save(output, format="PNG")
            return output.getvalue()
    except Exception:
        return image_bytes


def extract_docx_visual_asset_bytes(
    docx_path: Path,
    *targets: object,
) -> tuple[bytes | None, str | None]:
    for target in targets:
        if not isinstance(target, str) or not target.strip():
            continue
        normalized_target = target.strip()
        media_bytes = extract_docx_media(docx_path, normalized_target)
        if isinstance(media_bytes, (bytes, bytearray)) and media_bytes:
            return bytes(media_bytes), normalized_target
        package_bytes = extract_docx_package_bytes(docx_path, normalized_target)
        if isinstance(package_bytes, (bytes, bytearray)) and package_bytes:
            return bytes(package_bytes), normalized_target
    return None, None


def build_docx_media_preview_entry(row: dict, source_path: Path) -> dict | None:
    media_bytes, resolved_target = extract_docx_visual_asset_bytes(
        source_path,
        row.get("media_target"),
        row.get("ole_target"),
    )
    if not media_bytes or not resolved_target:
        return None

    ext = Path(resolved_target).suffix.lower().lstrip(".")
    role = str(row.get("role", "")).lower()
    normalized_png = None
    if ext in {"wmf", "emf"}:
        normalized_png = normalize_metafile_to_png_with_system_drawing(
            media_bytes,
            ext,
            scale=4.0,
            max_pixels=12000000,
            max_edge=4200,
        )
    normalized_png = normalized_png or normalize_image_to_png(media_bytes)
    if normalized_png:
        cropped_png = apply_viewport_crop_to_png(normalized_png, row.get("viewport_crop"))
        preview_png = cropped_png or normalized_png
        if role == "equation":
            preview_png = (
                apply_display_aspect_crop_to_png(
                    preview_png,
                    row.get("display_width_pt"),
                    row.get("display_height_pt"),
                )
                or preview_png
            )
        return finalize_preview_entry_for_row(row, {"bytes": preview_png, "ext": "png"})

    if isinstance(media_bytes, (bytes, bytearray)) and ext in {"png", "jpg", "jpeg", "gif", "bmp", "webp"}:
        return finalize_preview_entry_for_row(row, {"bytes": bytes(media_bytes), "ext": ext})
    return None


def build_alt_preview_entry(row: dict, source_path: Path | None, pdf_path: Path | None) -> dict | None:
    role = str(row.get("role", "")).lower()
    if role == "equation" and isinstance(source_path, Path) and source_path.suffix.lower() == ".docx":
        preview_entry = build_docx_media_preview_entry(row, source_path)
        if preview_entry is not None:
            return preview_entry

    if isinstance(pdf_path, Path) and row_has_pdf_preview_anchor(row):
        preview_kwargs = {}
        if not isinstance(source_path, Path):
            if role == "equation":
                preview_kwargs = PDF_FORMULA_PREVIEW_PADDING
            elif role == "image":
                preview_kwargs = PDF_FIGURE_PREVIEW_PADDING
        preview_bytes = render_alt_preview(
            pdf_path,
            int(row["preview_page"]),
            row.get("preview_bbox"),
            scale=PDF_PREVIEW_RENDER_SCALE,
            role=role,
            **preview_kwargs,
        )
        if preview_bytes:
            return finalize_preview_entry_for_row(row, {"bytes": preview_bytes, "ext": "png"})

    if isinstance(source_path, Path) and source_path.suffix.lower() == ".docx":
        preview_entry = build_docx_media_preview_entry(row, source_path)
        if preview_entry is not None:
            return preview_entry
        if role == "equation":
            hint_text = extract_equation_preview_hint_from_docx(
                source_path,
                row,
                media_bytes=extract_docx_media(source_path, row.get("media_target")) if isinstance(row.get("media_target"), str) else None,
            )
            if hint_text:
                return {"bytes": render_equation_text_tile(hint_text), "ext": "png"}

    title = f"{row.get('type', 'Item')} {row.get('id', '')}".strip()
    body = row.get("label") or row.get("existing_alt_text") or row.get("alt_text") or "Preview unavailable"
    return {"bytes": render_text_tile(title, body), "ext": "png"}


def finalize_preview_entry_for_row(row: dict, preview_entry: dict) -> dict:
    if str(row.get("role", "")).lower() != "equation":
        return preview_entry
    image_bytes = preview_entry.get("bytes")
    if not isinstance(image_bytes, (bytes, bytearray)) or not image_bytes:
        return preview_entry
    refined_png = refine_equation_preview_png(bytes(image_bytes))
    if isinstance(refined_png, (bytes, bytearray)) and equation_preview_has_visible_content(refined_png):
        return {"bytes": bytes(refined_png), "ext": "png"}
    return preview_entry


def flatten_image_onto_white(image: Image.Image) -> Image.Image:
    if "A" in image.getbands():
        base = Image.new("RGBA", image.size, (255, 255, 255, 255))
        base.alpha_composite(image)
        return base.convert("RGB")
    return image.convert("RGB")


def crop_image_whitespace(image: Image.Image, threshold: int = 10) -> Image.Image:
    working = image
    if "A" in working.getbands():
        alpha = working.getchannel("A").point(lambda value: 255 if value > 8 else 0)
        alpha_bbox = alpha.getbbox()
        if alpha_bbox:
            working = working.crop(alpha_bbox)

    rgb = flatten_image_onto_white(working)
    diff = ImageChops.difference(rgb, Image.new("RGB", rgb.size, "white")).convert("L")
    bbox = diff.point(lambda value: 255 if value > threshold else 0).getbbox()
    if not bbox:
        return working

    pad = 8
    left = max(0, bbox[0] - pad)
    top = max(0, bbox[1] - pad)
    right = min(working.width, bbox[2] + pad)
    bottom = min(working.height, bbox[3] + pad)
    return working.crop((left, top, right, bottom))


def refine_equation_preview_png(image_bytes: bytes) -> bytes | None:
    if not image_bytes:
        return None

    try:
        with Image.open(io.BytesIO(image_bytes)) as image:
            image.load()
            source = image.convert("RGBA")
    except Exception:
        return None

    trimmed = crop_image_whitespace(source)
    output = io.BytesIO()
    flatten_image_onto_white(trimmed).save(output, format="PNG")
    return output.getvalue()


def extract_legacy_equation_hint_from_media_bytes(media_bytes: bytes | None) -> str:
    if not media_bytes:
        return ""

    decoded = media_bytes.decode("latin1", errors="ignore")
    if "MathType" not in decoded:
        return ""

    filtered = "".join(ch if 32 <= ord(ch) < 127 else " " for ch in decoded)
    collapsed = re.sub(r"\s+", " ", filtered)
    tail = collapsed.rsplit("_E _A", 1)[-1]
    tail = tail.split('"System', 1)[0]
    tail = tail.strip(" -&")
    if not tail:
        return ""

    def compact_spaced_number(raw: str) -> str:
        return re.sub(r"\s+", "", raw)

    tail = re.sub(
        r'"\s*-\s*"\s*-\s*((?:\d\s*|\.\s*)+)\s*\(\s*\)',
        lambda match: f"-(-{compact_spaced_number(match.group(1))})",
        tail,
    )
    tail = re.sub(
        r'"\s*-\s*((?:\d\s*|\.\s*)+)',
        lambda match: f"-{compact_spaced_number(match.group(1))}",
        tail,
    )

    replacements = (
        ('d"', " >= "),
        ('e"', " <= "),
        ('`"', " != "),
        ('H"', " ~= "),
        ("= =", " = "),
        ("< <", " < "),
        ("> >", " > "),
        ('" - "', "-"),
        ('"-', "-"),
        ("( ", "("),
        (" )", ")"),
        (" . ", "."),
    )
    for old, new in replacements:
        tail = tail.replace(old, new)

    tail = re.sub(r"\b([A-Za-z])\s+([Nn])\s+([Dd])\b", r"\1 and", tail)
    tail = re.sub(r"\b([Aa])\s+n\s+d\b", r"\1 and", tail, flags=re.IGNORECASE)
    tail = re.sub(r"(?<=\d)\s+(?=\d)", "", tail)
    tail = re.sub(r"(?<=\d)\s+\.\s+(?=\d)", ".", tail)
    tail = re.sub(r"\(\s+", "(", tail)
    tail = re.sub(r"\s+\)", ")", tail)
    tail = re.sub(r"\s*,\s*", ", ", tail)
    tail = re.sub(r"\s+", " ", tail).strip(" ,-&")
    tail = tail.replace(" and a ", " and a ")

    if tail.count('"') > 2 or len(tail) < 2:
        return ""
    return tail[:220]


def extract_equation_preview_hint_from_docx(docx_path: Path, row: dict, media_bytes: bytes | None = None) -> str:
    candidates: list[str] = []

    media_hint = extract_legacy_equation_hint_from_media_bytes(media_bytes)
    if media_hint:
        candidates.append(media_hint)

    ole_target = row.get("ole_target")
    if isinstance(ole_target, str) and ole_target.strip():
        ole_bytes = extract_docx_package_bytes(docx_path, ole_target)
        ole_hint = extract_legacy_equation_hint_from_media_bytes(ole_bytes)
        if ole_hint:
            candidates.append(ole_hint)

    cleaned_candidates: list[str] = []
    for candidate in candidates:
        cleaned = normalize_alt_text(candidate)
        cleaned = re.sub(r'"?\bE\s*q\s*u\s*a\s*t\s*i\s*o\s*n\s*N\s*a\s*t\s*i\s*v\s*e\b"?', " ", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(r"\s+", " ", cleaned).strip(" \"'")
        cleaned = cleaned.replace("< =", "<=").replace("> =", ">=").replace("! =", "!=")
        cleaned = cleaned.replace("+ +", "+").replace("- -", "-")
        cleaned = normalize_equation_preview_text(cleaned)
        if cleaned:
            cleaned_candidates.append(cleaned)

    if not cleaned_candidates:
        return ""
    return max(cleaned_candidates, key=lambda text: (equation_quality_score(text), len(text)))


def normalize_equation_preview_text(text: str) -> str:
    cleaned = normalize_alt_text(text)
    if not cleaned:
        return ""

    cleaned = cleaned.replace("Ã—", " x ").replace("Â·", " x ").replace("*", " x ")
    cleaned = cleaned.replace("â‰¤", " <= ").replace("â‰¥", " >= ").replace("Ã·", " / ")
    cleaned = cleaned.replace("< =", "<=").replace("> =", ">=").replace("! =", "!=")
    cleaned = cleaned.replace("+ +", "+").replace("- -", "-")
    cleaned = re.sub(r"(?<=\d)\s*x\s*(?=\d)", " x ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"(?<=\d)\s+(?=[A-Za-z])", " ", cleaned)

    # Legacy MathType hints often flatten scientific notation like "2.961107"
    # or "a 10 n". Render it as OCR-friendly ASCII math instead of prose.
    cleaned = re.sub(
        r"\b([0-9]+(?:\.[0-9]+)?)10\s+([A-Za-z0-9]{1,3})\b",
        lambda match: f"{match.group(1)} x 10^{match.group(2)}",
        cleaned,
    )
    cleaned = re.sub(
        r"\b([A-Za-z0-9.]+)\s+10\s+([A-Za-z0-9]+)\b",
        lambda match: f"{match.group(1)} x 10^{match.group(2)}",
        cleaned,
    )
    cleaned = re.sub(
        r"\b([0-9]+(?:\.[0-9]+)?)10([A-Za-z0-9]{1,3})\b",
        lambda match: f"{match.group(1)} x 10^{match.group(2)}",
        cleaned,
    )
    cleaned = re.sub(
        r"\b([A-Za-z])10([A-Za-z0-9]{1,3})\b",
        lambda match: f"{match.group(1)} x 10^{match.group(2)}",
        cleaned,
    )

    cleaned = re.sub(r"\b([0-9]+(?:\.[0-9]+)?)\s*x\s*10\^([A-Za-z0-9]+)\b", r"\1 x 10^\2", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b([A-Za-z])\s*x\s*10\^([A-Za-z0-9]+)\b", r"\1 x 10^\2", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b([0-9A-Za-z]+)\s*<=\s*([0-9A-Za-z]+)\b", r"\1 <= \2", cleaned)
    cleaned = re.sub(r"\b([0-9A-Za-z]+)\s*>=\s*([0-9A-Za-z]+)\b", r"\1 >= \2", cleaned)
    cleaned = re.sub(r"\b([0-9A-Za-z]+)\s*!=\s*([0-9A-Za-z]+)\b", r"\1 != \2", cleaned)
    cleaned = re.sub(r"(?<=[0-9A-Za-z)])\s*([=+\-/<>\^])\s*", r" \1 ", cleaned)
    cleaned = re.sub(r"\b10\s*\^\s*([A-Za-z0-9]+)\b", r"10^\1", cleaned)
    cleaned = re.sub(r"\bx\s+10\s*\^\s*([A-Za-z0-9]+)\b", r"x 10^\1", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" \t\r\n,;:")
    cleaned = cleaned.replace("< =", "<=").replace("> =", ">=").replace("! =", "!=")
    cleaned = cleaned.replace("= =", "=")

    if not cleaned:
        return ""
    if is_low_quality_equation_text(cleaned):
        return ""
    return cleaned


def equation_preview_has_visible_content(image_bytes: bytes | None) -> bool:
    if not image_bytes:
        return False

    try:
        with Image.open(io.BytesIO(image_bytes)) as image:
            image.load()
            rgb = flatten_image_onto_white(image.convert("RGBA"))
    except Exception:
        return False

    diff = ImageChops.difference(rgb, Image.new("RGB", rgb.size, "white")).convert("L")
    bbox = diff.point(lambda value: 255 if value > 18 else 0).getbbox()
    if not bbox:
        return False

    ink_width = bbox[2] - bbox[0]
    ink_height = bbox[3] - bbox[1]
    if ink_width < 18 or ink_height < 12:
        return False
    return True


def render_equation_text_tile(body: str, width: int = 760, height: int = 220) -> bytes:
    document = fitz.open()
    page = document.new_page(width=width, height=height)
    page.draw_rect(fitz.Rect(0, 0, width, height), color=(0.82, 0.78, 0.72), fill=(0.99, 0.98, 0.96), width=1)
    page.insert_textbox(
        fitz.Rect(24, 26, width - 24, height - 22),
        normalize_alt_text(body or "Legacy equation preview"),
        fontname="Times-Roman",
        fontsize=26,
        color=(0.1, 0.1, 0.1),
        align=0,
    )
    pixmap = page.get_pixmap(matrix=fitz.Matrix(1.4, 1.4), alpha=False)
    return pixmap.tobytes("png")


def suggest_alt_text_for_row(
    row: dict,
    pdf_path: Path,
    docx_path: Path | None,
    rows: list[dict],
    row_index: int,
    omml_candidates: list[str],
    equation_order_map: dict[int, int],
) -> str:
    role = str(row.get("role", "")).lower()
    label = normalize_alt_text(row.get("label", ""))
    display_page = row.get("page")

    if role == "equation":
        ordinal = equation_order_map.get(row_index)
        if ordinal is not None and ordinal < len(omml_candidates):
            omml_text = omml_candidates[ordinal]
            if is_high_confidence_equation_text(omml_text):
                finalized = finalize_alt_text(omml_text, role, display_page)
                if finalized:
                    return finalized

        if is_high_confidence_equation_text(label):
            finalized = finalize_alt_text(label, role, display_page)
            if finalized:
                return finalized

    if role == "image" and label and label.lower() not in {"image", "equation"}:
        if is_useful_image_label(label):
            finalized = finalize_alt_text(label, role, display_page)
            if finalized:
                return finalized
    elif label and label.lower() not in {"image", "equation"}:
        finalized = finalize_alt_text(label, role, display_page)
        if finalized:
            return finalized

    page_number = row.get("preview_page")
    bbox = row.get("preview_bbox")
    extracted = ""
    if isinstance(page_number, int) and isinstance(bbox, dict):
        if role == "equation":
            extracted = normalize_alt_text(
                extract_best_equation_text(
                    pdf_path,
                    page_number,
                    bbox,
                    docx_path=docx_path,
                    media_target=row.get("media_target"),
                    ole_target=row.get("ole_target"),
                    viewport_crop=row.get("viewport_crop"),
                    display_width_pt=row.get("display_width_pt"),
                    display_height_pt=row.get("display_height_pt"),
                )
            )
        else:
            extracted = normalize_alt_text(
                extract_best_image_alt_text(
                    pdf_path,
                    page_number,
                    bbox,
                    docx_path=docx_path,
                    media_target=row.get("media_target"),
                    viewport_crop=row.get("viewport_crop"),
                )
            )

    if extracted:
        finalized = finalize_alt_text(extracted, role, display_page)
        if finalized:
            return finalized

    if role == "equation":
        return ""

    if isinstance(display_page, int):
        return f"Educational figure with visual information on page {display_page}."
    return "Educational figure with visual information."


def is_useful_image_label(label: str) -> bool:
    normalized = normalize_alt_text(label)
    if not normalized:
        return False
    lowered = normalized.lower()
    if looks_like_paragraph_noise(normalized):
        return False
    words = lowered.split()
    if len(words) > 14:
        return False
    if any(marker in lowered for marker in ("chart", "graph", "plot", "diagram", "figure", "photo", "image")):
        return True
    if all(token.isupper() and len(token) <= 4 for token in normalized.replace("=", " ").replace("+", " ").split() if token.isalpha()):
        return True
    return False


def extract_best_image_alt_text(
    pdf_path: Path,
    page_number: int,
    bbox_norm: dict,
    docx_path: Path | None = None,
    media_target: str | None = None,
    viewport_crop: dict | None = None,
) -> str:
    image_bytes = get_image_crop_png_bytes(
        pdf_path,
        page_number,
        bbox_norm,
        docx_path=docx_path,
        media_target=media_target,
        viewport_crop=viewport_crop,
    )
    if not image_bytes:
        return ""

    candidates: list[str] = []
    for psm in ("11", "6", "7"):
        text = compact_ocr_text(run_tesseract_on_png_bytes(image_bytes, psm_modes=(psm,)))
        cleaned = clean_image_ocr_candidate(text)
        if cleaned:
            candidates.append(cleaned)

    best_text = ""
    if not candidates:
        best_text = ""
    else:
        scored = sorted(
            ((image_text_quality_score(candidate), candidate) for candidate in candidates),
            key=lambda item: item[0],
            reverse=True,
        )
        best_text = scored[0][1]

    vision_alt = caption_image_with_local_vlm(image_bytes, best_text)
    if vision_alt:
        return vision_alt

    if best_text:
        return summarize_ordinary_image_alt(best_text)

    return "Educational figure with visual information."


def is_image_vlm_enabled() -> bool:
    raw = (os.getenv("MATCHA_ALT_IMAGE_VLM_ENABLED") or "").strip().lower()
    if not raw:
        return True
    return raw in {"1", "true", "yes", "on"}


def sanitize_generated_image_alt(text: str) -> str:
    cleaned = strip_llm_wrappers(normalize_alt_text(text))
    if not cleaned:
        return ""
    cleaned = re.sub(
        r"^(image showing|image of|picture of|photo of|this image shows|the image shows|there is|there are)\s+",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )
    cleaned = cleaned.strip().strip('"').strip("'")
    cleaned = strip_forbidden_alt_prefixes(cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    if not cleaned:
        return ""
    if looks_like_filename_metadata_text(cleaned):
        return ""
    if len(cleaned.split()) < 3:
        return ""
    sentence_limit = 20 if is_data_dense_image_alt(cleaned) else 2
    sentence_parts = re.split(r"(?<=[.!?])\s+", cleaned)
    cleaned = " ".join(part.strip() for part in sentence_parts[:sentence_limit] if part.strip()).strip()
    if looks_like_paragraph_noise(cleaned):
        return ""
    if any(
        marker in cleaned.lower()
        for marker in (
            "please provide",
            "i will create",
            "i can create",
            "accessibility alt text",
            "final alt text",
            "educational image description",
            "placeholder",
        )
    ):
        return ""
    if any(marker in cleaned.lower() for marker in ("answer:", "hint:", "chapter", "exercise", "question")):
        return ""
    return cleaned


def caption_image_with_local_vlm(image_bytes: bytes, ocr_hint: str = "") -> str:
    if not image_bytes:
        return ""
    if not is_image_vlm_enabled():
        return ""
    if not is_ollama_available():
        return ""

    model = (os.getenv("MATCHA_ALT_IMAGE_VLM_MODEL") or "llava:7b").strip()
    if not model:
        return ""

    prompt = textwrap.dedent(
        f"""
        Write one sentence of accessibility alt text for this educational image.
        Focus on what is visually happening and the key relationship or trend.
        Be specific and concise, about 12 to 28 words.
        Do not include surrounding page questions, answers, or chapter text.
        Do not start with "image of" or "picture of".
        If the image is a chart or diagram, mention the chart type and the key takeaway.
        If the image is a table or spreadsheet, mention the structure, headers, and the most important values or comparisons.
        OCR hint: {ocr_hint[:180]}
        """
    ).strip()

    try:
        encoded = base64.b64encode(image_bytes).decode("ascii")
        response = requests.post(
            "http://localhost:11434/api/generate",
            json={
                "model": model,
                "prompt": prompt,
                "images": [encoded],
                "stream": False,
                "options": {
                    "temperature": 0.15,
                    "top_p": 0.4,
                    "num_predict": 72,
                },
            },
            timeout=25,
        )
        if response.status_code != 200:
            return ""
        payload = response.json()
        raw = str(payload.get("response") or "")
        return sanitize_generated_image_alt(raw)
    except Exception:
        return ""


def get_image_crop_png_bytes(
    pdf_path: Path | None,
    page_number: int,
    bbox_norm: dict,
    docx_path: Path | None = None,
    media_target: str | None = None,
    viewport_crop: dict | None = None,
) -> bytes | None:
    if isinstance(docx_path, Path) and isinstance(media_target, str):
        image_bytes, _ = extract_docx_visual_asset_bytes(docx_path, media_target)
        png_bytes = normalize_image_to_png(image_bytes)
        if png_bytes:
            return apply_viewport_crop_to_png(png_bytes, viewport_crop) or png_bytes

    try:
        if not isinstance(pdf_path, Path):
            return None
        with fitz.open(str(pdf_path)) as document:
            if page_number < 0 or page_number >= len(document):
                return None
            page = document.load_page(page_number)
            clip = preview_clip_rect(page.rect, bbox_norm)
            if clip is None:
                return None
            pixmap = page.get_pixmap(matrix=fitz.Matrix(3.5, 3.5), clip=clip, alpha=False)
            return pixmap.tobytes("png")
    except Exception:
        return None


def clean_image_ocr_candidate(text: str) -> str:
    cleaned = compact_ocr_text(text)
    if not cleaned:
        return ""
    cleaned = strip_llm_wrappers(cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    if not cleaned:
        return ""
    return cleaned[:320]


def looks_like_paragraph_noise(text: str) -> bool:
    lowered = text.lower()
    words = lowered.split()
    chart_cues = (
        "table",
        "spreadsheet",
        "frequency",
        "tally",
        "class limits",
        "column",
        "columns",
        "row",
        "rows",
        "gdp",
        "epi",
        "score",
        "quantity extracted",
        "tc",
        "ec",
        "uc",
        "dollars",
    )
    if any(cue in lowered for cue in chart_cues):
        return False
    if len(words) > 28:
        return True
    if "?" in lowered and len(words) > 12:
        return True
    noise_markers = (
        "answer",
        "answers",
        "explain",
        "hint",
        "question",
        "chapter",
        "would",
        "should",
        "future profits",
        "in each case",
        "does this",
    )
    return any(marker in lowered for marker in noise_markers)


def image_text_quality_score(text: str) -> float:
    lowered = text.lower()
    words = lowered.split()
    score = 0.0

    if looks_like_paragraph_noise(text):
        score -= 2.5

    if len(words) <= 20:
        score += 0.7
    elif len(words) <= 30:
        score += 0.2

    if any(token in lowered for token in ("gdp", "epi", "score", "quantity", "dollars", "tc", "ec", "uc")):
        score += 1.5
    if any(token in lowered for token in ("chart", "graph", "plot")):
        score += 0.4
    if re.search(r"\bq[0-9]+\b", lowered):
        score += 0.5

    return score


def choose_alt_phrase(seed_text: str, options: tuple[str, ...]) -> str:
    if not options:
        return ""
    checksum = sum(ord(char) for char in (seed_text or ""))
    return options[checksum % len(options)]


def summarize_ordinary_image_alt(ocr_text: str) -> str:
    lowered = ocr_text.lower()
    if not lowered:
        return ""

    number_line_values = [int(token) for token in re.findall(r"(?<!\w)-?\d{1,2}(?!\w)", lowered)]
    if number_line_values:
        unique_values = sorted(set(number_line_values))
        if len(unique_values) >= 8 and min(unique_values) <= -5 and max(unique_values) >= 5:
            return (
                "Number line with highlighted endpoints and curved step arrows, illustrating movement across integer values."
            )

    if "gdp" in lowered and ("epi" in lowered or "score" in lowered):
        prefix = choose_alt_phrase(
            lowered,
            (
                "Scatter plot showing",
                "Scatter chart illustrating",
                "Scatter plot highlighting",
            ),
        )
        return f"{prefix} EPI score versus GDP per person, with a clear upward trend line."
    if "gdp per person" in lowered or ("gdp" in lowered and any(name in lowered for name in ("states", "greece", "colombia", "china", "iraq"))):
        prefix = choose_alt_phrase(
            lowered,
            (
                "Scatter plot showing",
                "Scatter chart mapping",
                "Scatter plot tracing",
            ),
        )
        return f"{prefix} data points against GDP per person, with an overall upward trend."

    if all(token in lowered for token in ("tc", "ec", "uc")) and "quantity" in lowered:
        prefix = choose_alt_phrase(
            lowered,
            (
                "Line chart comparing",
                "Line chart showing",
                "Line graph contrasting",
            ),
        )
        if any(token in lowered for token in ("q0", "q1", "q2", "first-year")):
            return (
                f"{prefix} TC = EC + UC, EC, and UC across first-year quantity extracted; "
                "the shift indicates how user cost changes the extraction level."
            )
        return f"{prefix} curves labeled TC = EC + UC, EC, and UC versus quantity extracted."

    if "quantity" in lowered and ("tc" in lowered or "ec" in lowered):
        prefix = choose_alt_phrase(
            lowered,
            (
                "Line chart showing",
                "Line graph describing",
                "Line chart tracing",
            ),
        )
        return f"{prefix} cost curves against quantity extracted."

    if any(token in lowered for token in ("chart", "graph", "plot", "axis", "axes")):
        prefix = choose_alt_phrase(
            lowered,
            (
                "Chart with",
                "Graph with",
                "Plot with",
            ),
        )
        return f"{prefix} labeled axes and plotted data."

    if looks_like_paragraph_noise(ocr_text):
        return "Educational diagram with labeled elements and visual relationships."

    label_summary = summarize_labels_from_ocr(ocr_text)
    if label_summary:
        return label_summary

    return "Educational figure with key visual details."


def summarize_labels_from_ocr(ocr_text: str) -> str:
    tokens = re.findall(r"[A-Za-z][A-Za-z\-]{1,}", ocr_text or "")
    if not tokens:
        return ""

    stopwords = {
        "the",
        "and",
        "for",
        "with",
        "from",
        "this",
        "that",
        "are",
        "was",
        "were",
        "into",
        "onto",
        "over",
        "under",
        "chart",
        "graph",
        "plot",
        "image",
        "figure",
    }
    selected: list[str] = []
    seen = set()
    for token in tokens:
        lowered = token.lower()
        if lowered in stopwords:
            continue
        if len(lowered) <= 2:
            continue
        if lowered in seen:
            continue
        seen.add(lowered)
        selected.append(token)
        if len(selected) == 5:
            break

    if not selected:
        return ""
    return f"Educational diagram with labels including {', '.join(selected)}."


def extract_best_equation_text(
    pdf_path: Path,
    page_number: int,
    bbox_norm: dict,
    docx_path: Path | None = None,
    media_target: str | None = None,
    ole_target: str | None = None,
    viewport_crop: dict | None = None,
    display_width_pt: object = None,
    display_height_pt: object = None,
) -> str:
    try:
        if page_number < 0:
            return ""
        stage1_image = get_equation_crop_png_bytes(
            pdf_path,
            page_number,
            bbox_norm,
            docx_path=docx_path,
            media_target=media_target,
            ole_target=ole_target,
            viewport_crop=viewport_crop,
            display_width_pt=display_width_pt,
            display_height_pt=display_height_pt,
        )
        candidates: list[str] = []
        if stage1_image:
            candidates.extend(
                [
                    compact_ocr_text(run_tesseract_on_png_bytes(stage1_image, psm_modes=("7", "6", "11"))),
                    compact_ocr_text(
                        run_tesseract_on_png_bytes(
                            stage1_image,
                            psm_modes=("7", "6"),
                            extra_configs=(
                                "tessedit_char_whitelist=0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ+-=^/().,×÷·−≤≥",
                                "classify_bln_numeric_mode=1",
                            ),
                        )
                    ),
                    compact_ocr_text(
                        run_tesseract_on_png_bytes(
                            stage1_image,
                            psm_modes=("7",),
                            extra_configs=(
                                "tessedit_char_whitelist=0123456789+-=^/().,",
                                "classify_bln_numeric_mode=1",
                            ),
                        )
                    ),
                ]
            )
        with fitz.open(str(pdf_path)) as document:
            if page_number >= len(document):
                return ""
            page = document.load_page(page_number)
            use_tight_pdf_formula_crop = not isinstance(docx_path, Path) and not media_target and not ole_target
            if use_tight_pdf_formula_crop:
                clip = preview_clip_rect(page.rect, bbox_norm, **PDF_FORMULA_PREVIEW_PADDING)
            else:
                clip = preview_clip_rect(page.rect, bbox_norm)

            if clip is not None:
                ocr_clip = clip
                if not use_tight_pdf_formula_crop:
                    ocr_clip = fitz.Rect(
                        max(page.rect.x0, clip.x0 - max(24, clip.width * 0.35)),
                        max(page.rect.y0, clip.y0 - max(16, clip.height * 0.35)),
                        min(page.rect.x1, clip.x1 + max(24, clip.width * 0.35)),
                        min(page.rect.y1, clip.y1 + max(16, clip.height * 0.35)),
                    )
                region_pixmap = page.get_pixmap(matrix=fitz.Matrix(4.0, 4.0), clip=ocr_clip, alpha=False)
                candidates.append(
                    compact_ocr_text(
                        run_tesseract_on_png_bytes(
                            region_pixmap.tobytes("png"),
                            psm_modes=("7", "6", "11"),
                            extra_configs=(
                                "tessedit_char_whitelist=0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ+-=^/().,×÷·−≤≥",
                                "classify_bln_numeric_mode=1",
                            ),
                        )
                    )
                )

            region_ocr = compact_ocr_text(
                ocr_pdf_region_text(
                    pdf_path,
                    page_number,
                    bbox_norm,
                    scale=4.0,
                    expansion=0.02 if use_tight_pdf_formula_crop else 0.3,
                )
            )
            if region_ocr:
                candidates.append(region_ocr)
    except Exception:
        return ""

    cleaned = []
    for candidate in candidates:
        text = clean_equation_candidate(candidate)
        if text:
            cleaned.append(text)

    if not cleaned:
        return ""

    scored = sorted(((equation_quality_score(text), text) for text in cleaned), key=lambda item: item[0], reverse=True)
    best_score, best_text = scored[0]
    strict = normalize_equation_candidate_to_alt(best_text)
    if strict:
        return strict

    fallback = best_effort_equation_alt_text(best_text)
    if fallback:
        return fallback
    return clean_equation_candidate(best_text)


def get_equation_crop_png_bytes(
    pdf_path: Path | None,
    page_number: int,
    bbox_norm: dict,
    docx_path: Path | None = None,
    media_target: str | None = None,
    ole_target: str | None = None,
    viewport_crop: dict | None = None,
    display_width_pt: object = None,
    display_height_pt: object = None,
) -> bytes | None:
    if isinstance(docx_path, Path):
        image_bytes, resolved_target = extract_docx_visual_asset_bytes(docx_path, media_target, ole_target)
        ext = Path(resolved_target or media_target or ole_target or "").suffix.lower().lstrip(".")
        png_bytes = normalize_metafile_to_png_with_system_drawing(
            image_bytes,
            ext,
            scale=4.5,
            max_pixels=18000000,
            max_edge=5200,
        )
        png_bytes = png_bytes or normalize_image_to_png(image_bytes)
        if png_bytes:
            cropped_png = apply_viewport_crop_to_png(png_bytes, viewport_crop) or png_bytes
            aspect_png = apply_display_aspect_crop_to_png(cropped_png, display_width_pt, display_height_pt) or cropped_png
            refined_png = refine_equation_preview_png(aspect_png)
            if refined_png:
                return refined_png
            return aspect_png

    try:
        if not isinstance(pdf_path, Path):
            return None
        with fitz.open(str(pdf_path)) as document:
            if page_number >= 0 and page_number < len(document):
                page = document.load_page(page_number)
                clip = preview_clip_rect(page.rect, bbox_norm, **PDF_FORMULA_PREVIEW_PADDING)
                if clip is not None:
                    pixmap = page.get_pixmap(matrix=fitz.Matrix(4.0, 4.0), clip=clip, alpha=False)
                    return pixmap.tobytes("png")
    except Exception:
        pass
    return None


def ocr_equation_with_vlm(
    image_bytes: bytes | None,
    page_number: int | None = None,
    bbox_norm: dict | None = None,
    pdf_path: Path | None = None,
) -> str:
    return ""


def _strip_equation_scaffold_noise(text: str) -> str:
    cleaned = compact_ocr_text(text)
    if not cleaned:
        return ""

    cleaned = re.sub(r"^\s*\d+\.\s*", "", cleaned)
    cleaned = re.sub(r"^\s*(?:->|=>|→|←|↔)\s*", "", cleaned)
    cleaned = re.sub(r"\s*(?:->|=>|→|←|↔)\s*$", "", cleaned)
    cleaned = re.sub(r"^\s*[•·▪◦]+\s*", "", cleaned)
    cleaned = re.sub(r"\s*[•·▪◦]+\s*$", "", cleaned)
    cleaned = re.sub(r"\s*(?:->|=>|→|←|↔)\s*", " ", cleaned)
    return " ".join(cleaned.split())


def _normalize_equation_blank_markers(text: str) -> str:
    cleaned = compact_ocr_text(text)
    if not cleaned:
        return ""

    blank_patterns = (
        r"(?<![A-Za-z0-9])(?:_\s*){2,}(?![A-Za-z0-9])",
        r"(?<![A-Za-z0-9])(?:[-–—]\s*){3,}(?![A-Za-z0-9])",
    )
    for pattern in blank_patterns:
        cleaned = re.sub(pattern, " ___ ", cleaned)
    return " ".join(cleaned.split())


def _collapse_duplicate_equation_operators(text: str) -> str:
    cleaned = compact_ocr_text(text)
    if not cleaned:
        return ""

    for _ in range(4):
        updated = re.sub(
            r"(?<=[A-Za-z0-9_)\]])\s*([+=])\s*(?:\1\s*)+(?=[A-Za-z0-9_(\[]|___)",
            r" \1 ",
            cleaned,
        )
        if updated == cleaned:
            break
        cleaned = updated
    return " ".join(cleaned.split())


def _normalize_equation_operator_spacing(text: str) -> str:
    cleaned = compact_ocr_text(text)
    if not cleaned:
        return ""

    cleaned = re.sub(r"(?<=\d)-(?=\d)", " - ", cleaned)
    cleaned = re.sub(r"(?<=[A-Za-z0-9_)])-(?=[A-Za-z_(])", " - ", cleaned)
    cleaned = re.sub(r"\s*([=+/^<>])\s*", r" \1 ", cleaned)
    cleaned = re.sub(r"\s*([×÷·])\s*", r" \1 ", cleaned)
    return " ".join(cleaned.split())


def _extract_equation_core(text: str) -> str:
    cleaned = compact_ocr_text(text)
    if not cleaned:
        return ""

    spans = re.findall(r"[A-Za-z0-9_+\-=/^().,<>≤≥×÷·\s]+", cleaned)
    candidates: list[str] = []
    for span in spans:
        candidate = " ".join(span.split()).strip(" \t\r\n,;:")
        if not candidate:
            continue
        if has_raw_equation_signal(candidate) or looks_equation_like(candidate):
            candidates.append(candidate)

    if not candidates:
        return cleaned
    return max(candidates, key=lambda candidate: (equation_quality_score(candidate), len(candidate)))


def strict_equation_transcription(candidate_text: str) -> str:
    text = compact_ocr_text(candidate_text)
    if not text:
        return ""

    text = strip_llm_wrappers(text)
    text = text.replace("â‰¤", "≤").replace("â‰¥", "≥")
    text = text.replace("Ã—", "×").replace("Â·", "·").replace("Ã·", "÷")
    text = text.replace("âˆ’", "-").replace("â€“", "-").replace("â€”", "-")
    text = re.sub(r"\(page\s+\d+\)", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\bpage\s+\d+\b", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\banswers?\b.*$", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\bchapter\b.*$", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\bspss\b.*$", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\bcont\.?\b.*$", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(document|image)\b.*$", "", text, flags=re.IGNORECASE)
    text = _strip_equation_scaffold_noise(text)
    text = _normalize_equation_blank_markers(text)
    text = _extract_equation_core(text)
    text = re.sub(r"[^0-9A-Za-z_+\-=/^().,<>≤≥×÷·\s]", " ", text)
    text = _collapse_duplicate_equation_operators(text)
    text = _normalize_equation_operator_spacing(text)
    text = _normalize_equation_blank_markers(text)
    text = re.sub(r"\s+", " ", text).strip(" \t\r\n,;:")
    if not text:
        return ""
    if not has_raw_equation_signal(text) and not looks_equation_like(text):
        return ""
    return text


def normalize_equation_candidate_to_alt(candidate_text: str) -> str:
    return strict_equation_transcription(candidate_text)


def best_effort_equation_alt_text(candidate_text: str) -> str:
    return strict_equation_transcription(candidate_text)


def transcribe_equation_alt_text(candidate_text: str) -> str:
    return strict_equation_transcription(candidate_text)


def formula_latex_to_spoken_alt_text(latex: str, kind: str | None = None) -> str:
    text = normalize_alt_text(latex)
    if not text:
        return ""
    text = text.replace("\\left", "").replace("\\right", "")
    text = re.sub(r"\\frac\s*\{\s*([^{}]+)\s*\}\s*\{\s*([^{}]+)\s*\}", r"\1/\2", text)
    text = re.sub(r"\\sqrt\s*\{\s*([^{}]+)\s*\}", r"sqrt(\1)", text)
    text = text.replace("\\", " ")
    text = text.replace("{", " ").replace("}", " ")
    text = text.replace("(", " ").replace(")", " ")
    text = text.replace("[", " ").replace("]", " ")
    text = text.replace(",", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return strict_equation_transcription(text)


def finalize_spoken_equation_alt_text(text: str) -> str:
    spoken = strict_equation_transcription(text)
    if not spoken:
        return ""
    return finalize_alt_text(spoken, "equation", None)


def has_raw_equation_signal(text: str) -> bool:
    candidate = compact_ocr_text(text)
    if not candidate:
        return False
    if any(marker in candidate for marker in ("=", "^", "_", "\\frac", "\\sqrt", "≤", "≥", "×", "÷", "·", "−", "–", "—")):
        return True
    if re.search(r"\b[a-zA-Z]\s*[\^_]\s*\d", candidate):
        return True
    if re.search(r"\d\s*[\/=+\-]\s*\d", candidate):
        return True
    if re.search(r"\b[a-zA-Z]\d+\b", candidate):
        return True
    return False


def equation_symbols_to_spoken_words(text: str) -> str:
    working = compact_ocr_text(text)
    if not working:
        return ""

    working = _normalize_equation_blank_markers(working)
    working = normalize_decimal_point_speech(working)
    working = re.sub(r"(?<=\d)-(?=\d)", " - ", working)
    working = re.sub(r"(?<=[A-Za-z0-9_)])-(?=[A-Za-z_(])", " - ", working)
    working = re.sub(r"(?<![A-Za-z0-9])_{2,}(?![A-Za-z0-9])", " blank ", working)
    working = working.replace(">=", " greater than or equal to ")
    working = working.replace("<=", " less than or equal to ")
    working = working.replace("!=", " not equal to ")
    working = working.replace("<>", " not equal to ")
    working = working.replace("==", " equals ")
    working = working.replace("->", " maps to ")
    working = working.replace("≤", " less than or equal to ")
    working = working.replace("≥", " greater than or equal to ")
    working = working.replace("×", " times ")
    working = working.replace("*", " times ")
    working = working.replace("·", " times ")
    working = working.replace("÷", " divided by ")
    working = working.replace("−", " minus ")
    working = working.replace("–", " minus ")
    working = working.replace("—", " minus ")
    working = working.replace(">", " greater than ")
    working = working.replace("<", " less than ")
    working = working.replace("=", " equals ")
    working = working.replace("+", " plus ")
    working = working.replace("/", " over ")
    working = re.sub(r"\s-\s", " minus ", working)

    working = re.sub(r"\\frac\s*\{\s*([^{}]+)\s*\}\s*\{\s*([^{}]+)\s*\}", r"\1 over \2", working)
    working = re.sub(r"\\sqrt\s*\{\s*([^{}]+)\s*\}", r"square root of \1", working)

    working = re.sub(r"(\b[a-zA-Z]\b)\s*\^\s*\{\s*2\s*\}", r"\1 squared", working)
    working = re.sub(r"(\b[a-zA-Z]\b)\s*\^\s*2\b", r"\1 squared", working)
    working = re.sub(r"(\b[a-zA-Z]\b)\s*\^\s*\{\s*3\s*\}", r"\1 cubed", working)
    working = re.sub(r"(\b[a-zA-Z]\b)\s*\^\s*3\b", r"\1 cubed", working)
    working = re.sub(r"(\b[a-zA-Z]\b)\s*\^\s*\{\s*([0-9]+)\s*\}", lambda m: f"{m.group(1)} to the {integer_to_words(m.group(2))}", working)
    working = re.sub(r"(\b[a-zA-Z]\b)\s*\^\s*([0-9]+)\b", lambda m: f"{m.group(1)} to the {integer_to_words(m.group(2))}", working)

    working = re.sub(r"\b([0-9]+)\b", lambda m: integer_to_words(m.group(1)), working)
    working = re.sub(r"\s+", " ", working).strip()
    working = re.sub(r"\bover equals\b", "over equals", working)
    return working


def looks_like_equation_transcription(text: str) -> bool:
    lowered = compact_ocr_text(text).lower()
    if not lowered:
        return False
    if any(marker in lowered for marker in ("showing", "centered", "graph", "page", "context", "transverse axis", "answers", "chapter", "spss", "provide", "feedback", "document", "image")):
        return False

    relation_terms = (" equals ", " equal to ", " greater than ", " less than ")
    math_terms = (
        " over ",
        " squared",
        " cubed",
        " times ",
        " divided by ",
        " square root",
        " to the ",
    )
    word_count = len(lowered.split())
    if word_count > 12 and not any(term in f" {lowered} " for term in relation_terms):
        return False
    if any(term in f" {lowered} " for term in relation_terms) and any(term in f" {lowered} " for term in math_terms):
        return True
    if any(term in f" {lowered} " for term in relation_terms) and len(lowered.split()) >= 3:
        return True
    return False


def integer_to_words(value: str) -> str:
    try:
        number = int(value)
    except (TypeError, ValueError):
        return str(value)

    ones = {
        0: "zero",
        1: "one",
        2: "two",
        3: "three",
        4: "four",
        5: "five",
        6: "six",
        7: "seven",
        8: "eight",
        9: "nine",
        10: "ten",
        11: "eleven",
        12: "twelve",
        13: "thirteen",
        14: "fourteen",
        15: "fifteen",
        16: "sixteen",
        17: "seventeen",
        18: "eighteen",
        19: "nineteen",
    }
    tens = {
        20: "twenty",
        30: "thirty",
        40: "forty",
        50: "fifty",
        60: "sixty",
        70: "seventy",
        80: "eighty",
        90: "ninety",
    }

    if number < 0:
        return f"minus {integer_to_words(str(abs(number)))}"
    if number in ones:
        return ones[number]
    if number < 100:
        ten_value = (number // 10) * 10
        remainder = number % 10
        if remainder == 0:
            return tens.get(ten_value, str(number))
        return f"{tens.get(ten_value, str(ten_value))} {ones.get(remainder, str(remainder))}"
    if number < 1000:
        hundred_value = number // 100
        remainder = number % 100
        if remainder == 0:
            return f"{ones.get(hundred_value, str(hundred_value))} hundred"
        return f"{ones.get(hundred_value, str(hundred_value))} hundred {integer_to_words(str(remainder))}"
    if number < 1_000_000:
        thousand_value = number // 1000
        remainder = number % 1000
        if remainder == 0:
            return f"{integer_to_words(str(thousand_value))} thousand"
        return f"{integer_to_words(str(thousand_value))} thousand {integer_to_words(str(remainder))}"
    if number < 1_000_000_000:
        million_value = number // 1_000_000
        remainder = number % 1_000_000
        if remainder == 0:
            return f"{integer_to_words(str(million_value))} million"
        return f"{integer_to_words(str(million_value))} million {integer_to_words(str(remainder))}"
    return str(number)


def normalize_decimal_point_speech(text: str) -> str:
    if not text:
        return ""

    def replace_numeric_decimal(match: re.Match[str]) -> str:
        sign = match.group(1)
        whole = match.group(2) or "0"
        fraction = match.group(3)
        prefix = "negative " if sign == "-" else ""
        fraction_words = " ".join(integer_to_words(digit) for digit in fraction)
        return f"{prefix}{integer_to_words(whole)} point {fraction_words}"

    normalized = re.sub(r"(?<![\w/])(-?)(\d*)\.(\d+)\b", replace_numeric_decimal, text)
    digit_word_pattern = "zero|one|two|three|four|five|six|seven|eight|nine"

    def replace_spoken_decimal(match: re.Match[str]) -> str:
        prefix = "negative " if match.group("negative") else ""
        whole = match.group("whole").lower()
        fraction = match.group("fraction").lower()
        return f"{prefix}{whole} point {fraction}"

    normalized = re.sub(
        rf"\b(?:(?P<negative>negative)\s+)?(?P<whole>{digit_word_pattern})\.(?P<fraction>{digit_word_pattern})\b",
        replace_spoken_decimal,
        normalized,
        flags=re.IGNORECASE,
    )
    return re.sub(r"\s+", " ", normalized).strip()


def looks_like_spoken_math_alt_text(text: str) -> bool:
    lowered = compact_ocr_text(text).lower()
    if not lowered:
        return False
    if any(marker in lowered for marker in ("showing", "centered", "graph", "page", "context", "transverse axis")):
        return False
    math_terms = (
        "equals",
        "plus",
        "minus",
        "over",
        "squared",
        "cubed",
        "times",
        "divided by",
        "square root",
        "to the",
    )
    return any(term in lowered for term in math_terms)


def polish_equation_alt_with_local_llm(
    candidate_text: str,
    row: dict | None,
    pdf_path: Path,
    page_number: int,
    bbox_norm: dict | None,
) -> str:
    return ""


def build_equation_style_examples_block(candidate_text: str, source_filename: str | None = None, limit: int = 3) -> str:
    if (os.getenv("MATCHA_ALT_USE_STYLE_EXAMPLES") or "").strip().lower() not in {"1", "true", "yes", "on"}:
        return ""

    examples = select_alt_style_examples(candidate_text, source_filename=source_filename, limit=limit)
    if not examples:
        return ""

    lines = []
    for example in examples:
        alt_text = normalize_alt_text(example.get("alt_text", "")).lower()
        if alt_text:
            lines.append(f"- {alt_text}")
    return "\n".join(lines)


def refine_equation_alt_with_local_llm(
    candidate_text: str,
    row: dict | None,
    pdf_path: Path,
    page_number: int,
    bbox_norm: dict | None,
) -> str:
    return ""


def is_ollama_available() -> bool:
    return bool(find_ollama_executable())


def find_ollama_executable() -> str | None:
    for candidate in (
        shutil.which("ollama"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Ollama", "ollama.exe"),
    ):
        if candidate and Path(candidate).exists():
            return candidate
    return None


def run_ollama_prompt(prompt: str, model: str = "phi3:mini") -> str:
    return ""


def run_ollama_image_prompt(prompt: str, image_bytes: bytes, model: str = "llava:7b") -> str:
    return ""


def strip_llm_wrappers(text: str) -> str:
    cleaned = compact_ocr_text(text)
    cleaned = cleaned.strip().strip('"').strip("'")
    cleaned = re.sub(r"^(alt text|alt|answer|response)[:\-]\s*", "", cleaned, flags=re.IGNORECASE)
    return cleaned


def is_safe_llm_alt_text(text: str) -> bool:
    cleaned = compact_ocr_text(text)
    if not cleaned:
        return False
    if any(ord(ch) > 126 for ch in cleaned):
        return False
    if any(marker in cleaned for marker in ("\\", "{", "}", "^", "=", "(", ")", "[", "]", "$")):
        return False
    if len(cleaned.split()) < 6:
        return False
    if any(marker in cleaned.lower() for marker in ("showing", "centered", "graph", "page", "transverse axis")):
        return False
    return True


def is_alt_vlm_enabled() -> bool:
    return (os.getenv("MATCHA_ALT_VLM_ENABLED") or "").strip().lower() in {"1", "true", "yes", "on"}


def is_alt_llm_polish_enabled() -> bool:
    return (os.getenv("MATCHA_ALT_LLM_POLISH_ENABLED") or "").strip().lower() in {"1", "true", "yes", "on"}


def ocr_pdf_page_text(pdf_path: Path, page_number: int, scale: float = 2.0) -> str:
    try:
        with fitz.open(str(pdf_path)) as document:
            if page_number < 0 or page_number >= len(document):
                return ""
            page = document.load_page(page_number)
            pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            return run_tesseract_on_png_bytes(pixmap.tobytes("png"), psm_modes=("11", "6"))
    except Exception:
        return ""


def ocr_pdf_region_text(
    pdf_path: Path,
    page_number: int,
    bbox_norm: dict,
    scale: float = 4.0,
    expansion: float = 0.7,
) -> str:
    try:
        with fitz.open(str(pdf_path)) as document:
            if page_number < 0 or page_number >= len(document):
                return ""
            page = document.load_page(page_number)
            clip = preview_clip_rect(page.rect, bbox_norm)
            if clip is None:
                return ""
            width_pad = clip.width * expansion
            height_pad = clip.height * expansion
            wider = fitz.Rect(
                max(page.rect.x0, clip.x0 - width_pad),
                max(page.rect.y0, clip.y0 - height_pad),
                min(page.rect.x1, clip.x1 + width_pad),
                min(page.rect.y1, clip.y1 + height_pad),
            )
            pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=wider, alpha=False)
            return run_tesseract_on_png_bytes(pixmap.tobytes("png"), psm_modes=("6", "11", "12"))
    except Exception:
        return ""


def run_tesseract_on_png_bytes(
    png_bytes: bytes,
    psm_modes: tuple[str, ...] = ("6", "11"),
    extra_configs: tuple[str, ...] = (),
) -> str:
    exe = find_tesseract_executable()
    if not exe or not png_bytes:
        return ""

    tmp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile("wb", suffix=".png", delete=False) as handle:
            handle.write(png_bytes)
            tmp_path = Path(handle.name)

        outputs = []
        for psm in psm_modes:
            command = [exe, str(tmp_path), "stdout", "--psm", psm]
            for config in extra_configs:
                if "=" not in config:
                    continue
                command.extend(["-c", config])
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=60,
                check=False,
            )
            text = " ".join(str(completed.stdout or "").replace("\n", " ").split())
            if text:
                outputs.append(text)
        return " ".join(outputs)
    except Exception:
        return ""
    finally:
        if tmp_path is not None and tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass


def compact_ocr_text(text: str) -> str:
    return " ".join(str(text or "").replace("\n", " ").split())


def find_tesseract_executable() -> str | None:
    for candidate in (
        shutil.which("tesseract"),
        r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
        r"C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe",
    ):
        if candidate and Path(candidate).exists():
            return candidate
    return None


def infer_semantic_equation_alt(page_context: str, crop_text: str, page_number: int | None = None) -> str:
    combined = f"{page_context}\n{crop_text}".lower()
    if "hyperbola" in combined:
        if "transverse axis: x-axis" in combined or "transverse axis x-axis" in combined:
            return "Equation showing the standard form of a hyperbola centered at the origin with transverse axis on the x-axis"
        if "transverse axis: y-axis" in combined or "transverse axis y-axis" in combined:
            return "Equation showing the standard form of a hyperbola centered at the origin with transverse axis on the y-axis"

    normalized_crop = normalize_alt_text(crop_text)
    if not normalized_crop:
        return ""

    if any(token in normalized_crop.lower() for token in ("squared", "over", "equals", "minus", "plus")):
        verbal = normalize_alt_text(normalized_crop)
        verbal = verbal.replace(" x-axis ", " x axis ")
        verbal = verbal.replace(" y-axis ", " y axis ")
        return finalize_alt_text(verbal, "equation", page_number)

    return ""


def extract_omml_equation_candidates(docx_path: Path) -> list[str]:
    try:
        elements = extract_doc_structure(str(docx_path))
    except Exception:
        return []

    candidates = []
    seen = set()
    for element in elements:
        if getattr(element, "role", "") != "equation":
            continue
        metadata = getattr(element, "metadata", None) or {}
        if metadata.get("math_source") != "omml":
            continue
        text = clean_equation_candidate(getattr(element, "text", ""))
        if not text:
            continue
        key = text.lower()
        if key in seen:
            continue
        seen.add(key)
        candidates.append(text)
    return candidates


def find_neighbor_equation_text(rows: list[dict], row_index: int, window: int = 6) -> str:
    for distance in range(1, window + 1):
        for candidate_index in (row_index - distance, row_index + distance):
            if candidate_index < 0 or candidate_index >= len(rows):
                continue
            candidate = rows[candidate_index]
            if str(candidate.get("role", "")).lower() != "equation":
                continue
            text = normalize_alt_text(candidate.get("existing_alt_text") or candidate.get("generated_alt_text") or candidate.get("label") or "")
            if is_high_confidence_equation_text(text):
                return text
    return ""


def is_high_confidence_equation_text(text: str) -> bool:
    cleaned = clean_equation_candidate(text)
    if not cleaned:
        return False
    if equation_quality_score(cleaned) < 1.4:
        return False
    return len(cleaned.split()) >= 4


def extract_text_from_pdf_region(pdf_path: Path, page_number: int, bbox_norm: dict) -> str:
    try:
        with fitz.open(str(pdf_path)) as document:
            if page_number < 0 or page_number >= len(document):
                return ""
            page = document.load_page(page_number)
            clip = preview_clip_rect(page.rect, bbox_norm)
            if clip is None:
                return ""
            return page.get_text("text", clip=clip) or ""
    except Exception:
        return ""


def preview_clip_rect(
    page_rect: fitz.Rect,
    bbox_norm: dict,
    *,
    min_pad_x: float = 10.0,
    min_pad_y: float = 10.0,
    relative_pad_x: float = 0.2,
    relative_pad_y: float = 0.2,
) -> fitz.Rect | None:
    x0 = min(max(float(bbox_norm.get("x0", 0.0)), 0.0), 1.0)
    y0 = min(max(float(bbox_norm.get("y0", 0.0)), 0.0), 1.0)
    x1 = min(max(float(bbox_norm.get("x1", x0)), 0.0), 1.0)
    y1 = min(max(float(bbox_norm.get("y1", y0)), 0.0), 1.0)

    if x1 <= x0:
        x1 = min(1.0, x0 + 0.08)
    if y1 <= y0:
        y1 = min(1.0, y0 + 0.08)

    clip = fitz.Rect(
        page_rect.x0 + (x0 * page_rect.width),
        page_rect.y0 + (y0 * page_rect.height),
        page_rect.x0 + (x1 * page_rect.width),
        page_rect.y0 + (y1 * page_rect.height),
    )

    pad_x = max(0.0, max(float(min_pad_x), clip.width * max(0.0, float(relative_pad_x))))
    pad_y = max(0.0, max(float(min_pad_y), clip.height * max(0.0, float(relative_pad_y))))
    expanded = fitz.Rect(
        max(page_rect.x0, clip.x0 - pad_x),
        max(page_rect.y0, clip.y0 - pad_y),
        min(page_rect.x1, clip.x1 + pad_x),
        min(page_rect.y1, clip.y1 + pad_y),
    )
    if expanded.is_empty:
        return None
    return expanded


def normalize_alt_text(text: str) -> str:
    if not text:
        return ""
    return " ".join(str(text).replace("\n", " ").split())


def clean_equation_candidate(text: str) -> str:
    normalized = normalize_alt_text(text)
    if not normalized:
        return ""

    # Drop obvious bleed-through suffixes that can appear in clipped OCR output.
    normalized = normalized.replace("(page 1).", "").replace("(page 2).", "").strip()
    normalized = normalized.replace(" - - ", " - ").replace("--", "-")
    normalized = " ".join(normalized.split())
    if is_low_quality_equation_text(normalized):
        return ""
    return normalized


def looks_equation_like(text: str) -> bool:
    lowered = text.lower()
    if "=" in text:
        return True
    keywords = ("fraction", "squared", "numerator", "denominator", "plus", "minus")
    if any(keyword in lowered for keyword in keywords):
        return True
    if sum(ch.isalpha() for ch in text) >= 2 and any(ch.isdigit() for ch in text):
        return True
    return False


def is_low_quality_equation_text(text: str) -> bool:
    stripped = text.strip()
    if not stripped:
        return True

    lowered = stripped.lower()
    tokens = stripped.split()
    math_ops = ("=", "+", "/", "^", "×", "÷", "·", "≤", "≥")
    noise_markers = ("answers", "chapter", "spss", "document", "image", "page", "transverse axis", "provide", "feedback")
    math_signal = bool(re.search(r"[=+/^×÷·≤≥]", stripped)) or bool(re.search(r"\b\d+\s*-\s*\d+\b", stripped)) or bool(
        re.search(r"\b[a-zA-Z]\s*-\s*[a-zA-Z\d]+\b", stripped)
    )

    if len(tokens) < 3:
        if re.fullmatch(r"[\d\s.,]+", stripped):
            return False
        if math_signal:
            return False
        if any(char.isdigit() for char in stripped):
            return False
        return True

    if re.fullmatch(r"[\d\s.,]+", stripped):
        return False
    if any(marker in lowered for marker in noise_markers) and not math_signal:
        return True

    single_char_tokens = sum(1 for token in tokens if len(token) == 1 and token.isalnum())
    single_ratio = single_char_tokens / max(1, len(tokens))
    if single_ratio > 0.7:
        return True

    meaningful = sum(1 for token in tokens if len(token) >= 3 and any(char.isalpha() for char in token))
    if any(char.isdigit() for char in stripped) and math_signal:
        return False
    if meaningful == 0 and not math_signal:
        return True
    return False


def equation_quality_score(text: str) -> float:
    score = 0.0
    lowered = text.lower()
    tokens = text.split()

    if re.fullmatch(r"[\d\s.,]+", text):
        score += 1.4
    if any(op in text for op in ("=", "+", "-", "/", "^", "×", "÷", "·", "≤", "≥")):
        score += 1.0
    if "=" in text:
        score += 1.2
    if any(variable in lowered for variable in (" x", "y", "squared", "fraction", "numerator", "denominator")):
        score += 0.9
    if len(tokens) >= 5:
        score += 0.4
    if any(bad in lowered for bad in ("answers", "chapter", "spss", "document", "image", "page", "transverse axis")) and not any(
        op in text for op in ("=", "+", "-", "/", "^", "×", "÷", "·", "≤", "≥")
    ):
        score -= 1.2
    if is_low_quality_equation_text(text):
        score -= 1.3
    score -= max(0, len(text) - 220) * 0.002
    return score


def finalize_alt_text(text: str, role: str, page: int | None) -> str:
    normalized = normalize_alt_text(text)
    if not normalized:
        return ""

    if role == "equation":
        normalized = normalize_decimal_point_speech(normalized)
        if is_low_quality_equation_text(normalized):
            return ""

    if not normalized.endswith((".", "!", "?")):
        normalized = normalized + "."

    if role not in {"equation", "image"} and page is not None and "page" not in normalized.lower():
        if len(normalized) < 250:
            normalized = normalized[:-1] + f" (page {page})."
    return normalized


def extract_docx_media(docx_path: Path, media_target: str | None) -> bytes | None:
    if not media_target:
        return None

    target = media_target.lstrip("/")
    if not target.lower().startswith("word/media/"):
        return None

    try:
        with zipfile.ZipFile(docx_path, "r") as archive:
            return archive.read(target)
    except (KeyError, FileNotFoundError, zipfile.BadZipFile):
        return None


def extract_docx_package_bytes(docx_path: Path, package_target: str | None) -> bytes | None:
    if not package_target:
        return None

    target = package_target.lstrip("/")
    try:
        with zipfile.ZipFile(docx_path, "r") as archive:
            return archive.read(target)
    except (KeyError, FileNotFoundError, zipfile.BadZipFile):
        return None


def normalize_image_to_png(image_bytes: bytes | None) -> bytes | None:
    if not image_bytes:
        return None
    if image_bytes[:8] == b"\x89PNG\r\n\x1a\n":
        return image_bytes

    pixmap = None
    try:
        source_pixmap = fitz.Pixmap(image_bytes)
    except Exception:
        source_pixmap = None

    try:
        if source_pixmap is not None:
            pixmap = source_pixmap

            # PNG export only works for grayscale/RGB pixmaps, so normalize any
            # embedded CMYK or other colorspaces before serialization.
            if pixmap.colorspace not in (fitz.csGRAY, fitz.csRGB) or pixmap.alpha:
                pixmap = fitz.Pixmap(fitz.csRGB, pixmap)

            return pixmap.tobytes("png")
        return normalize_image_to_png_with_pillow(image_bytes)
    finally:
        source_pixmap = None
        pixmap = None


def normalize_image_to_png_with_pillow(image_bytes: bytes | None, dpi: int = 600) -> bytes | None:
    if not image_bytes:
        return None

    try:
        with Image.open(io.BytesIO(image_bytes)) as image:
            image_format = (image.format or "").upper()
            if image_format in {"WMF", "EMF"}:
                # Render vector equation previews at a readable size while
                # preserving the exact single-equation object.
                image.load(dpi=max(300, int(dpi or 600)))
            else:
                image.load()

            if image.mode not in {"RGB", "RGBA"}:
                has_transparency = "A" in image.getbands() or "transparency" in image.info
                image = image.convert("RGBA" if has_transparency else "RGB")

            output = io.BytesIO()
            image.save(output, format="PNG")
            return output.getvalue()
    except Exception:
        return None


def normalize_metafile_to_png_with_system_drawing(
    image_bytes: bytes | None,
    ext: str,
    *,
    scale: float = 4.0,
    max_pixels: int = 18000000,
    max_edge: int = 5200,
) -> bytes | None:
    if not image_bytes or os.name != "nt":
        return None
    extension = str(ext or "").strip().lower()
    if extension not in {"wmf", "emf"}:
        return None

    input_path: Path | None = None
    output_path: Path | None = None
    script_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile("wb", suffix=f".{extension}", delete=False) as handle:
            handle.write(image_bytes)
            input_path = Path(handle.name)
        with tempfile.NamedTemporaryFile("wb", suffix=".png", delete=False) as handle:
            output_path = Path(handle.name)

        script_body = r"""
param(
  [string]$InputPath,
  [string]$OutputPath,
  [double]$Scale = 4.0,
  [int]$MaxPixels = 18000000,
  [int]$MaxEdge = 5200
)

Add-Type -AssemblyName System.Drawing
$metafile = New-Object System.Drawing.Imaging.Metafile($InputPath)
try {
  $width = [Math]::Max(1, [int][Math]::Round($metafile.Width * $Scale))
  $height = [Math]::Max(1, [int][Math]::Round($metafile.Height * $Scale))
  if ($width -gt $MaxEdge -or $height -gt $MaxEdge) {
    $edgeScale = [Math]::Min(($MaxEdge / [double]$width), ($MaxEdge / [double]$height))
    $width = [Math]::Max(1, [int][Math]::Round($width * $edgeScale))
    $height = [Math]::Max(1, [int][Math]::Round($height * $edgeScale))
  }
  $pixelCount = [double]$width * [double]$height
  if ($pixelCount -gt $MaxPixels) {
    $pixelScale = [Math]::Sqrt($MaxPixels / $pixelCount)
    $width = [Math]::Max(1, [int][Math]::Round($width * $pixelScale))
    $height = [Math]::Max(1, [int][Math]::Round($height * $pixelScale))
  }
  $bitmap = New-Object System.Drawing.Bitmap($width, $height)
  try {
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
      $graphics.Clear([System.Drawing.Color]::White)
      $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
      $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
      $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
      $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
      $graphics.DrawImage($metafile, 0, 0, $width, $height)
      $bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
      $graphics.Dispose()
    }
  } finally {
    $bitmap.Dispose()
  }
} finally {
  $metafile.Dispose()
}
"""

        with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8") as handle:
            handle.write(script_body)
            script_path = Path(handle.name)

        completed = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
                "-InputPath",
                str(input_path),
                "-OutputPath",
                str(output_path),
                "-Scale",
                str(max(1.0, float(scale or 4.0))),
                "-MaxPixels",
                str(max(1000000, int(max_pixels or 18000000))),
                "-MaxEdge",
                str(max(512, int(max_edge or 5200))),
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=45,
            check=False,
        )
        if completed.returncode != 0 or output_path is None or not output_path.exists():
            return None
        png_bytes = output_path.read_bytes()
        return png_bytes or None
    except Exception:
        return None
    finally:
        for path in (input_path, output_path, script_path):
            if path is not None and path.exists():
                try:
                    path.unlink()
                except OSError:
                    pass


def build_alt_excel(rows: list[dict], source_filename: str, preview_images: dict[int, object] | None = None) -> bytes:
    headers = [
        "File Name",
        "Image",
        "ALT Text",
        "Item ID",
        "Type",
        "Source Part",
        "Page",
        "Original ALT",
        "Generated ALT",
    ]

    records = [headers]
    preview_images = preview_images or {}
    image_entries = []

    for idx, row in enumerate(rows, start=1):
        row_id = row.get("id")
        image_entry = preview_images.get(row_id) if isinstance(row_id, int) else None
        image_bytes, image_ext = parse_preview_entry(image_entry)
        has_preview = isinstance(image_bytes, (bytes, bytearray)) and len(image_bytes) > 0

        if has_preview:
            width_px, height_px = png_dimensions(bytes(image_bytes))
            preview_role = str(row.get("role", "")).lower()
            max_width = 320 if preview_role == "equation" else 220
            max_height = 160 if preview_role == "equation" else 120
            if width_px <= 0 or height_px <= 0:
                width_px, height_px = max_width, max_height
            width_px, height_px = fit_within(width_px, height_px, max_width, max_height)
            image_entries.append(
                {
                    "row_zero_index": idx,  # header is row zero index 0, first data row is 1.
                    "col_zero_index": 1,   # Column B ("Image")
                    "bytes": bytes(image_bytes),
                    "ext": image_ext,
                    "width_px": width_px,
                    "height_px": height_px,
                    "name": f"{row.get('type', 'Item')} {idx}",
                }
            )

        records.append(
            [
                source_filename,
                "",
                row.get("alt_text", ""),
                row.get("id", idx - 1),
                row.get("type", ""),
                row.get("source_part", ""),
                row.get("page", ""),
                row.get("existing_alt_text", ""),
                row.get("generated_alt_text", ""),
            ]
        )

    sheet_rows = []
    for row_index, row_values in enumerate(records, start=1):
        cells = []
        for col_index, value in enumerate(row_values, start=1):
            cell_ref = f"{excel_column_name(col_index)}{row_index}"
            cell_text = escape("" if value is None else str(value))
            cells.append(f'<c r="{cell_ref}" t="inlineStr"><is><t>{cell_text}</t></is></c>')

        row_attrs = f' r="{row_index}"'
        if row_index > 1:
            preview_role = str(rows[row_index - 2].get("role", "")).lower() if row_index - 2 < len(rows) else ""
            row_height = "118" if preview_role == "equation" else "95"
            row_attrs += f' ht="{row_height}" customHeight="1"'
        sheet_rows.append(f"<row{row_attrs}>{''.join(cells)}</row>")

    cols_xml = (
        "<cols>"
        '<col min="1" max="1" width="34" customWidth="1"/>'
        '<col min="2" max="2" width="44" customWidth="1"/>'
        '<col min="3" max="3" width="72" customWidth="1"/>'
        '<col min="4" max="4" width="12" customWidth="1" hidden="1"/>'
        '<col min="5" max="5" width="14" customWidth="1" hidden="1"/>'
        '<col min="6" max="6" width="14" customWidth="1" hidden="1"/>'
        '<col min="7" max="7" width="10" customWidth="1" hidden="1"/>'
        '<col min="8" max="8" width="72" customWidth="1" hidden="1"/>'
        '<col min="9" max="9" width="72" customWidth="1" hidden="1"/>'
        "</cols>"
    )

    drawing_tag = '<drawing r:id="rId1"/>' if image_entries else ""
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"{cols_xml}<sheetData>{''.join(sheet_rows)}</sheetData>{drawing_tag}"
        "</worksheet>"
    )

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="ALT Management" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )

    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        "</Relationships>"
    )

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        "</styleSheet>"
    )

    root_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )

    image_exts = sorted({entry.get("ext", "png") for entry in image_entries})
    image_defaults = "".join(
        f'<Default Extension="{ext}" ContentType="{content_type_for_extension(ext)}"/>'
        for ext in image_exts
    )

    content_types_xml = "".join(
        [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
            '<Default Extension="xml" ContentType="application/xml"/>',
            image_defaults,
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
            '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
            if image_entries
            else "",
            "</Types>",
        ]
    )

    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types_xml)
        archive.writestr("_rels/.rels", root_rels_xml)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        archive.writestr("xl/styles.xml", styles_xml)

        if image_entries:
            archive.writestr("xl/worksheets/_rels/sheet1.xml.rels", build_sheet_relationships_xml())
            archive.writestr("xl/drawings/drawing1.xml", build_drawing_xml(image_entries))
            archive.writestr("xl/drawings/_rels/drawing1.xml.rels", build_drawing_relationships_xml(image_entries))
            for image_index, entry in enumerate(image_entries, start=1):
                ext = entry.get("ext", "png")
                archive.writestr(f"xl/media/image{image_index}.{ext}", entry["bytes"])

    return output.getvalue()


def excel_column_name(index: int) -> str:
    name = ""
    current = index
    while current > 0:
        current, remainder = divmod(current - 1, 26)
        name = chr(65 + remainder) + name
    return name


def build_sheet_relationships_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
        "</Relationships>"
    )


def build_drawing_relationships_xml(image_entries: list[dict]) -> str:
    relationships = []
    for image_index, entry in enumerate(image_entries, start=1):
        ext = entry.get("ext", "png")
        relationships.append(
            f'<Relationship Id="rId{image_index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image{image_index}.{ext}"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{''.join(relationships)}"
        "</Relationships>"
    )


def build_drawing_xml(image_entries: list[dict]) -> str:
    anchors = []
    for image_index, entry in enumerate(image_entries, start=1):
        cx = int(entry["width_px"]) * 9525
        cy = int(entry["height_px"]) * 9525
        row_zero_index = int(entry["row_zero_index"])
        col_zero_index = int(entry.get("col_zero_index", 10))
        name = escape(str(entry.get("name", f"Preview {image_index}")))
        anchors.append(
            '<xdr:oneCellAnchor>'
            f"<xdr:from><xdr:col>{col_zero_index}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row_zero_index}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
            f'<xdr:ext cx="{cx}" cy="{cy}"/>'
            "<xdr:pic>"
            f'<xdr:nvPicPr><xdr:cNvPr id="{image_index}" name="{name}"/><xdr:cNvPicPr/></xdr:nvPicPr>'
            f'<xdr:blipFill><a:blip r:embed="rId{image_index}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
            f'<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
            "</xdr:pic>"
            "<xdr:clientData/>"
            "</xdr:oneCellAnchor>"
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"{''.join(anchors)}"
        "</xdr:wsDr>"
    )


def png_dimensions(image_bytes: bytes) -> tuple[int, int]:
    if image_bytes[:8] != b"\x89PNG\r\n\x1a\n":
        return (0, 0)
    if len(image_bytes) < 24:
        return (0, 0)
    width = int.from_bytes(image_bytes[16:20], "big")
    height = int.from_bytes(image_bytes[20:24], "big")
    return (width, height)


def fit_within(width_px: int, height_px: int, max_width: int, max_height: int) -> tuple[int, int]:
    if width_px <= 0 or height_px <= 0:
        return (max_width, max_height)
    scale = min(max_width / width_px, max_height / height_px, 1.0)
    return (max(1, int(width_px * scale)), max(1, int(height_px * scale)))


def parse_preview_entry(entry: object) -> tuple[bytes | None, str]:
    if isinstance(entry, dict):
        image_bytes = entry.get("bytes")
        image_ext = str(entry.get("ext", "png")).lower().lstrip(".")
        if isinstance(image_bytes, (bytes, bytearray)) and image_ext:
            return (bytes(image_bytes), image_ext)
        return (None, "png")

    if isinstance(entry, (bytes, bytearray)):
        return (bytes(entry), "png")

    return (None, "png")


def render_text_tile(title: str, body: str, width: int = 720, height: int = 180) -> bytes:
    document = fitz.open()
    page = document.new_page(width=width, height=height)
    page.draw_rect(fitz.Rect(0, 0, width, height), color=(0.82, 0.78, 0.72), fill=(0.99, 0.98, 0.96), width=1)
    page.insert_textbox(
        fitz.Rect(24, 20, width - 24, 48),
        normalize_alt_text(title or "Preview"),
        fontname="helv",
        fontsize=15,
        color=(0.24, 0.18, 0.11),
    )
    page.insert_textbox(
        fitz.Rect(24, 58, width - 24, height - 20),
        normalize_alt_text(body or "Preview unavailable"),
        fontname="Times-Roman",
        fontsize=16,
        color=(0.16, 0.16, 0.16),
    )
    pixmap = page.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), alpha=False)
    return pixmap.tobytes("png")


def content_type_for_extension(ext: str) -> str:
    ext = ext.lower().lstrip(".")
    mapping = {
        "png": "image/png",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "gif": "image/gif",
        "bmp": "image/bmp",
        "wmf": "image/x-wmf",
        "emf": "image/x-emf",
    }
    return mapping.get(ext, "application/octet-stream")
