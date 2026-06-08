from __future__ import annotations

import math
import re
import uuid
from pathlib import Path

import fitz

from Altomizer.alt_management import (
    build_alt_preview_images,
    normalize_alt_text,
    summarize_alt_rows,
)

if hasattr(fitz, "TOOLS"):
    fitz.TOOLS.mupdf_display_errors(False)
    fitz.TOOLS.mupdf_display_warnings(False)


PDF_REF_RE = re.compile(r"(\d+)\s+\d+\s+R")
PDF_NUMBER_RE = re.compile(r"[-+]?(?:\d+\.\d+|\d+|\.\d+)")
PDF_BBOX_RE = re.compile(r"/BBox\s+(?P<value>\[[^\]]+\]|\d+\s+\d+\s+R)")
TARGET_STANDARD_ROLES = {"figure", "formula"}
ALT_NUMBER_WORDS = {
    "zero": "0",
    "one": "1",
    "two": "2",
    "three": "3",
    "four": "4",
    "five": "5",
    "six": "6",
    "seven": "7",
    "eight": "8",
    "nine": "9",
    "ten": "10",
    "eleven": "11",
    "twelve": "12",
    "thirteen": "13",
    "fourteen": "14",
    "fifteen": "15",
    "sixteen": "16",
    "seventeen": "17",
    "eighteen": "18",
    "nineteen": "19",
    "twenty": "20",
    "half": "2",
    "halves": "2",
    "third": "3",
    "thirds": "3",
    "quarter": "4",
    "quarters": "4",
    "fourth": "4",
    "fourths": "4",
}
FORMULA_RUN_BREAK_WORDS = {
    "and",
    "or",
    "but",
    "given",
    "where",
    "then",
    "therefore",
    "because",
    "determine",
    "solution",
    "equation",
    "formula",
    "definition",
    "definitions",
    "chapter",
    "section",
    "contents",
    "form",
    "point",
}
SHORT_FORMULA_PROSE_WORDS = {
    "and",
    "as",
    "by",
    "for",
    "from",
    "if",
    "in",
    "is",
    "of",
    "on",
    "or",
    "the",
    "then",
    "to",
    "with",
    "form",
    "formula",
    "definition",
    "definitions",
}
SHORT_FORMULA_FUNCTION_WORDS = {"sin", "cos", "tan", "sec", "csc", "cot", "log", "ln", "exp", "lim", "gcd", "lcm", "max", "min"}
PAINTING_OPERATORS = {
    "Tj",
    "TJ",
    "'",
    '"',
    "Do",
    "S",
    "s",
    "f",
    "F",
    "f*",
    "B",
    "B*",
    "b",
    "b*",
    "sh",
}
VECTOR_PAINTING_OPERATORS = {
    "S",
    "s",
    "f",
    "F",
    "f*",
    "B",
    "B*",
    "b",
    "b*",
    "sh",
}
STRUCTURE_OPERATORS = {"BDC", "BMC", "EMC"}
TEXT_OBJECT_OPERATORS = {"BT", "ET"}
GRAPHICS_STATE_OPERATORS = {"q", "Q"}
NON_PAINTING_OPERATORS = {
    "cm",
    "m",
    "l",
    "c",
    "v",
    "y",
    "h",
    "re",
    "W",
    "W*",
    "n",
    "CS",
    "cs",
    "SC",
    "SCN",
    "sc",
    "scn",
    "G",
    "g",
    "RG",
    "rg",
    "K",
    "k",
    "w",
    "J",
    "j",
    "M",
    "d",
    "ri",
    "gs",
    "Tf",
    "Tm",
    "Td",
    "TD",
    "T*",
    "Tc",
    "Tw",
    "Tz",
    "TL",
    "Tr",
    "Ts",
}
PDF_OPERATORS = PAINTING_OPERATORS | STRUCTURE_OPERATORS | TEXT_OBJECT_OPERATORS | GRAPHICS_STATE_OPERATORS | NON_PAINTING_OPERATORS


def pdf_name_key(value: object) -> str:
    return str(value or "").strip().lstrip("/").lower()


def parse_pdf_ref(value: object) -> int | None:
    match = PDF_REF_RE.search(str(value or ""))
    if not match:
        return None
    try:
        return int(match.group(1))
    except ValueError:
        return None


def parse_pdf_refs(value: object) -> list[int]:
    refs: list[int] = []
    for match in PDF_REF_RE.finditer(str(value or "")):
        try:
            refs.append(int(match.group(1)))
        except ValueError:
            continue
    return refs


def parse_number_array(value: object) -> list[float] | None:
    numbers = []
    for raw_number in PDF_NUMBER_RE.findall(str(value or "")):
        try:
            numbers.append(float(raw_number))
        except ValueError:
            continue
    if len(numbers) < 4:
        return None
    return numbers[:4]


def pdf_content_tokens(content: bytes) -> list[str]:
    tokens: list[str] = []
    index = 0
    length = len(content)
    whitespace = b"\x00\t\n\f\r "
    delimiters = b"()<>[]{}/%"

    while index < length:
        char = content[index]
        if char in whitespace:
            index += 1
            continue
        if char == ord("%"):
            while index < length and content[index] not in b"\r\n":
                index += 1
            continue
        if char == ord("("):
            depth = 1
            index += 1
            while index < length and depth > 0:
                current = content[index]
                if current == ord("\\"):
                    index += 2
                    continue
                if current == ord("("):
                    depth += 1
                elif current == ord(")"):
                    depth -= 1
                index += 1
            tokens.append("(...)")
            continue
        if char == ord("<"):
            if index + 1 < length and content[index + 1] == ord("<"):
                tokens.append("<<")
                index += 2
                continue
            index += 1
            while index < length:
                current = content[index]
                if current == ord(">"):
                    index += 1
                    break
                index += 1
            tokens.append("<...>")
            continue
        if char == ord(">") and index + 1 < length and content[index + 1] == ord(">"):
            tokens.append(">>")
            index += 2
            continue
        if char in b"[]{}":
            tokens.append(chr(char))
            index += 1
            continue
        if char == ord("/"):
            start = index
            index += 1
            while index < length and content[index] not in whitespace and content[index] not in delimiters:
                index += 1
            tokens.append(content[start:index].decode("latin1", errors="ignore"))
            continue

        start = index
        while index < length and content[index] not in whitespace and content[index] not in delimiters:
            index += 1
        tokens.append(content[start:index].decode("latin1", errors="ignore"))

    return tokens


def marked_content_mcid(operands: list[str]) -> int | None:
    for index, token in enumerate(operands):
        if token != "/MCID" or index + 1 >= len(operands):
            continue
        try:
            return int(float(operands[index + 1]))
        except ValueError:
            continue
    return None


def valid_content_bbox(bbox: object) -> bool:
    if not isinstance(bbox, (tuple, list)) or len(bbox) < 4:
        return False
    try:
        x0, y0, x1, y1 = (float(value) for value in bbox[:4])
    except (TypeError, ValueError):
        return False
    return x1 > x0 and y1 > y0


def bbox_area(bbox: object) -> float:
    if not valid_content_bbox(bbox):
        return 0.0
    x0, y0, x1, y1 = (float(value) for value in bbox[:4])
    return max(0.0, x1 - x0) * max(0.0, y1 - y0)


def fitz_rect_tuple(rect: fitz.Rect) -> tuple[float, float, float, float]:
    return (float(rect.x0), float(rect.y0), float(rect.x1), float(rect.y1))


def expand_fitz_bbox(
    bbox: list[float],
    page_rect: fitz.Rect,
    *,
    x_pad: float,
    y_pad: float,
) -> fitz.Rect:
    return fitz.Rect(
        max(page_rect.x0, float(bbox[0]) - x_pad),
        max(page_rect.y0, float(bbox[1]) - y_pad),
        min(page_rect.x1, float(bbox[2]) + x_pad),
        min(page_rect.y1, float(bbox[3]) + y_pad),
    )


def union_content_bboxes(bboxes: list[tuple[float, float, float, float]]) -> list[float] | None:
    valid_bboxes = [bbox for bbox in bboxes if valid_content_bbox(bbox)]
    if not valid_bboxes:
        return None
    return [
        min(bbox[0] for bbox in valid_bboxes),
        min(bbox[1] for bbox in valid_bboxes),
        max(bbox[2] for bbox in valid_bboxes),
        max(bbox[3] for bbox in valid_bboxes),
    ]


def bbox_gap(left: list[float] | tuple[float, float, float, float], right: list[float] | tuple[float, float, float, float]) -> tuple[float, float]:
    left_x0, left_y0, left_x1, left_y1 = (float(value) for value in left[:4])
    right_x0, right_y0, right_x1, right_y1 = (float(value) for value in right[:4])
    gap_x = max(0.0, max(right_x0 - left_x1, left_x0 - right_x1))
    gap_y = max(0.0, max(right_y0 - left_y1, left_y0 - right_y1))
    return gap_x, gap_y


def bbox_intersection_area(left: list[float] | tuple[float, float, float, float], right: list[float] | tuple[float, float, float, float]) -> float:
    left_x0, left_y0, left_x1, left_y1 = (float(value) for value in left[:4])
    right_x0, right_y0, right_x1, right_y1 = (float(value) for value in right[:4])
    overlap_x = max(0.0, min(left_x1, right_x1) - max(left_x0, right_x0))
    overlap_y = max(0.0, min(left_y1, right_y1) - max(left_y0, right_y0))
    return overlap_x * overlap_y


def bbox_center(left: list[float] | tuple[float, float, float, float]) -> tuple[float, float]:
    x0, y0, x1, y1 = (float(value) for value in left[:4])
    return ((x0 + x1) / 2.0, (y0 + y1) / 2.0)


def dedupe_bboxes(bboxes: list[tuple[float, float, float, float]]) -> list[tuple[float, float, float, float]]:
    deduped: list[tuple[float, float, float, float]] = []
    seen: set[tuple[float, float, float, float]] = set()
    for bbox in bboxes:
        rounded = tuple(round(float(value), 3) for value in bbox[:4])
        if rounded in seen:
            continue
        seen.add(rounded)
        deduped.append(tuple(float(value) for value in bbox[:4]))
    return deduped


def median_number(values: list[float]) -> float:
    if not values:
        return 0.0
    ordered = sorted(float(value) for value in values)
    middle = len(ordered) // 2
    if len(ordered) % 2:
        return ordered[middle]
    return (ordered[middle - 1] + ordered[middle]) / 2.0


def numeric_operands(operands: list[str]) -> list[float]:
    numbers: list[float] = []
    for token in operands:
        try:
            numbers.append(float(token))
        except ValueError:
            continue
    return numbers


def active_mcid_set(mcid_stack: list[int | None]) -> set[int]:
    return {mcid for mcid in mcid_stack if isinstance(mcid, int)}


def pdf_point_to_fitz(page: fitz.Page, x: float, y: float) -> fitz.Point:
    return fitz.Point(float(x), float(y)) * page.transformation_matrix


def multiply_pdf_matrices(
    left: tuple[float, float, float, float, float, float],
    right: tuple[float, float, float, float, float, float],
) -> tuple[float, float, float, float, float, float]:
    a1, b1, c1, d1, e1, f1 = left
    a2, b2, c2, d2, e2, f2 = right
    return (
        a1 * a2 + c1 * b2,
        b1 * a2 + d1 * b2,
        a1 * c2 + c1 * d2,
        b1 * c2 + d1 * d2,
        a1 * e2 + c1 * f2 + e1,
        b1 * e2 + d1 * f2 + f1,
    )


def pdf_translate_matrix(tx: float, ty: float) -> tuple[float, float, float, float, float, float]:
    return (1.0, 0.0, 0.0, 1.0, float(tx), float(ty))


def transform_pdf_point(
    matrix: tuple[float, float, float, float, float, float],
    x: float,
    y: float,
) -> tuple[float, float]:
    a, b, c, d, e, f = matrix
    return (a * x + c * y + e, b * x + d * y + f)


def image_bbox_from_pdf_matrix(
    page: fitz.Page,
    matrix: tuple[float, float, float, float, float, float],
) -> tuple[float, float, float, float] | None:
    points = [
        pdf_point_to_fitz(page, *transform_pdf_point(matrix, 0.0, 0.0)),
        pdf_point_to_fitz(page, *transform_pdf_point(matrix, 1.0, 0.0)),
        pdf_point_to_fitz(page, *transform_pdf_point(matrix, 0.0, 1.0)),
        pdf_point_to_fitz(page, *transform_pdf_point(matrix, 1.0, 1.0)),
    ]
    bbox = (
        min(point.x for point in points),
        min(point.y for point in points),
        max(point.x for point in points),
        max(point.y for point in points),
    )
    if not valid_content_bbox(bbox):
        return None
    return bbox


def trace_bboxes_by_seqno(page: fitz.Page) -> tuple[dict[int, list[tuple[float, float, float, float]]], list[dict]]:
    bboxes: dict[int, list[tuple[float, float, float, float]]] = {}
    traces = page.get_texttrace()
    for trace in traces:
        seqno = trace.get("seqno")
        bbox = trace.get("bbox")
        if isinstance(seqno, int) and valid_content_bbox(bbox):
            bboxes.setdefault(seqno, []).append(tuple(float(value) for value in bbox[:4]))
    return bboxes, traces


def nearest_trace_seqno_for_origin(
    traces: list[dict],
    origin: fitz.Point,
    *,
    tolerance: float = 2.5,
) -> int | None:
    best_seqno = None
    best_distance = float("inf")
    for trace in traces:
        seqno = trace.get("seqno")
        if not isinstance(seqno, int):
            continue
        for char in trace.get("chars") or []:
            if len(char) < 3:
                continue
            char_origin = char[2]
            if not isinstance(char_origin, (tuple, list)) or len(char_origin) < 2:
                continue
            try:
                dx = abs(float(char_origin[0]) - origin.x)
                dy = abs(float(char_origin[1]) - origin.y)
            except (TypeError, ValueError):
                continue
            distance = dx + dy
            if dx <= tolerance and dy <= tolerance and distance < best_distance:
                best_distance = distance
                best_seqno = seqno
    return best_seqno


def page_xref_map(document: fitz.Document) -> dict[int, int]:
    return {document.page_xref(page_index): page_index for page_index in range(document.page_count)}


def get_struct_tree_root_xref(document: fitz.Document) -> int | None:
    try:
        catalog_xref = document.pdf_catalog()
        value_type, value = document.xref_get_key(catalog_xref, "StructTreeRoot")
    except Exception:
        return None
    if value_type != "xref":
        return None
    return parse_pdf_ref(value)


def read_role_map(document: fitz.Document, struct_root_xref: int) -> dict[str, str]:
    try:
        value_type, value = document.xref_get_key(struct_root_xref, "RoleMap")
    except Exception:
        return {}
    if value_type != "xref":
        return {}
    role_map_xref = parse_pdf_ref(value)
    if not isinstance(role_map_xref, int):
        return {}

    role_map: dict[str, str] = {}
    try:
        for source_role in document.xref_get_keys(role_map_xref):
            mapped_type, mapped_value = document.xref_get_key(role_map_xref, source_role)
            if mapped_type == "name":
                role_map[pdf_name_key(source_role)] = pdf_name_key(mapped_value)
    except Exception:
        return role_map
    return role_map


def resolve_pdf_role(role: str, role_map: dict[str, str]) -> str:
    current = pdf_name_key(role)
    seen: set[str] = set()
    while current in role_map and current not in seen:
        seen.add(current)
        current = pdf_name_key(role_map[current])
    return current


def child_struct_refs(document: fitz.Document, xref: int) -> list[int]:
    refs: list[int] = []
    try:
        _value_type, value = document.xref_get_key(xref, "K")
    except Exception:
        return refs
    for ref in parse_pdf_refs(value):
        if ref != xref:
            refs.append(ref)
    return refs


def structure_standard_role(document: fitz.Document, struct_xref: int, role_map: dict[str, str]) -> str | None:
    try:
        role_type, role_value = document.xref_get_key(struct_xref, "S")
    except Exception:
        return None
    if role_type != "name":
        return None
    return resolve_pdf_role(pdf_name_key(role_value), role_map)


def structure_has_alt_field(document: fitz.Document, struct_xref: int) -> bool:
    try:
        value_type, value = document.xref_get_key(struct_xref, "Alt")
    except Exception:
        return False
    if value_type not in {"string", "name"}:
        return False
    return bool(normalize_alt_text(str(value or "").replace("\x00", "")))


def has_descendant_with_standard_role(
    document: fitz.Document,
    struct_xref: int,
    role_map: dict[str, str],
    standard_role: str,
) -> bool:
    seen: set[int] = set()
    stack = child_struct_refs(document, struct_xref)

    while stack:
        child_xref = stack.pop(0)
        if child_xref in seen:
            continue
        seen.add(child_xref)
        if structure_standard_role(document, child_xref, role_map) == standard_role:
            return True
        stack.extend(child_struct_refs(document, child_xref))
    return False


def structure_mcids(document: fitz.Document, struct_xref: int) -> set[int]:
    mcids: set[int] = set()
    try:
        value_type, value = document.xref_get_key(struct_xref, "K")
    except Exception:
        return mcids

    if value_type == "int":
        try:
            mcids.add(int(value))
        except ValueError:
            pass
    elif value_type == "array":
        raw_value = str(value or "")
        raw_value = re.sub(r"\b\d+\s+\d+\s+R\b", " ", raw_value)
        for raw_number in re.findall(r"(?<![\w/.-])-?\d+(?![\w.])", raw_value):
            try:
                mcids.add(int(raw_number))
            except ValueError:
                continue

    for raw_number in re.findall(r"/MCID\s+(-?\d+)", str(value or "")):
        try:
            mcids.add(int(raw_number))
        except ValueError:
            continue
    return mcids


def owned_structure_mcids(
    document: fitz.Document,
    struct_xref: int,
    role_map: dict[str, str],
) -> set[int]:
    seen: set[int] = set()

    def collect(current_xref: int, *, is_root: bool = False) -> set[int]:
        if current_xref in seen:
            return set()
        seen.add(current_xref)

        collected = set(structure_mcids(document, current_xref))
        for child_xref in child_struct_refs(document, current_xref):
            child_role = structure_standard_role(document, child_xref, role_map)
            if child_role in TARGET_STANDARD_ROLES:
                continue
            if not is_root and structure_has_alt_field(document, child_xref):
                continue
            collected.update(collect(child_xref))
        return collected

    return collect(struct_xref, is_root=True)


def reachable_struct_elements(document: fitz.Document, struct_root_xref: int) -> list[int]:
    ordered: list[int] = []
    seen: set[int] = set()
    stack = child_struct_refs(document, struct_root_xref)

    while stack:
        xref = stack.pop(0)
        if xref in seen:
            continue
        seen.add(xref)
        try:
            type_value = document.xref_get_key(xref, "Type")
        except Exception:
            type_value = ("null", "null")
        if type_value != ("name", "/StructElem"):
            try:
                role_value = document.xref_get_key(xref, "S")
            except Exception:
                continue
            if role_value[0] != "name":
                continue
        ordered.append(xref)
        stack.extend(child_struct_refs(document, xref))
    return ordered


def page_marked_content_components(
    document: fitz.Document,
    page_index: int,
    *,
    include_paint: bool = False,
) -> dict[int, dict[str, list[tuple[float, float, float, float]]]]:
    if page_index < 0 or page_index >= document.page_count:
        return {}

    page = document.load_page(page_index)
    mcid_stack: list[int | None] = []
    identity_matrix = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
    current_matrix = identity_matrix
    graphics_stack: list[tuple[float, float, float, float, float, float]] = []
    text_matrix = identity_matrix
    text_line_matrix = identity_matrix
    text_leading = 0.0
    in_text_object = False
    try:
        bboxlog = page.get_bboxlog()
    except Exception:
        bboxlog = []
    path_bboxlog = [
        entry
        for entry in bboxlog
        if isinstance(entry, tuple)
        and len(entry) > 1
        and isinstance(entry[0], str)
        and "path" in entry[0]
    ]
    path_bbox_index = 0
    text_origins_by_mcid: dict[int, list[fitz.Point]] = {}
    image_bboxes_by_mcid: dict[int, list[tuple[float, float, float, float]]] = {}
    paint_bboxes_by_mcid: dict[int, list[tuple[float, float, float, float]]] = {}
    image_name_to_xref = {
        f"/{image[7]}": int(image[0])
        for image in page.get_images(full=True)
        if len(image) > 7 and image[7] and int(image[0]) > 0
    }

    for content_xref in page.get_contents() or []:
        try:
            content = document.xref_stream(content_xref)
        except Exception:
            continue
        operands: list[str] = []
        for token in pdf_content_tokens(content):
            if token not in PDF_OPERATORS:
                operands.append(token)
                continue

            if token == "BDC":
                mcid_stack.append(marked_content_mcid(operands))
                operands = []
                continue
            if token == "BMC":
                mcid_stack.append(None)
                operands = []
                continue
            if token == "EMC":
                if mcid_stack:
                    mcid_stack.pop()
                operands = []
                continue
            if token == "q":
                graphics_stack.append(current_matrix)
                operands = []
                continue
            if token == "Q":
                current_matrix = graphics_stack.pop() if graphics_stack else identity_matrix
                operands = []
                continue
            if token == "cm":
                numbers = numeric_operands(operands)
                if len(numbers) >= 6:
                    matrix = tuple(float(value) for value in numbers[-6:])
                    current_matrix = multiply_pdf_matrices(current_matrix, matrix)
                operands = []
                continue
            if token == "BT":
                in_text_object = True
                text_matrix = identity_matrix
                text_line_matrix = identity_matrix
                operands = []
                continue
            if token == "ET":
                in_text_object = False
                text_matrix = identity_matrix
                text_line_matrix = identity_matrix
                operands = []
                continue
            if token == "Tm":
                numbers = numeric_operands(operands)
                if len(numbers) >= 6:
                    text_matrix = tuple(float(value) for value in numbers[-6:])
                    text_line_matrix = text_matrix
                operands = []
                continue
            if token == "Td" and in_text_object:
                numbers = numeric_operands(operands)
                if len(numbers) >= 2:
                    translation = pdf_translate_matrix(numbers[-2], numbers[-1])
                    text_line_matrix = multiply_pdf_matrices(text_line_matrix, translation)
                    text_matrix = text_line_matrix
                operands = []
                continue
            if token == "TD" and in_text_object:
                numbers = numeric_operands(operands)
                if len(numbers) >= 2:
                    text_leading = -float(numbers[-1])
                    translation = pdf_translate_matrix(numbers[-2], numbers[-1])
                    text_line_matrix = multiply_pdf_matrices(text_line_matrix, translation)
                    text_matrix = text_line_matrix
                operands = []
                continue
            if token == "T*" and in_text_object:
                translation = pdf_translate_matrix(0.0, -text_leading)
                text_line_matrix = multiply_pdf_matrices(text_line_matrix, translation)
                text_matrix = text_line_matrix
                operands = []
                continue

            active_mcids = active_mcid_set(mcid_stack)
            if token in VECTOR_PAINTING_OPERATORS:
                if active_mcids and path_bbox_index < len(path_bboxlog):
                    paint_entry = path_bboxlog[path_bbox_index]
                    paint_bbox = paint_entry[1] if len(paint_entry) > 1 else None
                    if include_paint and valid_content_bbox(paint_bbox):
                        bbox = tuple(float(value) for value in paint_bbox[:4])
                        for mcid in active_mcids:
                            paint_bboxes_by_mcid.setdefault(mcid, []).append(bbox)
                path_bbox_index += 1

            if active_mcids and token in {"Tj", "TJ", "'", '"'} and in_text_object:
                origin_x, origin_y = transform_pdf_point(current_matrix, text_matrix[4], text_matrix[5])
                origin = pdf_point_to_fitz(page, origin_x, origin_y)
                for mcid in active_mcids:
                    text_origins_by_mcid.setdefault(mcid, []).append(origin)
            elif active_mcids and token == "Do" and operands:
                image_xref = image_name_to_xref.get(operands[-1])
                if isinstance(image_xref, int):
                    bbox = image_bbox_from_pdf_matrix(page, current_matrix)
                    if bbox is not None:
                        if valid_content_bbox(bbox):
                            for mcid in active_mcids:
                                image_bboxes_by_mcid.setdefault(mcid, []).append(bbox)
                    else:
                        for rect in page.get_image_rects(image_xref):
                            bbox = (float(rect.x0), float(rect.y0), float(rect.x1), float(rect.y1))
                            if valid_content_bbox(bbox):
                                for mcid in active_mcids:
                                    image_bboxes_by_mcid.setdefault(mcid, []).append(bbox)
            operands = []

    seqno_bboxes, traces = trace_bboxes_by_seqno(page)
    components_by_mcid: dict[int, dict[str, list[tuple[float, float, float, float]]]] = {}
    for mcid, values in image_bboxes_by_mcid.items():
        components_by_mcid.setdefault(mcid, {"image": [], "paint": [], "text": []})["image"].extend(values)
    for mcid, values in paint_bboxes_by_mcid.items():
        components_by_mcid.setdefault(mcid, {"image": [], "paint": [], "text": []})["paint"].extend(values)
    for mcid, origins in text_origins_by_mcid.items():
        matched_seqnos = {
            seqno
            for origin in origins
            if (seqno := nearest_trace_seqno_for_origin(traces, origin)) is not None
        }
        for seqno in matched_seqnos:
            components_by_mcid.setdefault(mcid, {"image": [], "paint": [], "text": []})["text"].extend(seqno_bboxes.get(seqno, []))

    filtered: dict[int, dict[str, list[tuple[float, float, float, float]]]] = {}
    for mcid, groups in components_by_mcid.items():
        image_values = dedupe_bboxes([bbox for bbox in groups.get("image", []) if valid_content_bbox(bbox)])
        paint_values = dedupe_bboxes([bbox for bbox in groups.get("paint", []) if valid_content_bbox(bbox)])
        text_values = dedupe_bboxes([bbox for bbox in groups.get("text", []) if valid_content_bbox(bbox)])
        if image_values or paint_values or text_values:
            filtered[mcid] = {
                "image": image_values,
                "paint": paint_values,
                "text": text_values,
            }
    return filtered


def page_marked_content_bboxes(document: fitz.Document, page_index: int, *, include_paint: bool = False) -> dict[int, list[float]]:
    components_by_mcid = page_marked_content_components(document, page_index, include_paint=include_paint)
    bboxes_by_mcid: dict[int, list[float]] = {}
    for mcid, groups in components_by_mcid.items():
        values = groups.get("image", []) + groups.get("paint", []) + groups.get("text", [])
        bbox = union_content_bboxes(values)
        if bbox is not None:
            bboxes_by_mcid[mcid] = bbox
    return bboxes_by_mcid


def formula_alt_tokens(alt_text: str) -> set[str]:
    normalized = normalize_alt_text(str(alt_text or "")).lower()
    tokens = set(re.findall(r"\d+(?:\.\d+)?|[a-z]+", normalized))
    for token in list(tokens):
        if token in ALT_NUMBER_WORDS:
            tokens.add(ALT_NUMBER_WORDS[token])
    for variable, subscript in re.findall(r"\b([a-z])\s+(?:subscript|sub)\s+(\d+(?:\.\d+)?)\b", normalized):
        tokens.add(f"{variable}{subscript}")
    useful = {
        token
        for token in tokens
        if token not in {
            "open",
            "close",
            "parenthesis",
            "fraction",
            "over",
            "equals",
            "equal",
            "times",
            "plus",
            "minus",
            "sub",
            "subscript",
            "bar",
            "and",
            "bracket",
            "ellipsis",
            "divided",
            "by",
            "where",
            "showing",
            "cross",
            "cancellation",
            "the",
            "is",
            "replaced",
            "with",
            "to",
            "of",
            "or",
            "raised",
            "power",
            "formula",
            "definition",
            "definitions",
            "form",
            "negative",
        }
    }
    return useful


def visual_formula_tokens(text: str) -> set[str]:
    normalized = normalize_alt_text(str(text or "")).lower()
    tokens = set(re.findall(r"\d+(?:\.\d+)?|[a-z]+\d*|[a-z]+", normalized))
    expanded = set(tokens)
    if "{" in normalized or "}" in normalized or "brace" in normalized:
        expanded.add("brace")
    if "[" in normalized or "]" in normalized or "bracket" in normalized:
        expanded.add("bracket")
    if "(" in normalized or ")" in normalized or "parenthesis" in normalized or "paren" in normalized:
        expanded.add("parenthesis")
    if "√" in normalized or "sqrt" in normalized:
        expanded.update({"square", "root"})
    for token in list(tokens):
        match = re.match(r"([a-z]+)(\d+)$", token)
        if match:
            expanded.add(match.group(1))
            expanded.add(match.group(2))
    if "square root" in normalized or "root" in normalized:
        expanded.add("root")
        if "square" in normalized or "square root" in normalized:
            expanded.add("square")
    return expanded


def compact_formula_alpha_tokens(text: str, alt_tokens: set[str]) -> set[str]:
    normalized = normalize_alt_text(str(text or "")).lower()
    if not normalized or not alt_tokens:
        return set()
    compact_matches = re.findall(r"[a-z]{2,4}", normalized)
    expanded: set[str] = set()
    for token in compact_matches:
        if token in FORMULA_RUN_BREAK_WORDS or token in SHORT_FORMULA_PROSE_WORDS:
            continue
        matched_chars = {character for character in token if character in alt_tokens}
        if len(matched_chars) >= 2:
            expanded.update(matched_chars)
    return expanded


def is_strong_formula_token(token: str) -> bool:
    token = str(token or "").lower()
    return bool(re.search(r"\d+\.\d+", token) or re.match(r"[a-z]+\d+$", token) or len(token) >= 2 and token.isdigit())


def is_math_word(text: str) -> bool:
    cleaned = normalize_alt_text(str(text or ""))
    if not cleaned:
        return False
    if cleaned.lower() in SHORT_FORMULA_PROSE_WORDS:
        return False
    if re.search(r"[=+*/⁄_(){}\[\]|∣·×÷−-]", cleaned):
        return True
    if re.search(r"\d", cleaned) and len(cleaned) <= 12:
        return True
    if re.match(r"^[A-Za-z]\([^)]{1,20}\)$", cleaned):
        return True
    if re.match(r"^[A-Za-z]{1,3}\d*$", cleaned):
        return True
    return False


def is_long_prose_word(text: str) -> bool:
    cleaned = re.sub(r"[^A-Za-z]", "", str(text or ""))
    return len(cleaned) >= 3 and not re.match(r"^[Pp]\w*$", cleaned)


def pdf_word_text(word: tuple) -> str:
    if len(word) < 5:
        return ""
    return str(word[4] or "")


def formula_word_bbox(word: tuple) -> tuple[float, float, float, float]:
    return (float(word[0]), float(word[1]), float(word[2]), float(word[3]))


def formula_word_overlap_tokens(text: str, alt_tokens: set[str]) -> set[str]:
    if not alt_tokens:
        return set()
    overlap = visual_formula_tokens(text) & alt_tokens
    if overlap:
        return overlap | compact_formula_alpha_tokens(text, alt_tokens)
    return compact_formula_alpha_tokens(text, alt_tokens)


def formula_word_is_break(text: str, alt_tokens: set[str]) -> bool:
    cleaned = normalize_alt_text(str(text or "")).lower()
    if not cleaned or formula_word_overlap_tokens(cleaned, alt_tokens):
        return False
    tokens = re.findall(r"[a-z]+", cleaned)
    return any(token in FORMULA_RUN_BREAK_WORDS for token in tokens)


def weak_formula_edge_word(text: str, alt_tokens: set[str], segment_overlap: int) -> bool:
    if segment_overlap < 1:
        return False
    cleaned = normalize_alt_text(str(text or "")).strip().lower()
    if not cleaned or formula_word_overlap_tokens(cleaned, alt_tokens):
        return False
    if re.fullmatch(r"[=+\-*/_(){}\[\]|:.]+", cleaned):
        return False
    alpha = re.sub(r"[^a-z]", "", cleaned)
    if alpha and len(alpha) <= 3 and alpha not in SHORT_FORMULA_FUNCTION_WORDS:
        return True
    return is_long_prose_word(cleaned)


def trim_formula_segment_edges(segment: list[tuple], alt_tokens: set[str]) -> list[tuple]:
    if not segment:
        return []
    trimmed = list(segment)
    overlap = len(visual_formula_tokens(" ".join(pdf_word_text(word) for word in trimmed)) & alt_tokens) if alt_tokens else 0
    while len(trimmed) > 1 and weak_formula_edge_word(pdf_word_text(trimmed[0]), alt_tokens, overlap):
        trimmed.pop(0)
        overlap = len(visual_formula_tokens(" ".join(pdf_word_text(word) for word in trimmed)) & alt_tokens) if alt_tokens else 0
    while len(trimmed) > 1 and weak_formula_edge_word(pdf_word_text(trimmed[-1]), alt_tokens, overlap):
        trimmed.pop()
        overlap = len(visual_formula_tokens(" ".join(pdf_word_text(word) for word in trimmed)) & alt_tokens) if alt_tokens else 0
    return trimmed


def split_formula_word_segments(words: list[tuple], alt_tokens: set[str]) -> list[list[tuple]]:
    segments: list[list[tuple]] = []
    current: list[tuple] = []
    for word in words:
        text = pdf_word_text(word)
        overlap = formula_word_overlap_tokens(text, alt_tokens)
        is_formulaish = is_math_word(text) or bool(overlap)
        if formula_word_is_break(text, alt_tokens) or not is_formulaish:
            if current:
                trimmed = trim_formula_segment_edges(current, alt_tokens)
                if trimmed:
                    segments.append(trimmed)
                current = []
            continue
        current.append(word)

    if current:
        trimmed = trim_formula_segment_edges(current, alt_tokens)
        if trimmed:
            segments.append(trimmed)
    return segments


def formula_segment_stats(segment: list[tuple], alt_tokens: set[str]) -> dict:
    text = " ".join(pdf_word_text(word) for word in segment)
    tokens = visual_formula_tokens(text)
    overlap_tokens = tokens & alt_tokens if alt_tokens else set()
    overlap_tokens.update(compact_formula_alpha_tokens(text, alt_tokens))
    overlap = len(overlap_tokens)
    strong_overlap = sum(1 for token in overlap_tokens if is_strong_formula_token(token))
    math_words = sum(1 for word in segment if is_math_word(pdf_word_text(word)))
    symbol_count = len(re.findall(r"[=+\-*/_(){}\[\]|:]", text))
    long_words = sum(
        1
        for word in segment
        if is_long_prose_word(pdf_word_text(word)) and not formula_word_overlap_tokens(pdf_word_text(word), alt_tokens)
    )
    score = (overlap * 6.0) + (strong_overlap * 4.0) + (symbol_count * 1.5) + math_words - (long_words * 4.0)
    return {
        "segment": segment,
        "text": text,
        "overlap": overlap,
        "strong_overlap": strong_overlap,
        "symbol_count": symbol_count,
        "math_words": math_words,
        "long_words": long_words,
        "score": score,
    }


def choose_formula_line_words(words: list[tuple], alt_tokens: set[str]) -> list[tuple]:
    segments = split_formula_word_segments(words, alt_tokens)
    if not segments:
        return words
    if len(segments) == 1:
        return segments[0]

    stats = sorted(
        (formula_segment_stats(segment, alt_tokens) for segment in segments),
        key=lambda item: (float(item["score"]), int(item["strong_overlap"]), int(item["overlap"])),
        reverse=True,
    )
    if alt_tokens and int(stats[0]["overlap"]) > 0:
        runner_up = stats[1] if len(stats) > 1 else None
        if (
            runner_up is None
            or int(stats[0]["strong_overlap"]) > int(runner_up["strong_overlap"])
            or float(stats[0]["score"]) >= float(runner_up["score"]) + 3.0
            or int(stats[0]["overlap"]) >= int(runner_up["overlap"]) + 2
        ):
            return list(stats[0]["segment"])

    selected: list[tuple] = []
    for segment in segments:
        selected.extend(segment)
    return selected or words


def line_gap(left: dict, right: dict) -> float:
    if left["bbox"][3] < right["bbox"][1]:
        return float(right["bbox"][1] - left["bbox"][3])
    if right["bbox"][3] < left["bbox"][1]:
        return float(left["bbox"][1] - right["bbox"][3])
    return 0.0


def x_overlap_ratio(left: dict, right: dict) -> float:
    left_width = max(0.01, float(left["bbox"][2] - left["bbox"][0]))
    right_width = max(0.01, float(right["bbox"][2] - right["bbox"][0]))
    overlap = max(0.0, min(float(left["bbox"][2]), float(right["bbox"][2])) - max(float(left["bbox"][0]), float(right["bbox"][0])))
    return overlap / min(left_width, right_width)


def formula_line_candidates(page: fitz.Page, bbox: list[float], alt_text: str) -> list[dict]:
    if not valid_content_bbox(bbox):
        return []
    width = max(0.0, float(bbox[2]) - float(bbox[0]))
    height = max(0.0, float(bbox[3]) - float(bbox[1]))
    x_pad = 26.0 if width < 30 else 4.0
    y_pad = 12.0 if height < 24 else 4.0
    search_rect = expand_fitz_bbox(bbox, page.rect, x_pad=x_pad, y_pad=y_pad)
    alt_tokens = formula_alt_tokens(alt_text)

    grouped: dict[tuple[int, int], list[tuple]] = {}
    for word in page.get_text("words", clip=search_rect) or []:
        if len(word) < 8:
            continue
        key = (int(word[5]), int(word[6]))
        grouped.setdefault(key, []).append(word)

    candidates: list[dict] = []
    for words in grouped.values():
        # Formula reading order is left-to-right first; superscripts often sit slightly above the base glyph.
        words.sort(key=lambda item: (float(item[0]), float(item[1])))
        selected_words = choose_formula_line_words(words, alt_tokens)
        text = " ".join(pdf_word_text(word) for word in selected_words)
        word_bboxes = [formula_word_bbox(word) for word in selected_words]
        line_bbox = union_content_bboxes(word_bboxes)
        if line_bbox is None:
            continue
        math_words = sum(1 for word in selected_words if is_math_word(pdf_word_text(word)))
        long_words = sum(
            1
            for word in selected_words
            if is_long_prose_word(pdf_word_text(word)) and not formula_word_overlap_tokens(pdf_word_text(word), alt_tokens)
        )
        symbol_count = len(re.findall(r"[=+*/⁄_(){}\[\]|∣·×÷−-]", text))
        line_tokens = visual_formula_tokens(text)
        overlap_tokens = line_tokens & alt_tokens if alt_tokens else set()
        overlap_tokens.update(compact_formula_alpha_tokens(text, alt_tokens))
        overlap = len(overlap_tokens)
        strong_overlap = sum(1 for token in overlap_tokens if is_strong_formula_token(token))
        token_count = max(1, len(selected_words))
        score = (math_words * 2.0) + symbol_count + (overlap * 3.0) - (long_words * 2.5)
        if long_words >= 3 and strong_overlap == 0 and symbol_count < 3:
            continue
        if long_words >= 2 and math_words <= 2 and strong_overlap == 0 and symbol_count < 2:
            continue
        if long_words >= 3 and overlap == 0 and symbol_count < 2:
            continue
        if score < 1.5 and math_words / token_count < 0.45 and overlap == 0:
            continue
        candidates.append(
            {
                "bbox": line_bbox,
                "text": text,
                "line_tokens": line_tokens,
                "overlap_tokens": overlap_tokens,
                "score": score,
                "overlap": overlap,
                "strong_overlap": strong_overlap,
                "symbol_count": symbol_count,
                "math_words": math_words,
                "long_words": long_words,
            }
        )
    return sorted(candidates, key=lambda item: (float(item["bbox"][1]), float(item["bbox"][0])))


def cluster_formula_candidates(candidates: list[dict]) -> list[list[dict]]:
    if not candidates:
        return []
    ordered = sorted(candidates, key=lambda item: (float(item["bbox"][1]), float(item["bbox"][0])))
    clusters: list[list[dict]] = []
    current: list[dict] = []
    current_bbox: list[float] | None = None
    for candidate in ordered:
        if not current:
            current = [candidate]
            current_bbox = list(candidate["bbox"])
            continue
        gap = line_gap({"bbox": current_bbox}, candidate) if current_bbox is not None else 999.0
        if gap <= 0.75:
            current.append(candidate)
            current_bbox = union_content_bboxes([tuple(item["bbox"]) for item in current])
            continue
        clusters.append(current)
        current = [candidate]
        current_bbox = list(candidate["bbox"])
    if current:
        clusters.append(current)
    return clusters


def formula_cluster_stats(cluster: list[dict], alt_tokens: set[str], content_bbox: list[float]) -> dict:
    cluster_bbox = union_content_bboxes([tuple(candidate["bbox"]) for candidate in cluster])
    if cluster_bbox is None:
        cluster_bbox = list(content_bbox)
    cluster_tokens: set[str] = set()
    cluster_compact_tokens: set[str] = set()
    total_symbol_count = 0
    total_math_words = 0
    total_long_words = 0
    total_score = 0.0
    for candidate in cluster:
        cluster_tokens.update(set(candidate.get("line_tokens", set())))
        cluster_compact_tokens.update(compact_formula_alpha_tokens(str(candidate.get("text") or ""), alt_tokens))
        total_symbol_count += int(candidate.get("symbol_count", 0))
        total_math_words += int(candidate.get("math_words", 0))
        total_long_words += int(candidate.get("long_words", 0))
        total_score += float(candidate.get("score", 0.0))
    overlap_tokens = cluster_tokens & alt_tokens if alt_tokens else set()
    overlap_tokens.update(cluster_compact_tokens)
    strong_overlap = sum(1 for token in overlap_tokens if is_strong_formula_token(token))
    overlap = len(overlap_tokens)
    content_overlap_ratio = 0.0
    content_area = bbox_area(content_bbox)
    if content_area > 0:
        content_overlap_ratio = bbox_intersection_area(cluster_bbox, content_bbox) / content_area
    return {
        "candidates": cluster,
        "bbox": cluster_bbox,
        "cluster_tokens": cluster_tokens,
        "overlap_tokens": overlap_tokens,
        "overlap": overlap,
        "strong_overlap": strong_overlap,
        "symbol_count": total_symbol_count,
        "math_words": total_math_words,
        "long_words": total_long_words,
        "score": total_score,
        "content_overlap_ratio": content_overlap_ratio,
    }


def formula_cluster_sort_key(cluster: dict) -> tuple[float, float, float, float, float]:
    return (
        float(cluster["strong_overlap"]),
        float(cluster["overlap"]),
        float(cluster["symbol_count"]),
        float(cluster["math_words"] - (cluster["long_words"] * 2)),
        float(cluster["score"]),
    )


def strong_formula_alt_tokens(alt_tokens: set[str]) -> set[str]:
    return {token for token in alt_tokens if is_strong_formula_token(token)}


def should_add_formula_cluster(
    selected_clusters: list[dict],
    candidate_cluster: dict,
    covered_tokens: set[str],
    required_strong_tokens: set[str],
    *,
    max_gap: float,
    min_x_overlap: float,
) -> bool:
    new_tokens = set(candidate_cluster["overlap_tokens"]) - covered_tokens
    if not new_tokens:
        return False
    if required_strong_tokens:
        covered_strong_tokens = {token for token in covered_tokens if is_strong_formula_token(token)}
        if required_strong_tokens.issubset(covered_strong_tokens) and not any(
            is_strong_formula_token(token) for token in new_tokens
        ):
            return False
    return any(
        line_gap({"bbox": candidate_cluster["bbox"]}, {"bbox": chosen["bbox"]}) <= max_gap
        and x_overlap_ratio({"bbox": candidate_cluster["bbox"]}, {"bbox": chosen["bbox"]}) >= min_x_overlap
        for chosen in selected_clusters
    )


def refine_formula_content_bbox(page: fitz.Page, content_bbox: list[float] | None, alt_text: str) -> list[float] | None:
    if not valid_content_bbox(content_bbox):
        return content_bbox
    candidates = formula_line_candidates(page, list(content_bbox or []), alt_text)
    if not candidates:
        return content_bbox

    alt_tokens = formula_alt_tokens(alt_text)
    required_strong_tokens = strong_formula_alt_tokens(alt_tokens)
    clusters = [formula_cluster_stats(cluster, alt_tokens, list(content_bbox or [])) for cluster in cluster_formula_candidates(candidates)]
    if not clusters:
        return content_bbox

    anchor = max(clusters, key=formula_cluster_sort_key)
    selected_clusters = [anchor]
    covered_tokens = set(anchor["overlap_tokens"])
    changed = True
    while changed:
        changed = False
        for cluster in sorted(clusters, key=formula_cluster_sort_key, reverse=True):
            if cluster in selected_clusters:
                continue
            if not should_add_formula_cluster(
                selected_clusters,
                cluster,
                covered_tokens,
                required_strong_tokens,
                max_gap=6.5,
                min_x_overlap=0.12,
            ):
                continue
            selected_clusters.append(cluster)
            covered_tokens.update(set(cluster["overlap_tokens"]) - covered_tokens)
            changed = True

    refined = union_content_bboxes([tuple(cluster["bbox"]) for cluster in selected_clusters])
    if refined is None:
        return content_bbox
    original_area = bbox_area(content_bbox)
    refined_area = bbox_area(refined)
    if original_area > 0 and refined_area > original_area * 1.12:
        if covered_tokens and refined_area <= original_area * 1.75:
            return refined
        return content_bbox
    return refined


def formula_word_search_items(page: fitz.Page, alt_text: str) -> list[dict]:
    alt_tokens = formula_alt_tokens(alt_text)
    items: list[dict] = []
    for word in page.get_text("words") or []:
        text = pdf_word_text(word)
        cleaned = normalize_alt_text(text).strip()
        if not cleaned:
            continue
        overlap_tokens = formula_word_overlap_tokens(text, alt_tokens)
        formulaish = is_math_word(text) or bool(overlap_tokens)
        if not formulaish:
            continue
        lowered = cleaned.lower()
        alpha = re.sub(r"[^a-z]", "", lowered)
        if not overlap_tokens and (lowered in SHORT_FORMULA_PROSE_WORDS or lowered in FORMULA_RUN_BREAK_WORDS or is_long_prose_word(text)):
            continue
        if not overlap_tokens and alpha and len(alpha) <= 3 and alpha not in SHORT_FORMULA_FUNCTION_WORDS and not re.search(r"\d", cleaned):
            continue
        items.append(
            {
                "word": word,
                "text": text,
                "bbox": formula_word_bbox(word),
                "tokens": visual_formula_tokens(text),
                "overlap_tokens": overlap_tokens,
            }
        )
    return items


def score_formula_word_cluster(cluster: list[dict], alt_tokens: set[str]) -> tuple[float, float, float, float, float]:
    cluster_bbox = union_content_bboxes([tuple(item["bbox"]) for item in cluster]) or [0.0, 0.0, 0.0, 0.0]
    cluster_tokens: set[str] = set()
    symbol_count = 0
    formula_anchor_count = 0
    for item in cluster:
        cluster_tokens.update(set(item["tokens"]))
        symbol_count += len(re.findall(r"[=+\-*/_(){}\[\]|:]", str(item["text"] or "")))
        text = normalize_alt_text(str(item.get("text") or "")).strip()
        if re.search(r"[=+\-*/_(){}\[\]|:]", text) or re.search(r"\d", text) or len(re.sub(r"[^a-z]", "", text.lower())) >= 2:
            formula_anchor_count += 1
    overlap_tokens = cluster_tokens & alt_tokens if alt_tokens else set()
    for item in cluster:
        overlap_tokens.update(compact_formula_alpha_tokens(str(item.get("text") or ""), alt_tokens))
    overlap = len(overlap_tokens)
    strong_overlap = sum(1 for token in overlap_tokens if is_strong_formula_token(token))
    return (
        float(formula_anchor_count),
        float(symbol_count),
        float(overlap),
        float(strong_overlap),
        -bbox_area(cluster_bbox) / max(1.0, len(cluster)),
    )


def locate_formula_bbox_by_word_search(page: fitz.Page, alt_text: str) -> list[float] | None:
    alt_tokens = formula_alt_tokens(alt_text)
    required_strong_tokens = strong_formula_alt_tokens(alt_tokens)
    full_page_bbox = [float(page.rect.x0), float(page.rect.y0), float(page.rect.x1), float(page.rect.y1)]
    line_candidates = formula_line_candidates(page, full_page_bbox, alt_text)
    if line_candidates:
        clusters = [formula_cluster_stats(cluster, alt_tokens, full_page_bbox) for cluster in cluster_formula_candidates(line_candidates)]
        if clusters:
            anchor = max(clusters, key=formula_cluster_sort_key)
            selected_clusters = [anchor]
            covered_tokens = set(anchor["overlap_tokens"])
            changed = True
            while changed:
                changed = False
                for cluster in sorted(clusters, key=formula_cluster_sort_key, reverse=True):
                    if cluster in selected_clusters:
                        continue
                    if not should_add_formula_cluster(
                        selected_clusters,
                        cluster,
                        covered_tokens,
                        required_strong_tokens,
                        max_gap=4.0,
                        min_x_overlap=0.2,
                    ):
                        continue
                    selected_clusters.append(cluster)
                    covered_tokens.update(set(cluster["overlap_tokens"]) - covered_tokens)
                    changed = True
            refined = union_content_bboxes([tuple(cluster["bbox"]) for cluster in selected_clusters])
            if valid_content_bbox(refined):
                return refined

    items = formula_word_search_items(page, alt_text)
    if not items:
        return None

    seed_items = [item for item in items if item["overlap_tokens"]] or list(items)
    best_bbox: list[float] | None = None
    best_score: tuple[float, float, float, float] | None = None
    for seed in seed_items:
        cluster = [seed]
        cluster_ids = {id(seed)}
        current_bbox = list(seed["bbox"])
        changed = True
        while changed:
            changed = False
            for candidate in items:
                if id(candidate) in cluster_ids:
                    continue
                gap_x, gap_y = bbox_gap(current_bbox, candidate["bbox"])
                if gap_y > 6.0 or gap_x > 18.0:
                    continue
                cluster.append(candidate)
                cluster_ids.add(id(candidate))
                current_bbox = union_content_bboxes([tuple(item["bbox"]) for item in cluster]) or current_bbox
                changed = True
        score = score_formula_word_cluster(cluster, alt_tokens)
        cluster_bbox = union_content_bboxes([tuple(item["bbox"]) for item in cluster])
        if cluster_bbox is None:
            continue
        if best_score is None or score > best_score:
            best_score = score
            best_bbox = cluster_bbox
    return best_bbox


def locate_formula_bbox_by_page_search(page: fitz.Page, alt_text: str) -> list[float] | None:
    word_search_bbox = locate_formula_bbox_by_word_search(page, alt_text)
    if valid_content_bbox(word_search_bbox):
        return word_search_bbox
    full_page_bbox = [float(page.rect.x0), float(page.rect.y0), float(page.rect.x1), float(page.rect.y1)]
    refined = refine_formula_content_bbox(page, full_page_bbox, alt_text)
    if not valid_content_bbox(refined):
        return None
    page_area = max(1.0, float(page.rect.width * page.rect.height))
    refined_area = bbox_area(refined)
    if refined_area <= 0 or refined_area / page_area >= 0.12:
        return None
    return refined


def formula_bbox_signal(page: fitz.Page, bbox: list[float] | None, alt_text: str) -> dict[str, float]:
    if not valid_content_bbox(bbox):
        return {
            "overlap": 0.0,
            "strong_overlap": 0.0,
            "symbol_count": 0.0,
            "math_words": 0.0,
        }

    alt_tokens = formula_alt_tokens(alt_text)
    clip = expand_fitz_bbox(list(bbox or []), page.rect, x_pad=1.5, y_pad=1.0)
    words = list(page.get_text("words", clip=clip) or [])
    if not words:
        return {
            "overlap": 0.0,
            "strong_overlap": 0.0,
            "symbol_count": 0.0,
            "math_words": 0.0,
        }

    text = " ".join(pdf_word_text(word) for word in words)
    overlap_tokens = visual_formula_tokens(text) & alt_tokens if alt_tokens else set()
    overlap_tokens.update(compact_formula_alpha_tokens(text, alt_tokens))
    return {
        "overlap": float(len(overlap_tokens)),
        "strong_overlap": float(sum(1 for token in overlap_tokens if is_strong_formula_token(token))),
        "symbol_count": float(len(re.findall(r"[=+\-*/â„_(){}\[\]|âˆ£Â·Ã—Ã·âˆ’-]", text))),
        "math_words": float(sum(1 for word in words if is_math_word(pdf_word_text(word)))),
    }


def should_prefer_formula_word_search_bbox(
    page: fitz.Page,
    content_bbox: list[float] | None,
    word_search_bbox: list[float] | None,
    alt_text: str,
) -> bool:
    if not valid_content_bbox(word_search_bbox):
        return False
    if not valid_content_bbox(content_bbox):
        return True
    word_area = bbox_area(word_search_bbox)
    content_area = bbox_area(content_bbox)
    if word_area <= 0 or content_area <= 0:
        return False
    word_signal = formula_bbox_signal(page, word_search_bbox, alt_text)
    content_signal = formula_bbox_signal(page, content_bbox, alt_text)
    if word_signal["strong_overlap"] < content_signal["strong_overlap"]:
        return False
    if word_signal["overlap"] < content_signal["overlap"]:
        return False
    if word_signal["symbol_count"] + 1 < content_signal["symbol_count"] and word_signal["overlap"] <= content_signal["overlap"]:
        return False
    overlap = bbox_intersection_area(content_bbox, word_search_bbox)
    if overlap / word_area < 0.5:
        gap_x, gap_y = bbox_gap(content_bbox, word_search_bbox)
        if gap_x > 12.0 or gap_y > 8.0:
            return False
    return word_area <= content_area * 0.92


def tighten_formula_bbox_to_declared_wrapper_span(page: fitz.Page, content_bbox: list[float] | None, alt_text: str) -> list[float] | None:
    if not valid_content_bbox(content_bbox):
        return content_bbox

    normalized = normalize_alt_text(str(alt_text or "")).lower()
    wrapper_pairs: list[tuple[str, str]] = []
    if "brace" in normalized:
        wrapper_pairs.append(("{", "}"))
    if "parenthesis" in normalized or "paren" in normalized:
        wrapper_pairs.append(("(", ")"))
    if "bracket" in normalized:
        wrapper_pairs.append(("[", "]"))
    if not wrapper_pairs:
        return content_bbox

    clip = expand_fitz_bbox(list(content_bbox), page.rect, x_pad=2.5, y_pad=1.5)
    words = list(page.get_text("words", clip=clip) or [])
    if not words:
        return content_bbox
    words.sort(key=lambda item: (float(item[1]), float(item[0])))

    original_area = bbox_area(content_bbox)
    for opening_wrapper, closing_wrapper in wrapper_pairs:
        start_index = next((index for index, word in enumerate(words) if opening_wrapper in pdf_word_text(word)), None)
        end_index = next((index for index in range(len(words) - 1, -1, -1) if closing_wrapper in pdf_word_text(words[index])), None)
        if start_index is None or end_index is None or start_index > end_index:
            continue
        wrapper_bbox = union_content_bboxes([formula_word_bbox(word) for word in words[start_index : end_index + 1]])
        if not valid_content_bbox(wrapper_bbox):
            continue
        if bbox_area(wrapper_bbox) <= original_area * 1.02:
            return wrapper_bbox
    return content_bbox


def expand_formula_bbox_for_wrapping_glyphs(page: fitz.Page, content_bbox: list[float] | None, alt_text: str) -> list[float] | None:
    if not valid_content_bbox(content_bbox):
        return content_bbox

    normalized = normalize_alt_text(str(alt_text or "")).lower()
    left_pad = 0.0
    right_pad = 0.0
    if re.search(r"\b(open|left|start)\s+(parenthesis|paren|bracket|brace)\b", normalized):
        left_pad = 8.0
    if re.search(r"\b(close|right|end)\s+(parenthesis|paren|bracket|brace)\b", normalized):
        right_pad = 4.0
    if left_pad <= 0 and right_pad <= 0:
        return content_bbox

    clip = expand_fitz_bbox(list(content_bbox), page.rect, x_pad=0.75, y_pad=0.75)
    visual_text = " ".join(str(word[4]) for word in page.get_text("words", clip=clip) or [])
    has_left_wrapper = any(character in visual_text for character in "([{")
    has_right_wrapper = any(character in visual_text for character in ")]}")
    if (left_pad <= 0 or has_left_wrapper) and (right_pad <= 0 or has_right_wrapper):
        return content_bbox

    return [
        max(float(page.rect.x0), float(content_bbox[0]) - left_pad),
        float(content_bbox[1]),
        min(float(page.rect.x1), float(content_bbox[2]) + right_pad),
        float(content_bbox[3]),
    ]


def expand_formula_bbox_for_fine_math_glyphs(page: fitz.Page, content_bbox: list[float] | None, alt_text: str) -> list[float] | None:
    if not valid_content_bbox(content_bbox):
        return content_bbox

    normalized = normalize_alt_text(str(alt_text or "")).lower()
    x_pad = 1.1
    top_pad = 1.8
    bottom_pad = 1.1
    if any(
        marker in normalized
        for marker in (
            "superscript",
            "subscript",
            "squared",
            "cubed",
            "root",
            "fraction",
            "over",
            "power",
        )
    ):
        x_pad = 1.35
        top_pad = 2.3
        bottom_pad = 1.35
    if re.search(r"\b(open|close|left|right)\s+(parenthesis|paren|bracket|brace)\b", normalized):
        x_pad = max(x_pad, 1.5)

    return [
        max(float(page.rect.x0), float(content_bbox[0]) - x_pad),
        max(float(page.rect.y0), float(content_bbox[1]) - top_pad),
        min(float(page.rect.x1), float(content_bbox[2]) + x_pad),
        min(float(page.rect.y1), float(content_bbox[3]) + bottom_pad),
    ]


def structure_marked_content_bbox(
    document: fitz.Document,
    struct_xref: int,
    page_index: int,
    page_mcid_bboxes: dict[int, list[float]],
    role_map: dict[str, str],
) -> list[float] | None:
    mcids = owned_structure_mcids(document, struct_xref, role_map)
    if not mcids:
        return None
    bboxes = [page_mcid_bboxes[mcid] for mcid in mcids if mcid in page_mcid_bboxes]
    return union_content_bboxes([tuple(bbox) for bbox in bboxes])


def structure_marked_content_components(
    document: fitz.Document,
    struct_xref: int,
    page_mcid_components: dict[int, dict[str, list[tuple[float, float, float, float]]]],
    role_map: dict[str, str],
) -> dict[str, list[tuple[float, float, float, float]]]:
    merged = {"image": [], "paint": [], "text": []}
    mcids = owned_structure_mcids(document, struct_xref, role_map)
    if not mcids:
        return merged
    for mcid in mcids:
        groups = page_mcid_components.get(mcid)
        if not groups:
            continue
        for kind in ("image", "paint", "text"):
            merged[kind].extend(groups.get(kind, []))
    return {kind: dedupe_bboxes(values) for kind, values in merged.items()}


def cluster_component_bboxes(
    bboxes: list[tuple[float, float, float, float]],
    *,
    gap_x: float,
    gap_y: float,
) -> list[list[tuple[float, float, float, float]]]:
    remaining = list(dedupe_bboxes([bbox for bbox in bboxes if valid_content_bbox(bbox)]))
    clusters: list[list[tuple[float, float, float, float]]] = []
    while remaining:
        cluster = [remaining.pop(0)]
        changed = True
        while changed:
            changed = False
            current_union = union_content_bboxes(cluster)
            if current_union is None:
                continue
            next_remaining: list[tuple[float, float, float, float]] = []
            for candidate in remaining:
                candidate_gap_x, candidate_gap_y = bbox_gap(current_union, candidate)
                if candidate_gap_x <= gap_x and candidate_gap_y <= gap_y:
                    cluster.append(candidate)
                    changed = True
                else:
                    next_remaining.append(candidate)
            remaining = next_remaining
        clusters.append(cluster)
    return clusters


def score_figure_cluster(
    cluster_bbox: list[float],
    *,
    pdf_bbox: list[float] | None,
    page_rect: fitz.Rect,
    component_count: int,
) -> tuple[float, float, float, float]:
    cluster_area = bbox_area(cluster_bbox)
    page_area = max(1.0, float(page_rect.width * page_rect.height))
    if valid_content_bbox(pdf_bbox):
        pdf_area = bbox_area(pdf_bbox)
        overlap = bbox_intersection_area(cluster_bbox, pdf_bbox or [])
        overlap_ratio = overlap / max(1.0, min(cluster_area, pdf_area))
        gap_x, gap_y = bbox_gap(cluster_bbox, pdf_bbox or [])
        cluster_cx, cluster_cy = bbox_center(cluster_bbox)
        pdf_cx, pdf_cy = bbox_center(pdf_bbox or [])
        center_distance = math.hypot(cluster_cx - pdf_cx, cluster_cy - pdf_cy)
        return (
            overlap_ratio,
            float(component_count),
            cluster_area / page_area,
            -((gap_x + gap_y) + center_distance * 0.05),
        )
    return (
        0.0,
        float(component_count),
        cluster_area / page_area,
        -sum(bbox_center(cluster_bbox)),
    )


def select_anchor_cluster_bbox(
    page: fitz.Page,
    pdf_bbox: list[float] | None,
    components: dict[str, list[tuple[float, float, float, float]]],
) -> tuple[list[float], list[tuple[float, float, float, float]]] | None:
    visual_components = components.get("image", []) or components.get("paint", []) or []
    if not visual_components:
        visual_components = components.get("text", [])
    if not visual_components:
        return None

    clusters = cluster_component_bboxes(
        visual_components,
        gap_x=24.0 if components.get("image") else 16.0,
        gap_y=20.0 if components.get("image") else 16.0,
    )
    scored_clusters: list[tuple[tuple[float, float, float, float], list[tuple[float, float, float, float]]]] = []
    for cluster in clusters:
        cluster_bbox = union_content_bboxes(cluster)
        if cluster_bbox is None:
            continue
        scored_clusters.append(
            (
                score_figure_cluster(
                    cluster_bbox,
                    pdf_bbox=pdf_bbox,
                    page_rect=page.rect,
                    component_count=len(cluster),
                ),
                cluster,
            )
        )
    if not scored_clusters:
        return None
    best_cluster = max(scored_clusters, key=lambda item: item[0])[1]
    best_bbox = union_content_bboxes(best_cluster)
    if best_bbox is None:
        return None
    return best_bbox, best_cluster


def should_attach_label_to_figure(
    anchor_bbox: list[float],
    candidate_bbox: tuple[float, float, float, float],
    *,
    kind: str,
) -> bool:
    gap_x, gap_y = bbox_gap(anchor_bbox, candidate_bbox)
    anchor_width = max(1.0, float(anchor_bbox[2]) - float(anchor_bbox[0]))
    anchor_height = max(1.0, float(anchor_bbox[3]) - float(anchor_bbox[1]))
    candidate_width = max(1.0, float(candidate_bbox[2]) - float(candidate_bbox[0]))
    candidate_height = max(1.0, float(candidate_bbox[3]) - float(candidate_bbox[1]))
    candidate_area = bbox_area(candidate_bbox)
    anchor_area = max(1.0, bbox_area(anchor_bbox))

    if kind == "text":
        if candidate_height > max(28.0, anchor_height * 0.7):
            return False
        if candidate_width > max(180.0, anchor_width * 1.5):
            return False
        if candidate_area > anchor_area * 0.7:
            return False
        return gap_x <= max(42.0, anchor_width * 0.28) and gap_y <= max(26.0, anchor_height * 0.3)

    if kind == "paint":
        return gap_x <= max(20.0, anchor_width * 0.15) and gap_y <= max(20.0, anchor_height * 0.2)

    return gap_x <= 8.0 and gap_y <= 8.0


def is_decorative_paint_bbox(page: fitz.Page, bbox: tuple[float, float, float, float]) -> bool:
    width = max(0.0, float(bbox[2]) - float(bbox[0]))
    height = max(0.0, float(bbox[3]) - float(bbox[1]))
    page_width = max(1.0, float(page.rect.width))
    page_height = max(1.0, float(page.rect.height))
    is_thin_horizontal_rule = height <= 6.0 and width >= page_width * 0.45
    is_thin_vertical_rule = width <= 6.0 and height >= page_height * 0.28
    return is_thin_horizontal_rule or is_thin_vertical_rule


def filter_figure_paint_components(
    page: fitz.Page,
    paint_bboxes: list[tuple[float, float, float, float]],
) -> list[tuple[float, float, float, float]]:
    if not paint_bboxes:
        return []
    candidate_areas = [bbox_area(bbox) for bbox in paint_bboxes if bbox_area(bbox) > 0]
    median_area = median_number(candidate_areas)
    page_width = max(1.0, float(page.rect.width))
    page_height = max(1.0, float(page.rect.height))
    filtered: list[tuple[float, float, float, float]] = []
    for bbox in paint_bboxes:
        width = max(0.0, float(bbox[2]) - float(bbox[0]))
        height = max(0.0, float(bbox[3]) - float(bbox[1]))
        area = bbox_area(bbox)
        is_large_frame = (
            median_area > 0
            and area >= median_area * 18.0
            and width >= page_width * 0.55
            and height >= page_height * 0.4
        )
        if is_decorative_paint_bbox(page, bbox) or is_large_frame:
            continue
        filtered.append(bbox)
    return filtered


def refine_figure_content_bbox(
    page: fitz.Page,
    pdf_bbox: list[float] | None,
    components: dict[str, list[tuple[float, float, float, float]]],
) -> list[float] | None:
    filtered_paint = filter_figure_paint_components(page, list(components.get("paint", [])))
    normalized_components = {
        "image": list(components.get("image", [])),
        "paint": filtered_paint or list(components.get("paint", [])),
        "text": list(components.get("text", [])),
    }
    anchor_data = select_anchor_cluster_bbox(page, pdf_bbox, normalized_components)
    if anchor_data is None:
        values = normalized_components.get("image", []) + normalized_components.get("paint", []) + normalized_components.get("text", [])
        return union_content_bboxes(values)
    anchor_bbox, anchor_cluster = anchor_data

    selected: list[tuple[float, float, float, float]] = list(anchor_cluster)
    if not selected:
        selected.append(tuple(float(value) for value in anchor_bbox[:4]))

    current_bbox = union_content_bboxes(selected) or anchor_bbox
    changed = True
    while changed:
        changed = False
        for kind in ("image", "paint", "text"):
            for candidate in normalized_components.get(kind, []):
                candidate_tuple = tuple(float(value) for value in candidate[:4])
                if candidate_tuple in selected:
                    continue
                if should_attach_label_to_figure(current_bbox, candidate_tuple, kind=kind):
                    selected.append(candidate_tuple)
                    current_bbox = union_content_bboxes(selected) or current_bbox
                    changed = True
    return current_bbox


def structure_page_xref(document: fitz.Document, struct_xref: int) -> int | None:
    try:
        value_type, value = document.xref_get_key(struct_xref, "Pg")
    except Exception:
        value_type, value = ("null", "null")
    if value_type == "xref":
        page_xref = parse_pdf_ref(value)
        if isinstance(page_xref, int):
            return page_xref

    try:
        struct_object = document.xref_object(struct_xref, compressed=False)
    except Exception:
        return None
    match = re.search(r"/Pg\s+(\d+)\s+\d+\s+R", struct_object)
    if not match:
        return None
    try:
        return int(match.group(1))
    except ValueError:
        return None


def bbox_from_xref(document: fitz.Document, xref: int, visited: set[int] | None = None) -> list[float] | None:
    visited = visited or set()
    if xref in visited:
        return None
    visited.add(xref)

    try:
        raw_object = document.xref_object(xref, compressed=False)
    except Exception:
        raw_object = ""
    direct_bbox = parse_number_array(raw_object)
    if raw_object.strip().startswith("[") and direct_bbox is not None:
        return direct_bbox

    for key in ("BBox", "Rect"):
        try:
            value_type, value = document.xref_get_key(xref, key)
        except Exception:
            continue
        bbox = bbox_from_pdf_value(document, value_type, value, visited)
        if bbox is not None:
            return bbox

    match = PDF_BBOX_RE.search(raw_object)
    if match:
        bbox = bbox_from_pdf_literal(document, match.group("value"), visited)
        if bbox is not None:
            return bbox
    return None


def bbox_from_pdf_literal(document: fitz.Document, value: str, visited: set[int]) -> list[float] | None:
    ref = parse_pdf_ref(value)
    if isinstance(ref, int):
        return bbox_from_xref(document, ref, visited)
    return parse_number_array(value)


def bbox_from_pdf_value(
    document: fitz.Document,
    value_type: str,
    value: str,
    visited: set[int] | None = None,
) -> list[float] | None:
    visited = visited or set()
    if value_type == "xref":
        ref = parse_pdf_ref(value)
        return bbox_from_xref(document, ref, visited) if isinstance(ref, int) else None
    if value_type in {"array", "dict"}:
        return bbox_from_pdf_literal(document, value, visited)
    return None


def structure_bbox(document: fitz.Document, struct_xref: int) -> list[float] | None:
    try:
        value_type, value = document.xref_get_key(struct_xref, "BBox")
        bbox = bbox_from_pdf_value(document, value_type, value)
        if bbox is not None:
            return bbox
    except Exception:
        pass

    try:
        value_type, value = document.xref_get_key(struct_xref, "A")
    except Exception:
        return None

    bbox = bbox_from_pdf_value(document, value_type, value)
    if bbox is not None:
        return bbox

    for ref in parse_pdf_refs(value):
        bbox = bbox_from_xref(document, ref)
        if bbox is not None:
            return bbox
    return None


def pdf_bbox_to_preview_norm(page: fitz.Page, bbox: list[float] | None) -> dict[str, float] | None:
    if not bbox:
        return None
    x0, y0, x1, y1 = bbox[:4]
    left, right = sorted((x0, x1))
    bottom, top = sorted((y0, y1))
    page_rect = page.rect
    if page_rect.width <= 0 or page_rect.height <= 0:
        return None

    norm = {
        "x0": (left - page_rect.x0) / page_rect.width,
        "y0": (page_rect.y1 - top) / page_rect.height,
        "x1": (right - page_rect.x0) / page_rect.width,
        "y1": (page_rect.y1 - bottom) / page_rect.height,
    }
    return {key: max(0.0, min(1.0, float(value))) for key, value in norm.items()}


def fitz_bbox_to_preview_norm(page: fitz.Page, bbox: list[float] | None) -> dict[str, float] | None:
    if not bbox:
        return None
    page_rect = page.rect
    if page_rect.width <= 0 or page_rect.height <= 0:
        return None
    x0, y0, x1, y1 = bbox[:4]
    left, right = sorted((float(x0), float(x1)))
    top, bottom = sorted((float(y0), float(y1)))
    norm = {
        "x0": (left - page_rect.x0) / page_rect.width,
        "y0": (top - page_rect.y0) / page_rect.height,
        "x1": (right - page_rect.x0) / page_rect.width,
        "y1": (bottom - page_rect.y0) / page_rect.height,
    }
    return {key: max(0.0, min(1.0, float(value))) for key, value in norm.items()}


def should_prefer_content_bbox_for_figure(page: fitz.Page, pdf_bbox: list[float] | None, content_bbox: list[float] | None) -> bool:
    if not valid_content_bbox(content_bbox):
        return False
    if not valid_content_bbox(pdf_bbox):
        return True

    page_area = max(1.0, float(page.rect.width * page.rect.height))
    pdf_area = bbox_area(pdf_bbox)
    content_area = bbox_area(content_bbox)
    pdf_width = max(1.0, float(pdf_bbox[2]) - float(pdf_bbox[0]))
    pdf_height = max(1.0, float(pdf_bbox[3]) - float(pdf_bbox[1]))
    content_width = max(1.0, float(content_bbox[2]) - float(content_bbox[0]))
    content_height = max(1.0, float(content_bbox[3]) - float(content_bbox[1]))
    if content_area <= 0:
        return False
    overlap_area = bbox_intersection_area(pdf_bbox, content_bbox)
    content_coverage = overlap_area / content_area
    pdf_aspect = max(pdf_width, pdf_height) / min(pdf_width, pdf_height)
    content_aspect = max(content_width, content_height) / min(content_width, content_height)
    if content_coverage < 0.7:
        return True
    if pdf_width <= content_width * 0.72 or pdf_height <= content_height * 0.72:
        return True
    if pdf_aspect >= content_aspect * 2.5 or content_aspect >= pdf_aspect * 2.5:
        return True
    if pdf_area / page_area >= 0.40 and content_area / page_area <= 0.32:
        return True
    return pdf_area >= content_area * 3.0 and content_area >= 250.0


def pdf_alt_text(document: fitz.Document, struct_xref: int) -> str:
    try:
        value_type, value = document.xref_get_key(struct_xref, "Alt")
    except Exception:
        return ""
    if value_type not in {"string", "name"}:
        return ""
    return normalize_alt_text(str(value or "").replace("\x00", ""))


def validate_pdf_file(pdf_path: Path) -> str | None:
    try:
        with pdf_path.open("rb") as handle:
            head = handle.read(5)
    except OSError as exc:
        return f"{pdf_path.name} could not be read: {exc}"

    if head != b"%PDF-":
        return f"{pdf_path.name} is not a PDF file."

    try:
        with fitz.open(str(pdf_path)) as document:
            if document.page_count <= 0:
                return f"{pdf_path.name} does not contain any pages."
    except Exception as exc:
        return f"{pdf_path.name} could not be opened as a PDF: {exc}"
    return None


def build_pdf_alt_inventory(pdf_path: Path) -> dict:
    validation_error = validate_pdf_file(pdf_path)
    if validation_error:
        return unavailable_pdf_inventory(validation_error)

    rows: list[dict] = []
    try:
        with fitz.open(str(pdf_path)) as document:
            struct_root_xref = get_struct_tree_root_xref(document)
            if not isinstance(struct_root_xref, int):
                return unavailable_pdf_inventory("This PDF does not have a tagged structure tree.")

            roles = read_role_map(document, struct_root_xref)
            page_lookup = page_xref_map(document)
            struct_xrefs = reachable_struct_elements(document, struct_root_xref)
            marked_content_bboxes_by_page: dict[int, dict[int, list[float]]] = {}
            marked_content_figure_components_by_page: dict[int, dict[int, dict[str, list[tuple[float, float, float, float]]]]] = {}

            for struct_xref in struct_xrefs:
                try:
                    role_type, role_value = document.xref_get_key(struct_xref, "S")
                except Exception:
                    continue
                if role_type != "name":
                    continue
                raw_role = pdf_name_key(role_value)
                standard_role = resolve_pdf_role(raw_role, roles)
                if standard_role not in TARGET_STANDARD_ROLES:
                    continue
                if has_descendant_with_standard_role(
                    document,
                    struct_xref,
                    roles,
                    standard_role,
                ):
                    continue

                page_xref = structure_page_xref(document, struct_xref)
                page_index = page_lookup.get(page_xref) if isinstance(page_xref, int) else None
                preview_bbox = None
                content_bbox = None
                raw_content_bbox = None
                pdf_bbox = structure_bbox(document, struct_xref)
                existing_alt = pdf_alt_text(document, struct_xref)
                is_formula = standard_role == "formula"
                if isinstance(page_index, int):
                    page = document.load_page(page_index)
                    if is_formula:
                        if page_index not in marked_content_bboxes_by_page:
                            marked_content_bboxes_by_page[page_index] = page_marked_content_bboxes(document, page_index)
                        marked_content_bboxes = marked_content_bboxes_by_page[page_index]
                        raw_content_bbox = structure_marked_content_bbox(
                            document,
                            struct_xref,
                            page_index,
                            marked_content_bboxes,
                            roles,
                        )
                        content_bbox = raw_content_bbox
                    else:
                        if page_index not in marked_content_figure_components_by_page:
                            marked_content_figure_components_by_page[page_index] = page_marked_content_components(
                                document,
                                page_index,
                                include_paint=True,
                            )
                        figure_components = structure_marked_content_components(
                            document,
                            struct_xref,
                            marked_content_figure_components_by_page[page_index],
                            roles,
                        )
                        raw_content_bbox = union_content_bboxes(
                            figure_components.get("image", []) + figure_components.get("paint", []) + figure_components.get("text", [])
                        )
                        content_bbox = refine_figure_content_bbox(page, pdf_bbox, figure_components)
                    if is_formula and content_bbox is not None:
                        content_bbox = refine_formula_content_bbox(page, content_bbox, existing_alt)
                        content_bbox = tighten_formula_bbox_to_declared_wrapper_span(page, content_bbox, existing_alt)
                        content_bbox = expand_formula_bbox_for_wrapping_glyphs(page, content_bbox, existing_alt)
                        content_bbox = expand_formula_bbox_for_fine_math_glyphs(page, content_bbox, existing_alt)
                    elif is_formula and existing_alt:
                        content_bbox = locate_formula_bbox_by_page_search(page, existing_alt)
                        if content_bbox is not None:
                            content_bbox = tighten_formula_bbox_to_declared_wrapper_span(page, content_bbox, existing_alt)
                            content_bbox = expand_formula_bbox_for_wrapping_glyphs(page, content_bbox, existing_alt)
                            content_bbox = expand_formula_bbox_for_fine_math_glyphs(page, content_bbox, existing_alt)

                    content_preview_bbox = fitz_bbox_to_preview_norm(page, content_bbox)
                    pdf_preview_bbox = pdf_bbox_to_preview_norm(page, pdf_bbox)
                    if is_formula:
                        preview_bbox = content_preview_bbox or pdf_preview_bbox
                    elif should_prefer_content_bbox_for_figure(page, pdf_bbox, content_bbox):
                        preview_bbox = fitz_bbox_to_preview_norm(page, content_bbox)
                    else:
                        preview_bbox = pdf_preview_bbox or content_preview_bbox

                row = {
                    "id": len(rows),
                    "type": "Formula" if is_formula else "Figure",
                    "role": "equation" if is_formula else "image",
                    "source_part": "PDF structure tree",
                    "page": page_index + 1 if isinstance(page_index, int) else None,
                    "preview_page": page_index,
                    "preview_bbox": preview_bbox,
                    "preview_text": "",
                    "has_alt_text": bool(existing_alt),
                    "alt_text": existing_alt,
                    "existing_alt_text": existing_alt,
                    "generated_alt_text": "",
                    "effective_alt_text": existing_alt,
                    "alt_source": "existing" if existing_alt else "missing",
                    "label": existing_alt or ("Formula" if is_formula else "Figure"),
                    "struct_xref": struct_xref,
                    "struct_role": raw_role,
                    "standard_role": standard_role,
                    "pdf_bbox": pdf_bbox,
                    "content_bbox": content_bbox,
                    "raw_content_bbox": raw_content_bbox,
                }
                rows.append(row)
    except Exception as exc:
        return unavailable_pdf_inventory(f"PDF ALT inventory could not be generated: {exc}")

    summary = summarize_alt_rows(rows)
    return {
        "available": True,
        "message": f"Collected {summary['total_items']} tagged PDF ALT item(s).",
        "rows": rows,
        "summary": summary,
    }


def unavailable_pdf_inventory(message: str) -> dict:
    return {
        "available": False,
        "message": message,
        "rows": [],
        "summary": {
            "total_items": 0,
            "images": 0,
            "equations": 0,
            "with_alt_text": 0,
            "generated_alt_text": 0,
            "with_effective_alt_text": 0,
            "without_alt_text": 0,
        },
    }


def build_pdf_preview_images(rows: list[dict], pdf_path: Path) -> dict[int, dict]:
    return build_alt_preview_images(rows, None, pdf_path)


def inject_pdf_alt_texts(pdf_path: Path, rows: list[dict]) -> tuple[bytes, int]:
    applied_count = 0
    temp_path = pdf_path.with_name(f"{pdf_path.stem}.altomizer_update_{uuid.uuid4().hex}.pdf")
    temp_path.write_bytes(pdf_path.read_bytes())

    try:
        with fitz.open(str(temp_path)) as document:
            for row in rows:
                struct_xref = row.get("struct_xref")
                if not isinstance(struct_xref, int) or struct_xref <= 0 or struct_xref >= document.xref_length():
                    continue
                alt_text = normalize_alt_text(str(row.get("alt_text", "") or ""))
                existing_alt = pdf_alt_text(document, struct_xref)
                if alt_text == existing_alt:
                    continue
                document.xref_set_key(struct_xref, "Alt", fitz.get_pdf_str(alt_text))
                applied_count += 1

            if applied_count:
                document.saveIncr()

        return temp_path.read_bytes(), applied_count
    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass
