import hashlib
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path

import fitz


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}

REL_NS = {"pr": "http://schemas.openxmlformats.org/package/2006/relationships"}


@dataclass
class DocElement:
    index: int
    text: str
    role: str
    style: str | None
    num_id: str | None
    ilvl: str | None
    fingerprint: str
    source_part: str = "body"
    has_alt_text: bool | None = None
    first_row: bool | None = None
    first_column: bool | None = None
    repeat_header: bool | None = None
    allow_row_break_across_pages: bool | None = None
    metadata: dict | None = None


def normalize_text(text: str) -> str:
    text = (text or "").lower()
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^a-z0-9 ]", "", text)
    return text.strip()


def similarity(a: str, b: str) -> float:
    from difflib import SequenceMatcher

    return SequenceMatcher(None, normalize_text(a), normalize_text(b)).ratio()


def exact_text_match(a: str, b: str) -> bool:
    left = normalize_text(a)
    right = normalize_text(b)
    if not left or not right:
        return False
    if left == right:
        return True
    shorter, longer = (left, right) if len(left) <= len(right) else (right, left)
    return len(shorter) >= 10 and shorter in longer


def read_docx_xml(docx_path: str, xml_path: str = "word/document.xml") -> ET.Element:
    with zipfile.ZipFile(docx_path, "r") as archive:
        xml_data = archive.read(xml_path)
    return ET.fromstring(xml_data)


def read_docx_part_names(docx_path: str, prefix: str) -> list[str]:
    with zipfile.ZipFile(docx_path, "r") as archive:
        return sorted(name for name in archive.namelist() if name.startswith(prefix))


def resolve_target(base_part: str, target: str) -> str:
    if not target:
        return ""
    if target.startswith("/"):
        return target.lstrip("/")
    base_dir = Path(base_part).parent
    return str((base_dir / target).as_posix())


def inspect_docx_package(docx_path: str) -> dict:
    relationships = {}
    equation_targets = set()
    media_targets = set()

    with zipfile.ZipFile(docx_path, "r") as archive:
        names = set(archive.namelist())

        rels_name = "word/_rels/document.xml.rels"
        if rels_name in names:
            rels_root = ET.fromstring(archive.read(rels_name))
            for rel in rels_root.findall("pr:Relationship", REL_NS):
                rel_id = rel.attrib.get("Id")
                target = resolve_target("word/document.xml", rel.attrib.get("Target", ""))
                rel_type = rel.attrib.get("Type", "")
                if rel_id and target:
                    relationships[rel_id] = {"target": target, "type": rel_type}

        for name in names:
            lower_name = name.lower()
            if lower_name.startswith("word/media/"):
                media_targets.add(name)
            if not lower_name.startswith("word/embeddings/"):
                continue
            try:
                sample = archive.read(name)[:8192]
            except KeyError:
                continue
            decoded = sample.decode("latin1", errors="ignore").lower()
            if any(marker in decoded for marker in ("equation.dsmt", "mathtype", "microsoft equation 3.0", "equation editor")):
                equation_targets.add(name)

    return {
        "relationships": relationships,
        "equation_targets": equation_targets,
        "media_targets": media_targets,
    }


def read_numbering_definition(docx_path: str) -> dict:
    definition = {"abstract_levels": {}, "num_map": {}, "overrides": {}}

    with zipfile.ZipFile(docx_path, "r") as archive:
        try:
            xml_data = archive.read("word/numbering.xml")
        except KeyError:
            return definition

    root = ET.fromstring(xml_data)

    for abstract_num in root.findall(".//w:abstractNum", NS):
        abstract_id = abstract_num.attrib.get(f"{{{NS['w']}}}abstractNumId")
        if abstract_id is None:
            continue
        levels = {}
        for lvl in abstract_num.findall("./w:lvl", NS):
            ilvl = lvl.attrib.get(f"{{{NS['w']}}}ilvl")
            if ilvl is None:
                continue
            start = lvl.find("./w:start", NS)
            num_fmt = lvl.find("./w:numFmt", NS)
            lvl_text = lvl.find("./w:lvlText", NS)
            levels[ilvl] = {
                "start": int(start.attrib.get(f"{{{NS['w']}}}val", "1")) if start is not None else 1,
                "num_fmt": num_fmt.attrib.get(f"{{{NS['w']}}}val", "decimal") if num_fmt is not None else "decimal",
                "lvl_text": lvl_text.attrib.get(f"{{{NS['w']}}}val", f"%{int(ilvl) + 1}.") if lvl_text is not None else f"%{int(ilvl) + 1}.",
            }
        definition["abstract_levels"][abstract_id] = levels

    for num in root.findall(".//w:num", NS):
        num_id = num.attrib.get(f"{{{NS['w']}}}numId")
        abstract_ref = num.find("./w:abstractNumId", NS)
        if num_id is None or abstract_ref is None:
            continue
        abstract_id = abstract_ref.attrib.get(f"{{{NS['w']}}}val")
        if abstract_id:
            definition["num_map"][num_id] = abstract_id

        overrides = {}
        for lvl_override in num.findall("./w:lvlOverride", NS):
            ilvl = lvl_override.attrib.get(f"{{{NS['w']}}}ilvl")
            if ilvl is None:
                continue
            start_override = lvl_override.find("./w:startOverride", NS)
            if start_override is not None:
                overrides[ilvl] = int(start_override.attrib.get(f"{{{NS['w']}}}val", "1"))
        if overrides:
            definition["overrides"][num_id] = overrides

    return definition


def read_styles_definition(docx_path: str) -> dict:
    definition = {}

    with zipfile.ZipFile(docx_path, "r") as archive:
        try:
            xml_data = archive.read("word/styles.xml")
        except KeyError:
            return definition

    root = ET.fromstring(xml_data)
    for style in root.findall("./w:style", NS):
        if style.attrib.get(f"{{{NS['w']}}}type") != "paragraph":
            continue

        style_id = style.attrib.get(f"{{{NS['w']}}}styleId")
        if not style_id:
            continue

        name_node = style.find("./w:name", NS)
        based_on_node = style.find("./w:basedOn", NS)
        outline_node = style.find("./w:pPr/w:outlineLvl", NS)

        outline_level = None
        if outline_node is not None:
            try:
                outline_level = int(outline_node.attrib.get(f"{{{NS['w']}}}val", ""))
            except (TypeError, ValueError):
                outline_level = None

        definition[style_id] = {
            "name": name_node.attrib.get(f"{{{NS['w']}}}val", "") if name_node is not None else "",
            "based_on": based_on_node.attrib.get(f"{{{NS['w']}}}val") if based_on_node is not None else None,
            "outline_level": outline_level,
            "quick_style": style.find("./w:qFormat", NS) is not None,
        }

    return definition


def resolve_style_outline_level(style_name: str | None, styles_definition: dict, seen: set[str] | None = None) -> int | None:
    if not style_name:
        return None

    if seen is None:
        seen = set()
    if style_name in seen:
        return None
    seen.add(style_name)

    style_entry = styles_definition.get(style_name)
    if not style_entry:
        return None

    outline_level = style_entry.get("outline_level")
    if outline_level is not None:
        return outline_level

    based_on = style_entry.get("based_on")
    if based_on:
        return resolve_style_outline_level(based_on, styles_definition, seen)

    return None


def twips_to_pt(value: str | None) -> float | None:
    if not value:
        return None
    try:
        return float(value) / 20.0
    except (TypeError, ValueError):
        return None


def emu_to_pt(value: str | None) -> float | None:
    if not value:
        return None
    try:
        return float(value) / 12700.0
    except (TypeError, ValueError):
        return None


def extract_style_dimension_pt(style_value: str | None, name: str) -> float | None:
    if not style_value:
        return None
    match = re.search(rf"{name}\s*:\s*([0-9]+(?:\.[0-9]+)?)pt", style_value, re.IGNORECASE)
    if not match:
        return None
    return float(match.group(1))


def shape_dimensions_pt(shape: ET.Element | None) -> tuple[float | None, float | None]:
    if shape is None:
        return None, None
    style = shape.attrib.get("style", "")
    return (
        extract_style_dimension_pt(style, "width"),
        extract_style_dimension_pt(style, "height"),
    )


def clamp_crop_fraction(value: float | None) -> float:
    if value is None:
        return 0.0
    return max(0.0, min(float(value), 0.99))


def parse_drawingml_crop_value(value: str | None) -> float | None:
    if value is None:
        return None
    raw = str(value).strip()
    if not raw:
        return None
    try:
        return float(raw) / 100000.0
    except (TypeError, ValueError):
        return None


def parse_vml_crop_value(value: str | None) -> float | None:
    if value is None:
        return None
    raw = str(value).strip().lower()
    if not raw:
        return None
    try:
        if raw.endswith("%"):
            return float(raw[:-1]) / 100.0
        if raw.endswith("f"):
            return float(raw[:-1]) / 65536.0
        return float(raw)
    except (TypeError, ValueError):
        return None


def normalize_crop_rect(left: float | None, top: float | None, right: float | None, bottom: float | None) -> dict | None:
    crop = {
        "left": clamp_crop_fraction(left),
        "top": clamp_crop_fraction(top),
        "right": clamp_crop_fraction(right),
        "bottom": clamp_crop_fraction(bottom),
    }
    if all(value <= 0 for value in crop.values()):
        return None
    if crop["left"] + crop["right"] >= 0.98:
        return None
    if crop["top"] + crop["bottom"] >= 0.98:
        return None
    return crop


def extract_drawing_crop(node: ET.Element) -> dict | None:
    src_rect = node.find(".//a:srcRect", NS)
    if src_rect is None:
        return None
    return normalize_crop_rect(
        parse_drawingml_crop_value(src_rect.attrib.get("l")),
        parse_drawingml_crop_value(src_rect.attrib.get("t")),
        parse_drawingml_crop_value(src_rect.attrib.get("r")),
        parse_drawingml_crop_value(src_rect.attrib.get("b")),
    )


def extract_vml_crop(image: ET.Element | None, shape: ET.Element | None = None) -> dict | None:
    if image is None and shape is None:
        return None
    source = image if image is not None else shape
    if source is None:
        return None
    return normalize_crop_rect(
        parse_vml_crop_value(source.attrib.get("cropleft")),
        parse_vml_crop_value(source.attrib.get("croptop")),
        parse_vml_crop_value(source.attrib.get("cropright")),
        parse_vml_crop_value(source.attrib.get("cropbottom")),
    )


def visual_identity_key(kind: str, metadata: dict | None) -> tuple[str, tuple[tuple[str, str], ...]] | None:
    info = metadata or {}
    fields: list[tuple[str, str]] = []

    target_values = [
        str(info.get(name)).strip()
        for name in ("target", "image_target", "ole_target")
        if info.get(name) is not None and str(info.get(name)).strip()
    ]
    rel_values = [
        str(info.get(name)).strip()
        for name in ("rel_id", "image_rid", "ole_rid")
        if info.get(name) is not None and str(info.get(name)).strip()
    ]
    shape_values = [
        str(info.get(name)).strip()
        for name in ("docpr_id", "shape_id")
        if info.get(name) is not None and str(info.get(name)).strip()
    ]

    if target_values:
        fields.append(("target", "|".join(sorted(set(target_values)))))
    if rel_values:
        fields.append(("rel", "|".join(sorted(set(rel_values)))))
    if shape_values:
        fields.append(("shape", "|".join(sorted(set(shape_values)))))

    if not fields:
        return None
    return (kind, tuple(fields))


def dedupe_visual_entries(entries: list[dict]) -> list[dict]:
    deduped: list[dict] = []
    seen: set[tuple[str, tuple[tuple[str, str], ...]]] = set()
    for entry in entries:
        kind = str(entry.get("kind", "") or entry.get("role", "") or "visual")
        key = visual_identity_key(kind, entry.get("metadata"))
        if key is not None:
            if key in seen:
                continue
            seen.add(key)
        deduped.append(entry)
    return deduped


def drawing_dimensions_pt(drawing: ET.Element) -> tuple[float | None, float | None]:
    extent = drawing.find(".//wp:extent", NS)
    if extent is not None:
        width_pt = emu_to_pt(extent.attrib.get("cx"))
        height_pt = emu_to_pt(extent.attrib.get("cy"))
        if width_pt is not None and height_pt is not None:
            return width_pt, height_pt
    shape = drawing.find(".//v:shape", NS)
    return shape_dimensions_pt(shape)


def size_class_from_dimensions(width_pt: float | None, height_pt: float | None) -> str | None:
    if width_pt is None or height_pt is None:
        return None
    area = width_pt * height_pt
    longest_edge = max(width_pt, height_pt)
    if area <= 1800 and longest_edge <= 72:
        return "small"
    if area <= 12000 and longest_edge <= 180:
        return "medium"
    return "large"


def size_class_from_bbox(bbox: dict | None) -> str | None:
    if not bbox:
        return None
    width = max(0.0, bbox.get("x1", 0.0) - bbox.get("x0", 0.0))
    height = max(0.0, bbox.get("y1", 0.0) - bbox.get("y0", 0.0))
    area = width * height
    longest_edge = max(width, height)
    if area <= 0.006 and longest_edge <= 0.18:
        return "small"
    if area <= 0.04 and longest_edge <= 0.38:
        return "medium"
    return "large"


def bbox_metrics(bbox: dict | None) -> tuple[float, float, float]:
    if not bbox:
        return 0.0, 0.0, 0.0
    width = max(0.0, bbox.get("x1", 0.0) - bbox.get("x0", 0.0))
    height = max(0.0, bbox.get("y1", 0.0) - bbox.get("y0", 0.0))
    return width, height, width * height


def expand_bbox(bbox: dict, margin_x: float = 0.0, margin_y: float | None = None) -> dict:
    y_margin = margin_x if margin_y is None else margin_y
    return {
        "x0": max(0.0, bbox["x0"] - margin_x),
        "y0": max(0.0, bbox["y0"] - y_margin),
        "x1": min(1.0, bbox["x1"] + margin_x),
        "y1": min(1.0, bbox["y1"] + y_margin),
    }


def bboxes_intersect(left: dict, right: dict) -> bool:
    return not (
        left["x1"] < right["x0"]
        or right["x1"] < left["x0"]
        or left["y1"] < right["y0"]
        or right["y1"] < left["y0"]
    )


def bbox_contains_point(bbox: dict | None, x: float, y: float, margin: float = 0.0) -> bool:
    if not bbox:
        return False
    expanded = expand_bbox(bbox, margin)
    return expanded["x0"] <= x <= expanded["x1"] and expanded["y0"] <= y <= expanded["y1"]


def bbox_horizontal_overlap_ratio(left: dict | None, right: dict | None) -> float:
    if not left or not right:
        return 0.0
    overlap = min(left["x1"], right["x1"]) - max(left["x0"], right["x0"])
    if overlap <= 0:
        return 0.0
    width = min(left["x1"] - left["x0"], right["x1"] - right["x0"])
    if width <= 0:
        return 0.0
    return overlap / width


def aspect_ratio_from_bbox(bbox: dict | None) -> float | None:
    width, height, _ = bbox_metrics(bbox)
    if width <= 0 or height <= 0:
        return None
    return width / height


def aspect_ratio_similarity(expected: float | None, candidate: float | None) -> float:
    if expected is None or candidate is None or expected <= 0 or candidate <= 0:
        return 0.0
    difference = abs(expected - candidate) / max(expected, candidate)
    return max(0.0, 1.2 - (difference * 2.4))


def get_text(paragraph: ET.Element) -> str:
    text = paragraph_visible_text(paragraph)
    return re.sub(r"\s+", " ", text).strip()


def get_style(paragraph: ET.Element) -> str | None:
    style = paragraph.find("./w:pPr/w:pStyle", NS)
    if style is not None:
        return style.attrib.get(f"{{{NS['w']}}}val")
    return None


def run_text(run: ET.Element) -> str:
    parts = []
    for node in run.findall(".//w:t", NS):
        if node.text:
            parts.append(node.text)
    return "".join(parts)


def run_font_family(run: ET.Element) -> str | None:
    fonts = run.find("./w:rPr/w:rFonts", NS)
    if fonts is None:
        return None
    for attribute in ("ascii", "hAnsi", "cs", "eastAsia"):
        value = fonts.attrib.get(f"{{{NS['w']}}}{attribute}")
        if value:
            return value.strip()
    return None


def run_font_size_pt(run: ET.Element) -> float | None:
    size = run.find("./w:rPr/w:sz", NS)
    if size is None:
        size = run.find("./w:rPr/w:szCs", NS)
    if size is None:
        return None
    try:
        return float(size.attrib.get(f"{{{NS['w']}}}val", "")) / 2.0
    except (TypeError, ValueError):
        return None


def extract_paragraph_text_metadata(paragraph: ET.Element) -> dict:
    font_families = []
    font_sizes = []
    seen_fonts = set()
    seen_sizes = set()

    for run in paragraph.findall(".//w:r", NS):
        if not run_text(run).strip():
            continue

        font_family = run_font_family(run)
        if font_family:
            font_key = font_family.casefold()
            if font_key not in seen_fonts:
                seen_fonts.add(font_key)
                font_families.append(font_family)

        font_size = run_font_size_pt(run)
        if font_size is not None:
            size_key = round(font_size, 2)
            if size_key not in seen_sizes:
                seen_sizes.add(size_key)
                font_sizes.append(size_key)

    return {
        "font_family": font_families[0] if len(font_families) == 1 else None,
        "font_families": font_families,
        "font_size_pt": font_sizes[0] if len(font_sizes) == 1 else None,
        "font_sizes_pt": font_sizes,
    }


def get_numbering(paragraph: ET.Element):
    num_id = paragraph.find("./w:pPr/w:numPr/w:numId", NS)
    ilvl = paragraph.find("./w:pPr/w:numPr/w:ilvl", NS)
    num_id_val = num_id.attrib.get(f"{{{NS['w']}}}val") if num_id is not None else None
    ilvl_val = ilvl.attrib.get(f"{{{NS['w']}}}val") if ilvl is not None else None
    return num_id_val, ilvl_val


def format_list_number(value: int, num_fmt: str) -> str:
    if num_fmt == "decimal":
        return str(value)
    if num_fmt == "lowerLetter":
        result = ""
        current = value
        while current > 0:
            current -= 1
            result = chr(ord("a") + (current % 26)) + result
            current //= 26
        return result or str(value)
    if num_fmt == "upperLetter":
        return format_list_number(value, "lowerLetter").upper()
    if num_fmt == "lowerRoman":
        return int_to_roman(value).lower()
    if num_fmt == "upperRoman":
        return int_to_roman(value)
    return str(value)


def int_to_roman(value: int) -> str:
    numerals = [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ]
    current = max(1, value)
    parts = []
    for amount, symbol in numerals:
        while current >= amount:
            parts.append(symbol)
            current -= amount
    return "".join(parts)


def numbering_level_definition(numbering_definition: dict, num_id: str | None, ilvl: str | None) -> dict | None:
    if num_id is None:
        return None
    abstract_id = numbering_definition.get("num_map", {}).get(num_id)
    if abstract_id is None:
        return None
    level_key = ilvl or "0"
    return numbering_definition.get("abstract_levels", {}).get(abstract_id, {}).get(level_key)


def apply_numbering_prefix(
    text: str,
    num_id: str | None,
    ilvl: str | None,
    numbering_definition: dict,
    numbering_state: dict,
) -> str:
    if num_id is None:
        return text

    level = int(ilvl or "0")
    level_definition = numbering_level_definition(numbering_definition, num_id, ilvl)
    if level_definition is None:
        if text_looks_like_list_item(text):
            return text
        return f"{level + 1}. {text}".strip()

    counters = numbering_state.setdefault(num_id, {})
    for existing_level in list(counters):
        if existing_level > level:
            del counters[existing_level]

    if level not in counters:
        start_value = numbering_definition.get("overrides", {}).get(num_id, {}).get(str(level), level_definition.get("start", 1))
        counters[level] = start_value
    else:
        counters[level] += 1

    def counter_value(level_number: int) -> int:
        if level_number in counters:
            return counters[level_number]
        fallback_definition = numbering_level_definition(numbering_definition, num_id, str(level_number)) or {"start": 1}
        counters[level_number] = fallback_definition.get("start", 1)
        return counters[level_number]

    lvl_text = level_definition.get("lvl_text", f"%{level + 1}.")

    def replace_token(match: re.Match) -> str:
        token_level = max(0, int(match.group(1)) - 1)
        token_definition = numbering_level_definition(numbering_definition, num_id, str(token_level)) or {"num_fmt": "decimal"}
        return format_list_number(counter_value(token_level), token_definition.get("num_fmt", "decimal"))

    prefix = re.sub(r"%(\d+)", replace_token, lvl_text).strip()
    if not prefix:
        return text
    if text_looks_like_list_item(text) and text.lstrip().startswith(prefix):
        return text
    return f"{prefix} {text}".strip()


def get_table_style(table: ET.Element) -> str | None:
    style = table.find("./w:tblPr/w:tblStyle", NS)
    if style is not None:
        return style.attrib.get(f"{{{NS['w']}}}val")
    return None


def table_row_has_header_repeat(row: ET.Element) -> bool:
    tbl_header = row.find("./w:trPr/w:tblHeader", NS)
    if tbl_header is None:
        return False
    value = tbl_header.attrib.get(f"{{{NS['w']}}}val")
    return value not in {"0", "false", "off"}


def table_row_allows_break_across_pages(row: ET.Element) -> bool:
    cant_split = row.find("./w:trPr/w:cantSplit", NS)
    if cant_split is None:
        return True
    value = cant_split.attrib.get(f"{{{NS['w']}}}val")
    return value in {"0", "false", "off"}


def table_flags(table: ET.Element):
    look = table.find("./w:tblPr/w:tblLook", NS)
    first_row = False
    first_column = False
    if look is not None:
        first_row = look.attrib.get(f"{{{NS['w']}}}firstRow") == "1"
        first_column = look.attrib.get(f"{{{NS['w']}}}firstColumn") == "1"

    rows = table.findall("./w:tr", NS)
    repeat_header = table_row_has_header_repeat(rows[0]) if rows else False
    allow_break_across_pages = any(table_row_allows_break_across_pages(row) for row in rows) if rows else True
    return first_row, first_column, repeat_header, allow_break_across_pages


def paragraph_has_math(paragraph: ET.Element) -> bool:
    return paragraph.find(".//m:oMath", NS) is not None or paragraph.find(".//m:oMathPara", NS) is not None


def paragraph_instruction_text(paragraph: ET.Element) -> str:
    parts = []
    for node in paragraph.findall(".//w:instrText", NS):
        if node.text:
            parts.append(node.text)
    for field in paragraph.findall(".//w:fldSimple", NS):
        instruction = field.attrib.get(f"{{{NS['w']}}}instr")
        if instruction:
            parts.append(instruction)
    return " ".join(parts).strip()


def paragraph_outline_level(paragraph: ET.Element, styles_definition: dict | None = None) -> int | None:
    direct_outline = paragraph.find("./w:pPr/w:outlineLvl", NS)
    if direct_outline is not None:
        try:
            return int(direct_outline.attrib.get(f"{{{NS['w']}}}val", ""))
        except (TypeError, ValueError):
            pass

    return resolve_style_outline_level(get_style(paragraph), styles_definition or {})


def effective_node_children(node: ET.Element) -> list[ET.Element]:
    if node.tag == f"{{{NS['mc']}}}AlternateContent":
        choice = node.find("./mc:Choice", NS)
        if choice is not None and list(choice):
            return list(choice)
        fallback = node.find("./mc:Fallback", NS)
        if fallback is not None:
            return list(fallback)
        return []
    return list(node)


def office_math_fragments(node: ET.Element) -> list[dict]:
    if node.tag == f"{{{NS['m']}}}oMathPara":
        direct_objects = node.findall("./m:oMath", NS)
        if direct_objects:
            return [
                {
                    "kind": "equation",
                    "text": office_math_text(math_object) or "Equation",
                    "metadata": {"math_source": "omml"},
                }
                for math_object in direct_objects
            ]
        return [
            {
                "kind": "equation",
                "text": office_math_text(node) or "Equation",
                "metadata": {"math_source": "omml"},
            }
        ]

    return [
        {
            "kind": "equation",
            "text": office_math_text(node) or "Equation",
            "metadata": {"math_source": "omml"},
        }
    ]


def extract_mathtype_object_from_node(node: ET.Element, package_info: dict | None = None) -> dict | None:
    if node.tag != f"{{{NS['w']}}}object":
        return None

    relationships = (package_info or {}).get("relationships", {})
    ole = node.find(".//o:OLEObject", NS)
    if ole is None:
        return None

    prog_id = (ole.attrib.get("ProgID", "") or "").strip()
    prog_id_lower = prog_id.lower()
    if not any(keyword in prog_id_lower for keyword in ("equation", "mathtype", "math")):
        return None

    shape = node.find(".//v:shape", NS)
    image = node.find(".//v:imagedata", NS)
    image_rid = image.attrib.get(f"{{{NS['r']}}}id") if image is not None else None
    ole_rid = ole.attrib.get(f"{{{NS['r']}}}id")

    alt_text = ""
    if shape is not None:
        alt_text = (shape.attrib.get("alt", "") or shape.attrib.get("title", "") or "").strip()
    if not alt_text and image is not None:
        alt_text = (image.attrib.get("title", "") or "").strip()

    image_target = relationships.get(image_rid, {}).get("target") if image_rid else None
    ole_target = relationships.get(ole_rid, {}).get("target") if ole_rid else None
    width_pt, height_pt = shape_dimensions_pt(shape)
    if width_pt is None:
        width_pt = twips_to_pt(node.attrib.get(f"{{{NS['w']}}}dxaOrig"))
    if height_pt is None:
        height_pt = twips_to_pt(node.attrib.get(f"{{{NS['w']}}}dyaOrig"))
    crop = extract_vml_crop(image, shape)
    shape_id = shape.attrib.get("id", "") if shape is not None else ""

    return {
        "alt_text": alt_text,
        "has_alt_text": bool(alt_text),
        "prog_id": prog_id,
        "image_rid": image_rid,
        "ole_rid": ole_rid,
        "image_target": image_target,
        "ole_target": ole_target,
        "width_pt": width_pt,
        "height_pt": height_pt,
        "crop": crop,
        "shape_id": shape_id,
    }


def extract_generic_drawing_object_from_node(node: ET.Element, package_info: dict | None = None) -> dict | None:
    relationships = (package_info or {}).get("relationships", {})

    if node.tag == f"{{{NS['w']}}}drawing":
        doc_pr = node.find(".//wp:docPr", NS)
        title = doc_pr.attrib.get("title", "") if doc_pr is not None else ""
        descr = doc_pr.attrib.get("descr", "") if doc_pr is not None else ""
        alt_text = (descr or title).strip()
        blip = node.find(".//a:blip", NS)
        rel_id = None
        if blip is not None:
            rel_id = blip.attrib.get(f"{{{NS['r']}}}embed") or blip.attrib.get(f"{{{NS['r']}}}link")
        rel_info = relationships.get(rel_id, {}) if rel_id else {}
        target = rel_info.get("target") if rel_id else None
        width_pt, height_pt = drawing_dimensions_pt(node)
        crop = extract_drawing_crop(node)
        if not rel_id and not target and not alt_text:
            return None
        return {
            "alt_text": alt_text,
            "has_alt_text": bool(alt_text),
            "is_equation": generic_drawing_is_equation(
                alt_text,
                rel_info,
                package_info or {},
                width_pt,
                height_pt,
            ),
            "metadata": {
                "rel_id": rel_id,
                "target": target,
                "width_pt": width_pt,
                "height_pt": height_pt,
                "crop": crop,
                "docpr_id": doc_pr.attrib.get("id", "") if doc_pr is not None else "",
                "docpr_name": doc_pr.attrib.get("name", "") if doc_pr is not None else "",
            },
        }

    if node.tag != f"{{{NS['w']}}}pict":
        return None
    if node.find(".//o:OLEObject", NS) is not None:
        return None

    shape = node.find(".//v:shape", NS)
    image = node.find(".//v:imagedata", NS)
    alt_text = ""
    if shape is not None:
        alt_text = (shape.attrib.get("alt", "") or shape.attrib.get("title", "") or "").strip()
    if not alt_text and image is not None:
        alt_text = (image.attrib.get("title", "") or "").strip()
    rel_id = image.attrib.get(f"{{{NS['r']}}}id") if image is not None else None
    rel_info = relationships.get(rel_id, {}) if rel_id else {}
    target = rel_info.get("target") if rel_id else None
    width_pt, height_pt = shape_dimensions_pt(shape)
    crop = extract_vml_crop(image, shape)
    if not rel_id and not target and not alt_text:
        return None
    return {
        "alt_text": alt_text,
        "has_alt_text": bool(alt_text),
        "is_equation": generic_drawing_is_equation(
            alt_text,
            rel_info,
            package_info or {},
            width_pt,
            height_pt,
        ),
        "metadata": {
            "rel_id": rel_id,
            "target": target,
            "width_pt": width_pt,
            "height_pt": height_pt,
            "crop": crop,
            "shape_id": shape.attrib.get("id", "") if shape is not None else "",
        },
    }


def extract_inline_fragments(node: ET.Element, package_info: dict | None = None) -> list[dict]:
    text_substitutions = {
        f"{{{NS['w']}}}tab": "\t",
        f"{{{NS['w']}}}br": "\n",
        f"{{{NS['w']}}}cr": "\n",
        f"{{{NS['w']}}}softHyphen": "-",
        f"{{{NS['w']}}}noBreakHyphen": "-",
    }

    if node.tag == f"{{{NS['w']}}}t":
        return [{"kind": "text", "text": node.text or ""}]
    if node.tag in text_substitutions:
        return [{"kind": "text", "text": text_substitutions[node.tag]}]
    if node.tag in {f"{{{NS['w']}}}instrText", f"{{{NS['w']}}}delText"}:
        return []
    if node.tag in {f"{{{NS['m']}}}oMath", f"{{{NS['m']}}}oMathPara"}:
        return office_math_fragments(node)

    math_object = extract_mathtype_object_from_node(node, package_info)
    if math_object:
        return [
            {
                "kind": "equation",
                "text": math_object["alt_text"] or "Equation",
                "has_alt_text": math_object["has_alt_text"],
                "metadata": {
                    "prog_id": math_object["prog_id"],
                    "image_rid": math_object["image_rid"],
                    "ole_rid": math_object["ole_rid"],
                    "image_target": math_object["image_target"],
                    "ole_target": math_object["ole_target"],
                    "width_pt": math_object["width_pt"],
                    "height_pt": math_object["height_pt"],
                    "crop": math_object.get("crop"),
                    "shape_id": math_object.get("shape_id"),
                },
            }
        ]

    drawing_object = extract_generic_drawing_object_from_node(node, package_info)
    if drawing_object:
        fragment_kind = "equation" if drawing_object["is_equation"] else "image"
        return [
            {
                "kind": fragment_kind,
                "text": drawing_object["alt_text"] or ("Equation" if fragment_kind == "equation" else "Image"),
                "has_alt_text": drawing_object["has_alt_text"],
                "metadata": drawing_object["metadata"],
            }
        ]

    fragments = []
    for child in effective_node_children(node):
        fragments.extend(extract_inline_fragments(child, package_info))
    return fragments


def merge_inline_text_fragments(fragments: list[dict]) -> list[dict]:
    merged = []
    for fragment in fragments:
        if fragment.get("kind") != "text":
            merged.append(fragment)
            continue
        text = fragment.get("text", "")
        if not text:
            continue
        if merged and merged[-1].get("kind") == "text":
            merged[-1]["text"] = f"{merged[-1].get('text', '')}{text}"
            continue
        merged.append({"kind": "text", "text": text})
    return merged


def extract_paragraph_inline_fragments(paragraph: ET.Element, package_info: dict | None = None) -> list[dict]:
    fragments = []
    for child in effective_node_children(paragraph):
        if child.tag == f"{{{NS['w']}}}pPr":
            continue
        fragments.extend(extract_inline_fragments(child, package_info))
    return merge_inline_text_fragments(dedupe_visual_entries(fragments))


def paragraph_visible_text(paragraph: ET.Element) -> str:
    fragments = extract_paragraph_inline_fragments(paragraph)
    return "".join(fragment.get("text", "") for fragment in fragments if fragment.get("kind") == "text")


def apply_numbering_prefix_to_fragments(
    fragments: list[dict],
    num_id: str | None,
    ilvl: str | None,
    numbering_definition: dict,
    numbering_state: dict,
) -> list[dict]:
    if num_id is None:
        return fragments

    prefix = apply_numbering_prefix("", num_id, ilvl, numbering_definition, numbering_state)
    if not prefix:
        return fragments

    updated = [dict(fragment) for fragment in fragments]
    first_text_index = None
    for index, fragment in enumerate(updated):
        if fragment.get("kind") != "text":
            continue
        first_text_index = index
        text = fragment.get("text", "")
        if text_looks_like_list_item(text):
            if index > 0:
                return [{"kind": "text", "text": prefix}] + updated
            return updated
        fragment["text"] = f"{prefix} {text}".strip()
        return updated

    if first_text_index is not None:
        return updated
    return [{"kind": "text", "text": prefix}] + updated


def paragraph_text_from_fragments(fragments: list[dict]) -> str:
    text = "".join(fragment.get("text", "") for fragment in fragments if fragment.get("kind") == "text")
    return re.sub(r"\s+", " ", text).strip()


def split_text_segments(text: str) -> list[str]:
    return [segment.strip() for segment in re.split(r"[\t\r\n]+", text or "") if segment.strip()]


def fragment_text_segments(text: str, paragraph_role: str) -> list[str]:
    if paragraph_role in {"heading", "title", "subtitle", "caption", "quote"}:
        collapsed = re.sub(r"\s+", " ", text or "").strip()
        return [collapsed] if collapsed else []

    segments = split_text_segments(text)
    filtered = []
    for segment in segments:
        if not re.search(r"[A-Za-z0-9]", segment) and not text_looks_like_list_item(segment):
            continue
        filtered.append(segment)
    return filtered


def shape_alt_texts(paragraph: ET.Element) -> list[str]:
    texts = []
    for shape in paragraph.findall(".//v:shape", NS):
        alt_text = (shape.attrib.get("alt", "") or shape.attrib.get("title", "") or "").strip()
        if alt_text:
            texts.append(alt_text)
    for image in paragraph.findall(".//v:imagedata", NS):
        alt_text = (image.attrib.get("title", "") or "").strip()
        if alt_text:
            texts.append(alt_text)
    return texts


def office_math_text(node: ET.Element) -> str:
    parts = []
    for text_node in node.findall(".//m:t", NS):
        if text_node.text:
            parts.append(text_node.text)
    return "".join(parts).strip()


def extract_office_math_objects(paragraph: ET.Element) -> list[dict]:
    objects = []
    nested_math_ids = set()

    for math_paragraph in paragraph.findall(".//m:oMathPara", NS):
        paragraph_objects = math_paragraph.findall("./m:oMath", NS)
        if paragraph_objects:
            for math_object in paragraph_objects:
                nested_math_ids.add(id(math_object))
                text = office_math_text(math_object)
                objects.append(
                    {
                        "text": text or "Equation",
                        "metadata": {"math_source": "omml"},
                    }
                )
            continue

        text = office_math_text(math_paragraph)
        objects.append(
            {
                "text": text or "Equation",
                "metadata": {"math_source": "omml"},
            }
        )

    for math_object in paragraph.findall(".//m:oMath", NS):
        if id(math_object) in nested_math_ids:
            continue
        text = office_math_text(math_object)
        objects.append(
            {
                "text": text or "Equation",
                "metadata": {"math_source": "omml"},
            }
        )

    return objects


def extract_mathtype_objects(paragraph: ET.Element, package_info: dict | None = None) -> list[dict]:
    objects = []

    for obj in paragraph.findall(".//w:object", NS):
        math_object = extract_mathtype_object_from_node(obj, package_info)
        if math_object:
            objects.append(math_object)

    return dedupe_visual_entries(objects)


def extract_generic_drawing_objects(paragraph: ET.Element, package_info: dict | None = None) -> list[dict]:
    objects = []

    for drawing in paragraph.findall(".//w:drawing", NS):
        drawing_object = extract_generic_drawing_object_from_node(drawing, package_info)
        if drawing_object:
            objects.append(drawing_object)

    for pict in paragraph.findall(".//w:pict", NS):
        drawing_object = extract_generic_drawing_object_from_node(pict, package_info)
        if drawing_object:
            objects.append(drawing_object)

    return dedupe_visual_entries(objects)


def paragraph_relationship_ids(paragraph: ET.Element) -> set[str]:
    rel_ids = set()
    relationship_attributes = [
        f"{{{NS['r']}}}id",
        f"{{{NS['r']}}}embed",
        f"{{{NS['r']}}}link",
    ]
    for node in paragraph.iter():
        for attr_name in relationship_attributes:
            value = node.attrib.get(attr_name)
            if value:
                rel_ids.add(value)
    return rel_ids


def relationship_target_looks_like_equation(rel_info: dict, package_info: dict) -> bool:
    target = rel_info.get("target", "")
    rel_type = (rel_info.get("type", "") or "").lower()
    target_lower = target.lower()
    if target in package_info.get("equation_targets", set()):
        return True
    if any(keyword in rel_type for keyword in ("oleobject", "package")) and target_lower.startswith("word/embeddings/"):
        return True
    if target_lower.startswith("word/embeddings/") and any(keyword in target_lower for keyword in ("equation", "mathtype", "oleobject")):
        return True
    if target_lower.startswith("word/media/") and any(keyword in target_lower for keyword in ("equation", "mathtype", "formula", "math")):
        return True
    return False


def paragraph_has_embedded_equation(paragraph: ET.Element, package_info: dict | None = None, text: str = "", style: str | None = None) -> bool:
    if extract_mathtype_objects(paragraph, package_info):
        return True

    if extract_office_math_objects(paragraph):
        return True

    instruction_text = paragraph_instruction_text(paragraph).lower()
    if re.search(r"\b(eq|equation|mathtype|math|professional|officemath)\b", instruction_text):
        return True

    for ole_object in paragraph.findall(".//o:OLEObject", NS):
        prog_id = (ole_object.attrib.get("ProgID", "") or "").lower()
        if any(keyword in prog_id for keyword in ("equation", "mathtype", "math")):
            return True

    if package_info:
        relationships = package_info.get("relationships", {})
        for rel_id in paragraph_relationship_ids(paragraph):
            rel_info = relationships.get(rel_id)
            if rel_info and relationship_target_looks_like_equation(rel_info, package_info):
                return True

    has_object = paragraph.find(".//w:object", NS) is not None
    has_pict = paragraph.find(".//w:pict", NS) is not None
    style_lower = (style or "").lower()
    if (has_object or has_pict) and any(keyword in style_lower for keyword in ("equation", "formula", "math", "vector")):
        return True

    return False


def paragraph_drawings(
    paragraph: ET.Element,
    force_equation: bool = False,
    package_info: dict | None = None,
) -> list[dict]:
    drawings = []

    math_objects = extract_mathtype_objects(paragraph, package_info)
    for math_object in math_objects:
        drawings.append(
            {
                "alt_text": math_object["alt_text"],
                "has_alt_text": math_object["has_alt_text"],
                "is_equation": True,
                "metadata": {
                    "prog_id": math_object["prog_id"],
                    "image_rid": math_object["image_rid"],
                    "ole_rid": math_object["ole_rid"],
                    "image_target": math_object["image_target"],
                    "ole_target": math_object["ole_target"],
                    "width_pt": math_object["width_pt"],
                    "height_pt": math_object["height_pt"],
                    "crop": math_object.get("crop"),
                    "shape_id": math_object.get("shape_id"),
                },
            }
        )

    generic_objects = extract_generic_drawing_objects(paragraph, package_info)
    for generic_object in generic_objects:
        drawings.append(
            {
                "alt_text": generic_object["alt_text"],
                "has_alt_text": generic_object["has_alt_text"],
                "is_equation": force_equation or generic_object["is_equation"],
                "metadata": generic_object["metadata"],
            }
        )
    return dedupe_visual_entries(drawings)


def drawing_looks_like_equation(alt_text: str | None) -> bool:
    if not alt_text:
        return False
    text = normalize_text(alt_text)
    words = text.split()
    if looks_like_answer_section_heading(text):
        return False
    image_context_words = {
        "graph",
        "axis",
        "axes",
        "grid",
        "curve",
        "curves",
        "rectangle",
        "picture",
        "figure",
        "plot",
        "diagram",
        "plane",
        "hyperbola",
        "ellipse",
        "parabola",
        "triangle",
        "line",
        "notes",
        "handwritten",
        "steps",
        "coordinate",
    }
    if any(word in words for word in image_context_words):
        return False
    if re.search(r"\b(professional|officemath)\b", text):
        return True
    if re.search(r"\b(equation|formula|math|vector)\b", text):
        if len(words) <= 12:
            return True
        return False

    spoken_math_hits = sum(
        phrase in text
        for phrase in (
            "equals",
            "plus",
            "minus",
            "over",
            "squared",
            "cubed",
            "open parenthesis",
            "close parenthesis",
            "fraction",
        )
    )
    word_count = len(words)
    return spoken_math_hits >= 4 and word_count <= 12


def generic_drawing_is_equation(
    alt_text: str | None,
    rel_info: dict,
    package_info: dict,
    width_pt: float | None,
    height_pt: float | None,
) -> bool:
    if relationship_target_looks_like_equation(rel_info, package_info):
        return True

    if not drawing_looks_like_equation(alt_text):
        return False

    size_class = size_class_from_dimensions(width_pt, height_pt)
    if size_class == "large":
        return False

    normalized = normalize_text(alt_text or "")
    if looks_like_answer_section_heading(normalized):
        return False
    if re.search(r"\b(graph|diagram|notes|handwritten|coordinate|plane|hyperbola|ellipse|parabola|axis|grid|curve)\b", normalized):
        return False

    return True


def looks_like_answer_section_heading(text: str) -> bool:
    normalized = normalize_text(text)
    if not normalized:
        return False
    if re.match(r"^\d+(?:[-.]\d+)+\s+", normalized):
        if re.search(r"\b(answers?|chapter|section|exercise|problem|review|solutions?|cont)\b", normalized):
            return True
    if re.search(r"\b(answers?|chapter|section|exercise|problem|review|solutions?|cont)\b", normalized):
        if re.match(r"^\d+(?:[-.]\d+)?\b", normalized):
            return True
    return False


def text_looks_like_equation(text: str) -> bool:
    normalized = (text or "").strip()
    if len(normalized) < 3:
        return False
    if re.match(r"^\d+(?:[-.]\d+)+\s+", normalize_text(normalized)) and re.search(
        r"\b(answers?|chapter|section|exercise|problem|review|solutions?|cont)\b",
        normalized,
        re.IGNORECASE,
    ):
        return False
    word_count = len(re.findall(r"\b[\w']+\b", normalized))
    alpha_words = re.findall(r"[A-Za-z]{2,}", normalized)
    if word_count >= 8 and len(alpha_words) >= 5:
        return False
    if re.search(r"\b(answers?|chapter|section|exercise|problem|review|solutions?)\b", normalized, re.IGNORECASE):
        return False
    operator_score = sum(symbol in normalized for symbol in ("=", "+", "/", "^", "×", "÷", "·", "≤", "≥"))
    grouping_score = sum(symbol in normalized for symbol in ("(", ")", "[", "]"))
    digit_count = sum(char.isdigit() for char in normalized)
    letter_count = sum(char.isalpha() for char in normalized)
    variable_count = len(re.findall(r"\b[a-z]\b", normalized.lower()))
    if operator_score >= 1 and (digit_count >= 1 or letter_count >= 2):
        return True
    if grouping_score >= 2 and (digit_count >= 1 or variable_count >= 1):
        return True
    return bool(re.search(r"\b[a-z]\s*=\s*.+", normalized.lower()))


def block_looks_like_math_text(text: str) -> bool:
    normalized = (text or "").strip()
    if text_looks_like_equation(normalized):
        return True
    if not normalized:
        return False
    has_digit = any(char.isdigit() for char in normalized)
    has_math_mark = bool(re.search(r"[=+\-/*^()]|[²³¼½¾≤≥]", normalized))
    return has_digit and has_math_mark


def line_looks_like_equation_fragment(text: str) -> bool:
    normalized = (text or "").strip()
    if not normalized:
        return False
    if block_looks_like_math_text(normalized):
        return True
    if len(re.findall(r"[A-Za-z]{2,}", normalized)) >= 3:
        return False
    if re.fullmatch(r"[A-Za-z]", normalized):
        return True
    if re.fullmatch(r"[=+\-/*^()\[\]]", normalized):
        return True
    if re.fullmatch(r"\d+(?:\.\d+)?", normalized):
        return True
    if re.fullmatch(r"[A-Za-z]?\d+(?:\.\d+)?", normalized):
        return True
    if any(symbol in normalized for symbol in ("\uf0b4", "×", "·", "≤", "≥")):
        return True
    if len(normalized) <= 8 and any(char.isdigit() for char in normalized):
        return True
    return False


def join_equation_fragments(fragments: list[str]) -> str:
    parts = [fragment.strip() for fragment in fragments if fragment and fragment.strip()]
    if not parts:
        return ""
    return " ".join(parts)


def math_token_signature(text: str) -> tuple[set[str], set[str]]:
    normalized = normalize_text(text)
    numbers = set(re.findall(r"\d+(?:\.\d+)?", normalized))
    variables = set(re.findall(r"\b[a-z]\b", normalized))
    return numbers, variables


def equation_text_similarity(source_text: str, block_text: str) -> float:
    source_numbers, source_variables = math_token_signature(source_text)
    block_numbers, block_variables = math_token_signature(block_text)

    number_overlap = len(source_numbers & block_numbers)
    variable_overlap = len(source_variables & block_variables)

    score = 0.0
    if exact_text_match(source_text, block_text):
        score += 2.5
    score += number_overlap * 0.7
    score += variable_overlap * 0.4
    if block_looks_like_math_text(block_text):
        score += 0.8
    return score


def text_looks_like_list_item(text: str) -> bool:
    normalized = (text or "").strip()
    if not normalized:
        return False
    return bool(re.match(r"^(\d+\.|\d+\)|[a-zA-Z]\.|[a-zA-Z]\)|[ivxlcdmIVXLCDM]+\.)($|\s+)", normalized))


def first_words_look_like_normal_text(text: str, count: int = 5) -> bool:
    words = re.findall(r"[A-Za-z]+", (text or "").strip())
    if len(words) < count:
        return False
    first_words = words[:count]
    return all(len(word) >= 2 for word in first_words)


def paragraph_is_equation_only(text: str) -> bool:
    normalized = (text or "").strip()
    if not normalized:
        return False
    if looks_like_answer_section_heading(normalized):
        return False
    if not text_looks_like_equation(normalized):
        return False
    word_count = len(re.findall(r"\b[\w']+\b", normalized))
    alpha_words = re.findall(r"[A-Za-z]{2,}", normalized)
    if word_count <= 7 and len(alpha_words) <= 4:
        return True
    return bool(re.fullmatch(r"[\sA-Za-z0-9=+\-*/^().,\[\]{}:;]+", normalized) and word_count <= 10 and "=" in normalized)


def canonical_style(style_name: str | None, outline_level: int | None = None) -> str:
    if outline_level is not None and 0 <= outline_level <= 8:
        return f"h{outline_level + 1}"
    name = (style_name or "").strip().lower()
    if not name:
        return "unstyled"
    if any(keyword in name for keyword in ("equation", "formula", "math", "vector", "professional")):
        return "equation"
    heading_match = re.search(r"heading\s*([1-9])", name)
    if heading_match:
        return f"h{heading_match.group(1)}"
    if name in {"title", "subtitle", "caption", "quote"}:
        return name
    if "list" in name:
        return "list"
    if name in {"normal", "bodytext", "body text"}:
        return "p"
    if "header" in name:
        return "header"
    if "footer" in name:
        return "footer"
    return name


def classify_paragraph(
    text: str,
    style: str | None,
    num_id: str | None,
    has_math: bool,
    has_equation_object: bool = False,
    outline_level: int | None = None,
) -> str:
    style_lower = (style or "").lower()
    equation_style = any(keyword in style_lower for keyword in ("equation", "formula", "math", "vector", "professional"))
    if looks_like_answer_section_heading(text):
        return "normal_text"
    if outline_level is not None or "heading" in style_lower or re.match(r"^h[1-6]$", style_lower):
        return "heading"
    if "title" in style_lower:
        return "title"
    if "subtitle" in style_lower:
        return "subtitle"
    if num_id is not None or text_looks_like_list_item(text):
        return "list"
    if has_math or has_equation_object:
        if paragraph_is_equation_only(text):
            return "equation"
        if text == "":
            return "empty"
        return "normal_text"
    if equation_style:
        return "equation" if paragraph_is_equation_only(text) else "normal_text"
    if "caption" in style_lower:
        return "caption"
    if "quote" in style_lower:
        return "quote"
    if text == "":
        return "empty"
    return "normal_text"


def fingerprint_text(text: str) -> str:
    normalized = re.sub(r"\s+", " ", text.lower().strip())
    return hashlib.sha1(normalized.encode("utf-8")).hexdigest()


def starts_with_capital(text: str) -> bool:
    match = re.search(r"[A-Za-z]", text or "")
    if match is None:
        return True
    return (text or "")[match.start()].isupper()


def is_sentence_case(text: str) -> bool:
    words = re.findall(r"[A-Za-z][A-Za-z']*", text or "")
    if not words:
        return True
    if not starts_with_capital(text):
        return False
    if len(words) == 1:
        return not words[0].isupper() or len(words[0]) == 1
    if any(word.isupper() and len(word) > 1 for word in words[1:]):
        return False
    return True


def footer_rule_issues(item: dict) -> list[str]:
    metadata = item.get("metadata") or {}
    issues = []

    font_families = metadata.get("font_families") or []
    if not font_families:
        issues.append("font family could not be verified")
    elif any((font or "").casefold() != "times new roman" for font in font_families):
        issues.append(f"font should be Times New Roman, found {', '.join(font_families)}")

    font_sizes = metadata.get("font_sizes_pt") or []
    if not font_sizes:
        issues.append("font size could not be verified")
    elif any(abs(size - 10.0) > 0.05 for size in font_sizes):
        rendered_sizes = ", ".join(f"{size:g} pt" for size in font_sizes)
        issues.append(f"font size should be 10 pt, found {rendered_sizes}")

    text = (item.get("text") or "").strip()
    if re.search(r"[A-Za-z]", text):
        if not starts_with_capital(text):
            issues.append("text should start with a capital letter")
        if not is_sentence_case(text):
            issues.append("text should be in sentence case")

    return issues


def build_text_element(
    index: int,
    text: str,
    role: str,
    style: str | None,
    num_id: str | None,
    ilvl: str | None,
    source_part: str,
    metadata: dict | None = None,
) -> DocElement:
    return DocElement(
        index=index,
        text=text,
        role=role,
        style=style,
        num_id=num_id,
        ilvl=ilvl,
        fingerprint=fingerprint_text(text or f"{source_part}-{index}-{role}"),
        source_part=source_part,
        metadata=metadata,
    )


def append_paragraph_object_elements(
    elements: list[DocElement],
    paragraph: ET.Element,
    running_index: int,
    style: str | None,
    source_part: str,
    package_info: dict | None,
    force_equation: bool = False,
) -> int:
    for math_object in extract_office_math_objects(paragraph):
        elements.append(
            DocElement(
                index=running_index,
                text=math_object["text"],
                role="equation",
                style=style or "Equation",
                num_id=None,
                ilvl=None,
                fingerprint=fingerprint_text(f"equation-{running_index}-{math_object['text']}"),
                source_part=source_part,
                has_alt_text=False,
                metadata=math_object.get("metadata"),
            )
        )
        running_index += 1

    for drawing in paragraph_drawings(paragraph, force_equation=force_equation, package_info=package_info):
        drawing_role = "equation" if drawing.get("is_equation") else "image"
        elements.append(
            DocElement(
                index=running_index,
                text=drawing["alt_text"] or ("Equation" if drawing_role == "equation" else "Image"),
                role=drawing_role,
                style=style or ("Equation" if drawing_role == "equation" else "Image"),
                num_id=None,
                ilvl=None,
                fingerprint=fingerprint_text(f"{drawing_role}-{running_index}-{drawing['alt_text']}"),
                source_part=source_part,
                has_alt_text=drawing["has_alt_text"],
                metadata=drawing.get("metadata"),
            )
        )
        running_index += 1
    return running_index


def paragraph_text_metadata(paragraph: ET.Element, outline_level: int | None = None) -> dict | None:
    metadata = extract_paragraph_text_metadata(paragraph)
    if outline_level is None:
        return metadata or None
    merged = dict(metadata or {})
    merged["outline_level"] = outline_level
    return merged


def append_paragraph_inline_elements(
    elements: list[DocElement],
    fragments: list[dict],
    running_index: int,
    style: str | None,
    num_id: str | None,
    ilvl: str | None,
    paragraph_role: str,
    source_part: str,
    text_metadata: dict | None = None,
) -> int:
    text_segments = []
    for fragment in fragments:
        if fragment.get("kind") != "text":
            continue
        text_segments.extend(fragment_text_segments(fragment.get("text", ""), paragraph_role))
    non_text_count = sum(1 for fragment in fragments if fragment.get("kind") != "text")
    text_count = len(text_segments)
    emitted_text = 0

    if paragraph_role == "empty" and not text_count and not non_text_count:
        elements.append(build_text_element(running_index, "", "empty", style, num_id, ilvl, source_part, metadata=text_metadata))
        return running_index + 1

    for fragment in fragments:
        kind = fragment.get("kind")
        if kind == "text":
            for segment in fragment_text_segments(fragment.get("text", ""), paragraph_role):
                text = re.sub(r"\s+", " ", segment).strip()
                if not text:
                    continue

                if paragraph_role == "list":
                    role = "list" if emitted_text == 0 else "normal_text"
                elif paragraph_role in {"heading", "title", "subtitle", "caption", "quote"}:
                    role = paragraph_role
                elif paragraph_role == "equation" and non_text_count == 0:
                    role = "equation"
                else:
                    role = "normal_text"

                elements.append(build_text_element(running_index, text, role, style, num_id, ilvl, source_part, metadata=text_metadata))
                running_index += 1
                emitted_text += 1
            continue

        role = "equation" if kind == "equation" else "image"
        default_text = "Equation" if role == "equation" else "Image"
        elements.append(
            DocElement(
                index=running_index,
                text=(fragment.get("text") or default_text).strip(),
                role=role,
                style=style or default_text,
                num_id=None,
                ilvl=None,
                fingerprint=fingerprint_text(f"{role}-{running_index}-{fragment.get('text', '')}"),
                source_part=source_part,
                has_alt_text=fragment.get("has_alt_text"),
                metadata=fragment.get("metadata"),
            )
        )
        running_index += 1

    return running_index


def is_visual_equation_element(element: DocElement) -> bool:
    if element.role != "equation":
        return False
    metadata = element.metadata or {}
    if metadata.get("prog_id") or metadata.get("ole_target") or metadata.get("image_target"):
        return True
    if not normalize_text(element.text):
        return True
    return False


def element_outline_level(element: DocElement) -> int | None:
    metadata = element.metadata or {}
    outline_level = metadata.get("outline_level")
    if isinstance(outline_level, int):
        return outline_level
    return None


def iter_block_children(parent: ET.Element):
    for child in list(parent):
        if child.tag in {f"{{{NS['w']}}}p", f"{{{NS['w']}}}tbl"}:
            yield child
            continue
        if child.tag in {f"{{{NS['w']}}}sdt", f"{{{NS['w']}}}customXml"}:
            for nested in iter_block_children(child):
                yield nested


def extract_doc_structure(docx_path: str) -> list[DocElement]:
    root = read_docx_xml(docx_path)
    body = root.find(".//w:body", NS)
    if body is None:
        return []
    package_info = inspect_docx_package(docx_path)
    styles_definition = read_styles_definition(docx_path)
    numbering_definition = read_numbering_definition(docx_path)
    numbering_state = {}

    elements = []
    running_index = 0

    for child in iter_block_children(body):
        if child.tag == f"{{{NS['w']}}}p":
            style = get_style(child)
            outline_level = paragraph_outline_level(child, styles_definition)
            num_id, ilvl = get_numbering(child)
            fragments = apply_numbering_prefix_to_fragments(
                extract_paragraph_inline_fragments(child, package_info),
                num_id,
                ilvl,
                numbering_definition,
                numbering_state,
            )
            text = paragraph_text_from_fragments(fragments)
            has_math = paragraph_has_math(child)
            has_equation_object = paragraph_has_embedded_equation(child, package_info, text, style)
            role = classify_paragraph(
                text,
                style,
                num_id,
                has_math,
                has_equation_object,
                outline_level=outline_level,
            )
            running_index = append_paragraph_inline_elements(
                elements,
                fragments,
                running_index,
                style,
                num_id,
                ilvl,
                role,
                "body",
                text_metadata=paragraph_text_metadata(child, outline_level),
            )

        elif child.tag == f"{{{NS['w']}}}tbl":
            cell_texts = []
            for paragraph in child.findall(".//w:p", NS):
                paragraph_num_id, paragraph_ilvl = get_numbering(paragraph)
                paragraph_text = apply_numbering_prefix(
                    get_text(paragraph),
                    paragraph_num_id,
                    paragraph_ilvl,
                    numbering_definition,
                    numbering_state,
                )
                paragraph_style = get_style(paragraph) or get_table_style(child) or "Table"
                if paragraph_text:
                    cell_texts.append(paragraph_text)
                for math_object in extract_office_math_objects(paragraph):
                    if math_object["text"]:
                        cell_texts.append(math_object["text"])
                running_index = append_paragraph_object_elements(
                    elements,
                    paragraph,
                    running_index,
                    paragraph_style,
                    "body",
                    package_info,
                    force_equation=False,
                )
            table_text = "\n".join(cell_texts).strip()
            style = get_table_style(child) or "Table"
            first_row, first_column, repeat_header, allow_row_break_across_pages = table_flags(child)
            elements.append(
                DocElement(
                    index=running_index,
                    text=table_text,
                    role="table",
                    style=style,
                    num_id=None,
                    ilvl=None,
                    fingerprint=fingerprint_text(table_text or f"table-{running_index}"),
                    source_part="body",
                    first_row=first_row,
                    first_column=first_column,
                    repeat_header=repeat_header,
                    allow_row_break_across_pages=allow_row_break_across_pages,
                )
            )
            running_index += 1

    return elements


def extract_header_footer_structure(docx_path: str, prefix: str, role_name: str) -> list[DocElement]:
    elements = []
    running_index = 0
    package_info = inspect_docx_package(docx_path)
    for part_name in read_docx_part_names(docx_path, prefix):
        root = read_docx_xml(docx_path, part_name)
        for paragraph in root.findall(".//w:p", NS):
            text = get_text(paragraph)
            style = get_style(paragraph) or role_name.title()
            if text:
                elements.append(
                    build_text_element(
                        running_index,
                        text,
                        role_name,
                        style,
                        None,
                        None,
                        role_name,
                        metadata=extract_paragraph_text_metadata(paragraph),
                    )
                )
                running_index += 1
            for drawing in paragraph_drawings(paragraph, force_equation=False, package_info=package_info):
                elements.append(
                    DocElement(
                        index=running_index,
                        text=drawing["alt_text"] or role_name.title(),
                        role=role_name,
                        style=style,
                        num_id=None,
                        ilvl=None,
                        fingerprint=fingerprint_text(f"{role_name}-{running_index}-{drawing['alt_text']}"),
                        source_part=role_name,
                        has_alt_text=drawing["has_alt_text"],
                        metadata=drawing.get("metadata"),
                    )
                )
                running_index += 1
    return elements


def extract_pdf_regions(pdf_path: str | Path) -> dict:
    text_blocks = []
    image_blocks = []
    equation_line_blocks = []

    def collect_vector_blocks(
        page: fitz.Page,
        page_number: int,
        width: float,
        height: float,
        start_order: int,
        page_text_blocks: list[dict],
    ) -> list[dict]:
        raw_boxes = []
        for order, drawing in enumerate(page.get_drawings()):
            rect = drawing.get("rect")
            if rect is None:
                continue
            x0, y0, x1, y1 = rect
            bbox = {
                "x0": max(0.0, x0 / width),
                "y0": max(0.0, y0 / height),
                "x1": min(1.0, x1 / width),
                "y1": min(1.0, y1 / height),
            }
            box_width, box_height, box_area = bbox_metrics(bbox)
            if box_width <= 0 and box_height <= 0:
                continue
            if box_width < 0.002 and box_height < 0.002:
                continue
            raw_boxes.append({"bbox_norm": bbox, "order": order, "area": box_area})

        if not raw_boxes:
            return []

        clusters = []
        for raw_box in raw_boxes:
            expanded = expand_bbox(raw_box["bbox_norm"], 0.01, 0.012)
            matching_clusters = [
                cluster for cluster in clusters
                if bboxes_intersect(expanded, expand_bbox(cluster["bbox_norm"], 0.01, 0.012))
            ]

            if not matching_clusters:
                clusters.append(
                    {
                        "bbox_norm": raw_box["bbox_norm"],
                        "count": 1,
                        "order": raw_box["order"],
                        "area_sum": raw_box["area"],
                    }
                )
                continue

            primary = matching_clusters[0]
            primary["bbox_norm"] = merge_bboxes([primary["bbox_norm"], raw_box["bbox_norm"]])
            primary["count"] += 1
            primary["order"] = min(primary["order"], raw_box["order"])
            primary["area_sum"] += raw_box["area"]

            for extra in matching_clusters[1:]:
                primary["bbox_norm"] = merge_bboxes([primary["bbox_norm"], extra["bbox_norm"]])
                primary["count"] += extra["count"]
                primary["order"] = min(primary["order"], extra["order"])
                primary["area_sum"] += extra["area_sum"]
                clusters.remove(extra)

        vector_blocks = []
        vector_order = start_order
        for cluster in clusters:
            bbox = cluster["bbox_norm"]
            box_width, box_height, box_area = bbox_metrics(bbox)
            longest_edge = max(box_width, box_height)
            expected_text_blocks = [
                block
                for block in page_text_blocks
                if bbox_contains_point(
                    bbox,
                    ((block["bbox_norm"]["x0"] + block["bbox_norm"]["x1"]) / 2),
                    ((block["bbox_norm"]["y0"] + block["bbox_norm"]["y1"]) / 2),
                    margin=0.01,
                )
            ]
            text_block_count = len(expected_text_blocks)
            text_char_count = sum(len(normalize_text(block.get("text", ""))) for block in expected_text_blocks)
            is_thin_rule = (box_width >= 0.35 and box_height <= 0.012) or (box_height >= 0.35 and box_width <= 0.012)
            if is_thin_rule:
                continue
            if box_area < 0.01 and cluster["count"] < 6:
                continue
            if longest_edge < 0.08:
                continue
            if text_block_count >= 4 and text_char_count >= 28:
                continue
            vector_blocks.append(
                {
                    "order": vector_order,
                    "page": page_number,
                    "bbox_norm": bbox,
                    "kind": "vector",
                    "path_count": cluster["count"],
                    "text_block_count": text_block_count,
                    "text_char_count": text_char_count,
                }
            )
            vector_order += 1

        return vector_blocks

    with fitz.open(str(pdf_path)) as document:
        global_order = 0
        image_order = 0
        for page_number, page in enumerate(document):
            width = page.rect.width or 1
            height = page.rect.height or 1
            text_dict = page.get_text("dict")
            page_text_blocks = []

            for block in text_dict.get("blocks", []):
                if block.get("type") == 0:
                    line_entries = []
                    lines = []
                    for line in block.get("lines", []):
                        spans = []
                        for span in line.get("spans", []):
                            if span.get("text"):
                                spans.append(span["text"])
                        if spans:
                            line_text = "".join(spans)
                            lines.append(line_text)
                            x0, y0, x1, y1 = line["bbox"]
                            line_entries.append(
                                {
                                    "text": line_text,
                                    "bbox_norm": {"x0": x0 / width, "y0": y0 / height, "x1": x1 / width, "y1": y1 / height},
                                }
                            )
                    block_text = "\n".join(lines).strip()
                    if not block_text:
                        continue
                    x0, y0, x1, y1 = block["bbox"]
                    text_blocks.append(
                        {
                            "order": global_order,
                            "page": page_number,
                            "text": block_text,
                            "bbox_norm": {"x0": x0 / width, "y0": y0 / height, "x1": x1 / width, "y1": y1 / height},
                        }
                    )
                    page_text_blocks.append(text_blocks[-1])
                    global_order += 1

                    grouped_rows: list[list[dict]] = []
                    for line_entry in line_entries:
                        line_bbox = line_entry["bbox_norm"]
                        line_center = (line_bbox["y0"] + line_bbox["y1"]) / 2
                        if not grouped_rows:
                            grouped_rows.append([line_entry])
                            continue
                        previous_row = grouped_rows[-1]
                        previous_bbox = merge_bboxes([entry["bbox_norm"] for entry in previous_row])
                        previous_center = (previous_bbox["y0"] + previous_bbox["y1"]) / 2 if previous_bbox else line_center
                        if abs(line_center - previous_center) <= 0.012:
                            previous_row.append(line_entry)
                        else:
                            grouped_rows.append([line_entry])

                    for row_entries in grouped_rows:
                        fragments = [entry for entry in row_entries if line_looks_like_equation_fragment(entry.get("text", ""))]
                        if not fragments:
                            continue
                        row_text = join_equation_fragments([entry["text"] for entry in fragments])
                        if not row_text:
                            continue
                        if len(fragments) == 1 and not block_looks_like_math_text(row_text):
                            continue
                        merged = merge_bboxes([entry["bbox_norm"] for entry in fragments])
                        if merged is None:
                            continue
                        equation_line_blocks.append(
                            {
                                "order": len(equation_line_blocks),
                                "page": page_number,
                                "bbox_norm": expand_bbox(merged, 0.004, 0.003),
                                "text": row_text,
                                "kind": "equation_line",
                            }
                        )
                elif block.get("type") == 1:
                    x0, y0, x1, y1 = block["bbox"]
                    image_blocks.append(
                        {
                            "order": image_order,
                            "page": page_number,
                            "bbox_norm": {"x0": x0 / width, "y0": y0 / height, "x1": x1 / width, "y1": y1 / height},
                            "kind": "image",
                        }
                    )
                    image_order += 1

            vector_blocks = collect_vector_blocks(page, page_number, width, height, image_order, page_text_blocks)
            image_blocks.extend(vector_blocks)
            image_order += len(vector_blocks)

    return {"text_blocks": text_blocks, "image_blocks": image_blocks, "equation_line_blocks": equation_line_blocks}


def merge_bboxes(boxes: list[dict]) -> dict | None:
    if not boxes:
        return None
    return {
        "x0": min(box["x0"] for box in boxes),
        "y0": min(box["y0"] for box in boxes),
        "x1": max(box["x1"] for box in boxes),
        "y1": max(box["y1"] for box in boxes),
    }


def find_text_anchor_with_index(element: DocElement, text_blocks: list) -> tuple[dict, int | None]:
    if not text_blocks or not normalize_text(element.text):
        return {"page": None, "bbox_norm": None, "text": element.text}, None

    best = None
    best_score = -1.0
    best_index = None
    for order, block in enumerate(text_blocks):
        text_score = similarity(element.text, block.get("text", ""))
        proximity = 1 / (1 + abs(element.index - order))
        score = (text_score * 0.88) + (proximity * 0.12)
        if element.role == "header":
            score += max(0, 0.18 - block["bbox_norm"]["y0"])
        if element.role == "footer":
            score += max(0, block["bbox_norm"]["y1"] - 0.82)
        if score > best_score:
            best_score = score
            best = block
            best_index = order

    if best is None:
        return {"page": None, "bbox_norm": None, "text": element.text}, None

    return (
        {"page": best.get("page"), "bbox_norm": best.get("bbox_norm"), "text": best.get("text", element.text)},
        best_index,
    )


def find_text_anchor(element: DocElement, text_blocks: list) -> dict:
    preview, _ = find_text_anchor_with_index(element, text_blocks)
    return preview


def find_text_anchor_exact(
    element: DocElement,
    text_blocks: list,
    used_indexes: set[int],
    start_index: int = 0,
    zone: str | None = None,
) -> tuple[dict, int | None]:
    def zone_ok(block: dict) -> bool:
        bbox = block.get("bbox_norm") or {}
        if zone == "header":
            return bbox.get("y0", 1) <= 0.22
        if zone == "footer":
            return bbox.get("y1", 0) >= 0.78
        return True

    for idx in range(start_index, len(text_blocks)):
        if idx in used_indexes:
            continue
        block = text_blocks[idx]
        if not zone_ok(block):
            continue
        if exact_text_match(element.text, block.get("text", "")):
            used_indexes.add(idx)
            return (
                {
                    "page": block.get("page"),
                    "bbox_norm": block.get("bbox_norm"),
                    "text": block.get("text", element.text),
                },
                idx,
            )

    for idx, block in enumerate(text_blocks):
        if idx in used_indexes:
            continue
        if not zone_ok(block):
            continue
        if exact_text_match(element.text, block.get("text", "")):
            used_indexes.add(idx)
            return (
                {
                    "page": block.get("page"),
                    "bbox_norm": block.get("bbox_norm"),
                    "text": block.get("text", element.text),
                },
                idx,
            )

    return find_text_anchor_with_index(element, text_blocks)


def find_equation_text_anchor_exact(
    element: DocElement,
    text_blocks: list,
    used_indexes: set[int],
    start_index: int = 0,
    preferred_page: int | None = None,
) -> tuple[dict | None, int | None]:
    if not text_blocks:
        return None, None

    candidate_indexes = list(range(start_index, len(text_blocks))) + list(range(0, start_index))
    best_index = None
    best_score = -1.0

    for idx in candidate_indexes:
        if idx in used_indexes:
            continue
        block = text_blocks[idx]
        block_text = block.get("text", "")
        if not block_looks_like_math_text(block_text):
            continue

        score = equation_text_similarity(element.text, block_text)
        if preferred_page is not None and block.get("page") == preferred_page:
            score += 0.6
        score += 0.2 / (1 + abs(idx - start_index))

        if score > best_score:
            best_score = score
            best_index = idx

    if best_index is None or best_score < 0.8:
        return None, None

    used_indexes.add(best_index)
    block = text_blocks[best_index]
    return (
        {
            "page": block.get("page"),
            "bbox_norm": block.get("bbox_norm"),
            "text": block.get("text", element.text),
        },
        best_index,
    )


def find_visual_equation_text_anchor_local(
    element: DocElement,
    text_blocks: list,
    used_indexes: set[int],
    start_index: int = 0,
    preferred_page: int | None = None,
) -> tuple[dict | None, int | None]:
    if not text_blocks:
        return None, None

    def ordered_candidates(offsets: tuple[int, ...]) -> list[int]:
        indexes = []
        for offset in offsets:
            idx = start_index + offset
            if 0 <= idx < len(text_blocks) and idx not in indexes:
                indexes.append(idx)
        if preferred_page is not None:
            indexes = [idx for idx in indexes if text_blocks[idx].get("page") == preferred_page]
        return indexes

    def best_candidate(candidate_indexes: list[int]) -> tuple[int | None, float | None]:
        best_index = None
        best_score = None
        for idx in candidate_indexes:
            if idx in used_indexes:
                continue

            block = text_blocks[idx]
            block_text = block.get("text", "")
            if not block_text.strip():
                continue

            score = equation_text_similarity(element.text, block_text)
            if any(symbol in block_text for symbol in ("=", "+", "/", "^", "(", ")", "[", "]")):
                score += 0.4
            if preferred_page is not None and block.get("page") == preferred_page:
                score += 0.6
            if idx < start_index:
                score += 0.35
            score -= 0.35 * abs(idx - start_index)

            if best_score is None or score > best_score:
                best_score = score
                best_index = idx
        return best_index, best_score

    best_index, best_score = best_candidate(ordered_candidates((-1, 0, 1)))
    if best_index is None or (best_score is not None and best_score < 1.0):
        best_index, best_score = best_candidate(ordered_candidates((-2, 2, -3, 3)))

    if best_index is None or (best_score is not None and best_score < 1.0):
        return None, None

    used_indexes.add(best_index)
    block = text_blocks[best_index]
    return (
        {
            "page": block.get("page"),
            "bbox_norm": block.get("bbox_norm"),
            "text": block.get("text", element.text),
        },
        best_index,
    )


def find_equation_line_anchor(
    element: DocElement,
    equation_line_blocks: list,
    used_indexes: set[int],
    start_index: int = 0,
    preferred_page: int | None = None,
) -> tuple[dict | None, list[int]]:
    if not equation_line_blocks:
        return None, []

    expected_size_class = element_visual_size_class(element)
    expected_aspect_ratio = element_visual_aspect_ratio(element)

    def ordered_candidates() -> list[int]:
        ordered = list(range(start_index, len(equation_line_blocks))) + list(range(0, start_index))
        seen = set()
        result = []
        for idx in ordered:
            if idx in seen:
                continue
            seen.add(idx)
            result.append(idx)
        return result

    def candidate_group(index: int) -> list[int]:
        indexes = [index]
        anchor = equation_line_blocks[index]
        anchor_bbox = anchor.get("bbox_norm")
        for next_index in range(index + 1, min(index + 3, len(equation_line_blocks))):
            if next_index in used_indexes:
                break
            candidate = equation_line_blocks[next_index]
            if candidate.get("page") != anchor.get("page"):
                break
            candidate_bbox = candidate.get("bbox_norm")
            if not anchor_bbox or not candidate_bbox:
                break
            vertical_gap = candidate_bbox["y0"] - anchor_bbox["y1"]
            if vertical_gap > 0.04:
                break
            if bbox_horizontal_overlap_ratio(anchor_bbox, candidate_bbox) < 0.35:
                break
            indexes.append(next_index)
            anchor_bbox = merge_bboxes([anchor_bbox, candidate_bbox])
        return indexes

    best_indexes: list[int] = []
    best_score = None
    for idx in ordered_candidates():
        if idx in used_indexes:
            continue
        candidate_indexes = candidate_group(idx)
        blocks = [equation_line_blocks[candidate_index] for candidate_index in candidate_indexes]
        merged_bbox = merge_bboxes([block.get("bbox_norm") for block in blocks if block.get("bbox_norm")])
        if merged_bbox is None:
            continue

        score = 0.0
        if preferred_page is not None and blocks[0].get("page") == preferred_page:
            score += 1.8
        elif preferred_page is not None:
            score -= 0.4

        candidate_size_class = size_class_from_bbox(merged_bbox)
        if expected_size_class and candidate_size_class == expected_size_class:
            score += 1.6
        elif expected_size_class and candidate_size_class is not None:
            score -= 0.8

        score += aspect_ratio_similarity(expected_aspect_ratio, aspect_ratio_from_bbox(merged_bbox))

        candidate_text = " ".join(block.get("text", "") for block in blocks)
        if block_looks_like_math_text(candidate_text):
            score += 0.6
        elif any(symbol in candidate_text for symbol in ("=", "+", "-", "/", "^")):
            score += 0.3

        score -= abs(idx - start_index) * 0.03

        if best_score is None or score > best_score:
            best_score = score
            best_indexes = candidate_indexes

    if not best_indexes or (best_score is not None and best_score < 0.45):
        return None, []

    for idx in best_indexes:
        used_indexes.add(idx)

    matched_blocks = [equation_line_blocks[idx] for idx in best_indexes]
    merged_bbox = merge_bboxes([block["bbox_norm"] for block in matched_blocks if block.get("bbox_norm")])
    if merged_bbox is None:
        return None, []

    return (
        {
            "page": matched_blocks[0].get("page"),
            "bbox_norm": expand_bbox(merged_bbox, 0.004, 0.004),
            "text": "\n".join(block.get("text", element.text) for block in matched_blocks),
        },
        best_indexes,
    )


def element_visual_size_class(element: DocElement) -> str | None:
    metadata = element.metadata or {}
    size_class = size_class_from_dimensions(metadata.get("width_pt"), metadata.get("height_pt"))
    if size_class:
        return size_class
    if is_visual_equation_element(element):
        return "small"
    return None


def element_visual_aspect_ratio(element: DocElement) -> float | None:
    metadata = element.metadata or {}
    width = metadata.get("width_pt")
    height = metadata.get("height_pt")
    if not width or not height:
        return None
    if width <= 0 or height <= 0:
        return None
    return width / height


def find_image_anchor(
    element: DocElement,
    image_blocks: list,
    used_indexes: set[int],
    start_index: int = 0,
    preferred_page: int | None = None,
) -> tuple[dict, int | None]:
    expected_size_class = element_visual_size_class(element)
    expected_aspect_ratio = element_visual_aspect_ratio(element)
    visual_equation = is_visual_equation_element(element)

    def preview_for(index: int) -> tuple[dict, int]:
        block = image_blocks[index]
        used_indexes.add(index)
        return (
            {"page": block.get("page"), "bbox_norm": block.get("bbox_norm"), "text": element.text},
            index,
        )

    def candidate_indexes() -> list[int]:
        ordered = list(range(start_index, len(image_blocks))) + list(range(0, start_index))
        seen = set()
        result = []
        for idx in ordered:
            if idx in seen:
                continue
            seen.add(idx)
            result.append(idx)
        return result

    best_index = None
    best_score = None
    for idx in candidate_indexes():
        if idx in used_indexes:
            continue
        block = image_blocks[idx]
        score = 0.0

        if preferred_page is not None and block.get("page") == preferred_page:
            score += 2.0
        elif preferred_page is not None:
            score -= 0.4

        candidate_size_class = size_class_from_bbox(block.get("bbox_norm"))
        if expected_size_class and candidate_size_class == expected_size_class:
            score += 1.6
        elif expected_size_class and candidate_size_class is not None:
            score -= 0.8

        score += aspect_ratio_similarity(expected_aspect_ratio, aspect_ratio_from_bbox(block.get("bbox_norm")))

        if block.get("kind") == "vector":
            text_block_count = block.get("text_block_count", 0)
            text_char_count = block.get("text_char_count", 0)
            if text_block_count >= 3:
                score -= 0.45 * text_block_count
            if text_char_count >= 18:
                score -= min(1.6, text_char_count / 30.0)
            if visual_equation and (text_block_count >= 2 or text_char_count >= 10):
                score -= 2.2

        distance_penalty = abs(idx - start_index) * 0.03
        score -= distance_penalty

        if best_score is None or score > best_score:
            best_score = score
            best_index = idx

    if best_index is not None and (best_score is None or best_score >= 0.25):
        return preview_for(best_index)

    if expected_size_class:
        for idx in candidate_indexes():
            if idx in used_indexes:
                continue
            if preferred_page is not None and image_blocks[idx].get("page") != preferred_page:
                continue
            if size_class_from_bbox(image_blocks[idx].get("bbox_norm")) == expected_size_class:
                return preview_for(idx)
        for idx in candidate_indexes():
            if idx in used_indexes:
                continue
            if size_class_from_bbox(image_blocks[idx].get("bbox_norm")) == expected_size_class:
                return preview_for(idx)

    if visual_equation:
        return {"page": None, "bbox_norm": None, "text": element.text}, None

    if preferred_page is not None:
        for idx in candidate_indexes():
            if idx in used_indexes:
                continue
            if image_blocks[idx].get("page") == preferred_page:
                return preview_for(idx)

    for idx in candidate_indexes():
        if idx in used_indexes:
            continue
        return preview_for(idx)

    return {"page": None, "bbox_norm": None, "text": element.text}, None


def expand_table_bbox_with_images(page: int, merged: dict | None, image_blocks: list) -> dict | None:
    if merged is None:
        return None
    extra_boxes = [merged]
    for block in image_blocks:
        if block.get("page") != page or not block.get("bbox_norm"):
            continue
        box = block["bbox_norm"]
        center_x = (box["x0"] + box["x1"]) / 2
        center_y = (box["y0"] + box["y1"]) / 2
        if (
            merged["x0"] - 0.08 <= center_x <= merged["x1"] + 0.08
            and merged["y0"] - 0.08 <= center_y <= merged["y1"] + 0.08
        ):
            extra_boxes.append(box)
    return merge_bboxes(extra_boxes)


def find_table_anchor(element: DocElement, text_blocks: list, image_blocks: list | None = None) -> dict:
    lines = [line.strip() for line in element.text.splitlines() if line.strip()]
    if not lines:
        return {"page": None, "bbox_norm": None, "text": element.text}

    matches = []
    for block in text_blocks:
        best_line_score = max((similarity(line, block.get("text", "")) for line in lines), default=0.0)
        if best_line_score >= 0.46:
            matches.append((best_line_score, block))

    if not matches:
        return find_text_anchor(element, text_blocks)

    page_counts = {}
    for _, block in matches:
        page = block.get("page")
        page_counts[page] = page_counts.get(page, 0) + 1
    dominant_page = max(page_counts, key=page_counts.get)
    page_matches = [block for _, block in matches if block.get("page") == dominant_page]
    merged = merge_bboxes([block["bbox_norm"] for block in page_matches if block.get("bbox_norm")])
    if image_blocks:
        merged = expand_table_bbox_with_images(dominant_page, merged, image_blocks)
    return {"page": dominant_page, "bbox_norm": merged, "text": element.text}


def find_table_anchor_exact(element: DocElement, text_blocks: list, image_blocks: list, used_indexes: set[int]) -> dict:
    lines = [line.strip() for line in element.text.splitlines() if line.strip()]
    if not lines:
        return find_table_anchor(element, text_blocks, image_blocks)

    matched_indexes = []
    for idx, block in enumerate(text_blocks):
        if idx in used_indexes:
            continue
        block_text = block.get("text", "")
        if any(exact_text_match(line, block_text) for line in lines):
            matched_indexes.append(idx)

    if not matched_indexes:
        return find_table_anchor(element, text_blocks, image_blocks)

    for idx in matched_indexes:
        used_indexes.add(idx)

    page_counts = {}
    for idx in matched_indexes:
        page = text_blocks[idx].get("page")
        page_counts[page] = page_counts.get(page, 0) + 1
    dominant_page = max(page_counts, key=page_counts.get)
    page_matches = [text_blocks[idx] for idx in matched_indexes if text_blocks[idx].get("page") == dominant_page]
    merged = merge_bboxes([block["bbox_norm"] for block in page_matches if block.get("bbox_norm")])
    merged = expand_table_bbox_with_images(dominant_page, merged, image_blocks)
    return {"page": dominant_page, "bbox_norm": merged, "text": element.text}


def build_style_item(
    element: DocElement,
    text_blocks: list,
    image_blocks: list,
    used_image_indexes: set[int],
    image_cursor: int,
    preferred_page: int | None = None,
) -> tuple[dict, int]:
    item = {
        "index": element.index,
        "text": element.text,
        "role": element.role,
        "style_name": element.style or "Unstyled",
        "style_tag": canonical_style(element.style, element_outline_level(element)),
        "num_id": element.num_id,
        "ilvl": element.ilvl,
        "source_part": element.source_part,
        "has_alt_text": element.has_alt_text,
        "first_row": element.first_row,
        "first_column": element.first_column,
        "repeat_header": element.repeat_header,
        "allow_row_break_across_pages": element.allow_row_break_across_pages,
        "metadata": element.metadata,
    }

    if element.role == "image" or is_visual_equation_element(element):
        preview, matched_image_index = find_image_anchor(
            element,
            image_blocks,
            used_image_indexes,
            image_cursor,
            preferred_page=preferred_page,
        )
        if matched_image_index is not None:
            image_cursor = max(image_cursor, matched_image_index + 1)
    elif element.role == "table":
        preview = find_table_anchor(element, text_blocks, image_blocks)
    elif element.role == "empty":
        preview = {"page": None, "bbox_norm": None, "text": ""}
    else:
        preview = find_text_anchor(element, text_blocks)

    item["preview"] = preview
    return item, image_cursor


def build_body_style_item(
    element: DocElement,
    text_blocks: list,
    equation_line_blocks: list,
    image_blocks: list,
    used_text_indexes: set[int],
    used_equation_line_indexes: set[int],
    used_visual_text_indexes: set[int],
    used_image_indexes: set[int],
    equation_cursor: int,
    image_cursor: int,
    text_cursor: int,
    preferred_page: int | None = None,
) -> tuple[dict, int, int, int]:
    item = {
        "index": element.index,
        "text": element.text,
        "role": element.role,
        "style_name": element.style or "Unstyled",
        "style_tag": canonical_style(element.style, element_outline_level(element)),
        "num_id": element.num_id,
        "ilvl": element.ilvl,
        "source_part": element.source_part,
        "has_alt_text": element.has_alt_text,
        "first_row": element.first_row,
        "first_column": element.first_column,
        "repeat_header": element.repeat_header,
        "allow_row_break_across_pages": element.allow_row_break_across_pages,
        "metadata": element.metadata,
    }

    if element.role == "equation":
        if is_visual_equation_element(element):
            preview, matched_line_indexes = find_equation_line_anchor(
                element,
                equation_line_blocks,
                used_equation_line_indexes,
                equation_cursor,
                preferred_page=preferred_page,
            )
            if matched_line_indexes:
                equation_cursor = max(equation_cursor, matched_line_indexes[-1] + 1)
            else:
                preview, matched_image_index = find_image_anchor(
                    element,
                    image_blocks,
                    used_image_indexes,
                    image_cursor,
                    preferred_page=preferred_page,
                )
                if matched_image_index is not None:
                    image_cursor = max(image_cursor, matched_image_index + 1)
                else:
                    preview, matched_index = find_visual_equation_text_anchor_local(
                        element,
                        text_blocks,
                        used_visual_text_indexes,
                        start_index=text_cursor,
                        preferred_page=preferred_page,
                    )
                    if matched_index is None:
                        preview = {"page": None, "bbox_norm": None, "text": element.text}
        else:
            preview, matched_index = find_equation_text_anchor_exact(
                element,
                text_blocks,
                used_text_indexes,
                start_index=text_cursor,
                preferred_page=preferred_page,
            )
            if matched_index is not None:
                text_cursor = matched_index + 1
            else:
                preview, matched_index = find_text_anchor_exact(
                    element,
                    text_blocks,
                    used_text_indexes,
                    start_index=text_cursor,
                )
                if matched_index is not None:
                    text_cursor = matched_index + 1
    elif element.role == "image" or is_visual_equation_element(element):
        preview, matched_image_index = find_image_anchor(
            element,
            image_blocks,
            used_image_indexes,
            image_cursor,
            preferred_page=preferred_page,
        )
        if matched_image_index is not None:
            image_cursor = max(image_cursor, matched_image_index + 1)
    elif element.role == "table":
        preview = find_table_anchor_exact(element, text_blocks, image_blocks, used_text_indexes)
    elif element.role == "empty":
        preview = {"page": None, "bbox_norm": None, "text": ""}
    else:
        preview, matched_index = find_text_anchor_exact(element, text_blocks, used_text_indexes, start_index=text_cursor)
        if matched_index is not None:
            text_cursor = matched_index + 1

    item["preview"] = preview
    return item, equation_cursor, image_cursor, text_cursor


def expand_repeating_part_items(
    elements: list[DocElement],
    text_blocks: list,
    role_name: str,
) -> list[dict]:
    if not elements:
        return []

    zone = "header" if role_name == "header" else "footer"
    repeated_items = []

    for block in text_blocks:
        bbox = block.get("bbox_norm") or {}
        if role_name == "header" and bbox.get("y0", 1) > 0.22:
            continue
        if role_name == "footer" and bbox.get("y1", 0) < 0.78:
            continue

        matched = None
        for element in elements:
            if exact_text_match(element.text, block.get("text", "")):
                matched = element
                break

        if matched is None:
            continue

        repeated_items.append(
            {
                "index": matched.index,
                "text": matched.text,
                "role": matched.role,
                "style_name": matched.style or role_name.title(),
                "style_tag": canonical_style(matched.style, element_outline_level(matched)),
                "num_id": matched.num_id,
                "ilvl": matched.ilvl,
                "source_part": matched.source_part,
                "has_alt_text": matched.has_alt_text,
                "first_row": matched.first_row,
                "first_column": matched.first_column,
                "repeat_header": matched.repeat_header,
                "allow_row_break_across_pages": matched.allow_row_break_across_pages,
                "metadata": matched.metadata,
                "preview": {
                    "page": block.get("page"),
                    "bbox_norm": block.get("bbox_norm"),
                    "text": block.get("text", matched.text),
                },
                "zone": zone,
            }
        )

    return repeated_items


def nearest_page(items: list[dict], index: int) -> int | None:
    if not items:
        return None
    best_page = None
    best_gap = None
    for item in items:
        page = item.get("preview", {}).get("page")
        if page is None:
            continue
        gap = abs((item.get("index") or 0) - index)
        if best_gap is None or gap < best_gap:
            best_gap = gap
            best_page = page
    return best_page


def synthesize_empty_preview(items: list[dict], position: int) -> dict:
    current = items[position]
    previous = next(
        (
            item for item in reversed(items[:position])
            if item.get("preview", {}).get("page") is not None and item.get("preview", {}).get("bbox_norm")
        ),
        None,
    )
    following = next(
        (
            item for item in items[position + 1:]
            if item.get("preview", {}).get("page") is not None and item.get("preview", {}).get("bbox_norm")
        ),
        None,
    )

    if previous and following and previous["preview"]["page"] == following["preview"]["page"]:
        prev_box = previous["preview"]["bbox_norm"]
        next_box = following["preview"]["bbox_norm"]
        top = min(max(prev_box["y1"] + 0.004, 0), 0.97)
        bottom = max(min(next_box["y0"] - 0.004, 1), top + 0.012)
        return {
            "page": previous["preview"]["page"],
            "bbox_norm": {
                "x0": min(prev_box["x0"], next_box["x0"]),
                "y0": top,
                "x1": max(prev_box["x1"], next_box["x1"]),
                "y1": bottom,
            },
            "text": current.get("text", ""),
        }

    anchor = following or previous
    if anchor:
        anchor_box = anchor["preview"]["bbox_norm"]
        page = anchor["preview"]["page"]
        if following and not previous:
            y0 = max(anchor_box["y0"] - 0.03, 0)
            y1 = max(y0 + 0.015, min(anchor_box["y0"] - 0.004, 1))
        else:
            y0 = min(anchor_box["y1"] + 0.004, 0.97)
            y1 = min(y0 + 0.02, 0.995)
        return {
            "page": page,
            "bbox_norm": {
                "x0": anchor_box["x0"],
                "y0": y0,
                "x1": anchor_box["x1"],
                "y1": y1,
            },
            "text": current.get("text", ""),
        }

    return {"page": None, "bbox_norm": None, "text": current.get("text", "")}


def backfill_empty_previews(items: list[dict]) -> None:
    for position, item in enumerate(items):
        if item.get("role") != "empty":
            continue
        if item.get("preview", {}).get("bbox_norm"):
            continue
        item["preview"] = synthesize_empty_preview(items, position)


def alert_page_for(item: dict, items: list[dict]) -> int | None:
    page = item.get("preview", {}).get("page")
    if page is not None:
        return page
    return nearest_page(items, item["index"])


def build_style_map(docx_path: str | Path, pdf_path: str | Path) -> dict:
    path = Path(docx_path)
    if path.suffix.lower() != ".docx":
        return {
            "available": False,
            "message": "Style map is available only for DOCX files.",
            "items": [],
            "alerts": [],
            "summary": {"count": 0, "roles": {}, "alerts": 0},
        }

    body_items = extract_doc_structure(str(path))
    header_items = extract_header_footer_structure(str(path), "word/header", "header")
    footer_items = extract_header_footer_structure(str(path), "word/footer", "footer")

    pdf_regions = extract_pdf_regions(pdf_path)
    text_blocks = pdf_regions["text_blocks"]
    equation_line_blocks = pdf_regions["equation_line_blocks"]
    image_blocks = pdf_regions["image_blocks"]

    items = []
    equation_cursor = 0
    image_cursor = 0
    used_text_indexes = set()
    used_equation_line_indexes = set()
    used_visual_text_indexes = set()
    used_image_indexes = set()
    text_cursor = 0
    last_body_page = None

    for item in expand_repeating_part_items(header_items, text_blocks, "header"):
        item["map_index"] = len(items)
        items.append(item)

    for element in body_items:
        item, equation_cursor, image_cursor, text_cursor = build_body_style_item(
            element,
            text_blocks,
            equation_line_blocks,
            image_blocks,
            used_text_indexes,
            used_equation_line_indexes,
            used_visual_text_indexes,
            used_image_indexes,
            equation_cursor,
            image_cursor,
            text_cursor,
            preferred_page=last_body_page,
        )
        item["map_index"] = len(items)
        items.append(item)
        preview_page = item.get("preview", {}).get("page")
        if preview_page is not None:
            last_body_page = preview_page

    for item in expand_repeating_part_items(footer_items, text_blocks, "footer"):
        item["map_index"] = len(items)
        items.append(item)

    backfill_empty_previews(items)

    role_counts = {}
    for item in items:
        role = item.get("role") or "unknown"
        role_counts[role] = role_counts.get(role, 0) + 1

    alerts = []
    empty_run = []
    for item in items:
        if item.get("role") == "empty" and item.get("source_part") == "body":
            empty_run.append(item)
        else:
            if len(empty_run) >= 2:
                alerts.append(
                    {
                        "type": "double_empty_paragraph",
                        "message": f"{len(empty_run)} consecutive empty paragraphs",
                        "page": alert_page_for(empty_run[0], items),
                        "target_map_index": empty_run[0]["map_index"],
                    }
                )
            elif len(empty_run) == 1:
                alerts.append(
                    {
                        "type": "empty_paragraph",
                        "message": "Empty paragraph detected",
                        "page": alert_page_for(empty_run[0], items),
                        "target_map_index": empty_run[0]["map_index"],
                    }
                )
            empty_run = []

        if item.get("role") == "image" and not item.get("has_alt_text"):
            alerts.append(
                {
                    "type": "image_missing_alt_text",
                    "message": "Image without alt text",
                    "page": item.get("preview", {}).get("page"),
                    "target_map_index": item.get("map_index"),
                }
            )

        if (
            item.get("role") == "table"
            and (
                not item.get("first_row")
                or not item.get("first_column")
                or not item.get("repeat_header")
                or item.get("allow_row_break_across_pages") is not False
            )
        ):
            missing = []
            if not item.get("first_row"):
                missing.append("header row")
            if not item.get("first_column"):
                missing.append("first column")
            if not item.get("repeat_header"):
                missing.append("repeat header row")
            if item.get("allow_row_break_across_pages") is not False:
                missing.append("break across pages disabled")
            alerts.append(
                {
                    "type": "table_missing_structure",
                    "message": f"Table missing {' and '.join(missing)}",
                    "page": item.get("preview", {}).get("page"),
                    "target_map_index": item.get("map_index"),
                }
            )

        if (
            item.get("role") == "footer"
            and item.get("source_part") == "footer"
            and isinstance(item.get("metadata"), dict)
            and (
                item["metadata"].get("font_families") is not None
                or item["metadata"].get("font_sizes_pt") is not None
            )
        ):
            issues = footer_rule_issues(item)
            if issues:
                alerts.append(
                    {
                        "type": "footer_format_mismatch",
                        "message": f"Footer should be 10 pt Times New Roman in sentence case and start with a capital: {'; '.join(issues)}",
                        "page": item.get("preview", {}).get("page"),
                        "target_map_index": item.get("map_index"),
                    }
                )

    if len(empty_run) >= 2:
        alerts.append(
            {
                "type": "double_empty_paragraph",
                "message": f"{len(empty_run)} consecutive empty paragraphs",
                "page": alert_page_for(empty_run[0], items),
                "target_map_index": empty_run[0]["map_index"],
            }
        )
    elif len(empty_run) == 1:
        alerts.append(
            {
                "type": "empty_paragraph",
                "message": "Empty paragraph detected",
                "page": alert_page_for(empty_run[0], items),
                "target_map_index": empty_run[0]["map_index"],
            }
        )

    return {
        "available": True,
        "message": "",
        "items": items,
        "alerts": alerts,
        "summary": {"count": len(items), "roles": role_counts, "alerts": len(alerts)},
    }


def compare_docx_styles(path_a: str | Path, path_b: str | Path, layout_a: list | None = None, layout_b: list | None = None) -> dict:
    right_path = Path(path_b)
    if right_path.suffix.lower() != ".docx":
        return {
            "results": [],
            "summary": {
                "available": False,
                "differences": 0,
                "message": "Style inspection is available only when the right input is a DOCX file.",
            },
        }

    # Kept as a compatibility wrapper for the existing API shape.
    return {"results": [], "summary": {"available": True, "differences": 0, "left_items": 0, "right_items": 0}}
