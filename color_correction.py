from __future__ import annotations

import colorsys
import os
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

from lxml import etree

try:
    from colour import Color  # type: ignore
except ImportError:
    class Color:  # type: ignore[override]
        def __init__(self, value: str | None = None, hsl: tuple[float, float, float] | None = None):
            if hsl is not None:
                hue, saturation, lightness = hsl
                red, green, blue = colorsys.hls_to_rgb(hue, lightness, saturation)
            else:
                candidate = str(value or "").strip()
                if not candidate.startswith("#"):
                    candidate = f"#{candidate}"
                if len(candidate) != 7:
                    raise ValueError("Color value must be a 6-digit hex string.")
                red = int(candidate[1:3], 16) / 255
                green = int(candidate[3:5], 16) / 255
                blue = int(candidate[5:7], 16) / 255
            self._rgb = (red, green, blue)

        def get_rgb(self) -> tuple[float, float, float]:
            return self._rgb

        @property
        def hsl(self) -> tuple[float, float, float]:
            hue, lightness, saturation = colorsys.rgb_to_hls(*self._rgb)
            return hue, saturation, lightness

        @property
        def hex_l(self) -> str:
            red, green, blue = [round(max(0.0, min(channel, 1.0)) * 255) for channel in self._rgb]
            return f"#{red:02x}{green:02x}{blue:02x}"


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
DRAWING_NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
EXTRA_NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}
THEME_ALIASES = {
    "text1": "dk1",
    "text2": "dk2",
    "background1": "lt1",
    "background2": "lt2",
    "hyperlink": "hlink",
    "followedhyperlink": "folHlink",
}


@dataclass(slots=True)
class ProcessedDocument:
    filename: str
    output_bytes: bytes
    fixed_elements: int
    changed: bool


@dataclass(slots=True)
class StyleInfo:
    style_type: str | None
    based_on: str | None
    color: str | None
    shading: str | None


@dataclass(slots=True)
class StyleContext:
    styles: dict[str, StyleInfo]
    default_color: str
    default_shading: str | None


HIGHLIGHT_COLORS = {
    "black": "000000",
    "blue": "0000FF",
    "cyan": "00FFFF",
    "darkBlue": "00008B",
    "darkCyan": "008B8B",
    "darkGray": "A9A9A9",
    "darkGreen": "006400",
    "darkMagenta": "8B008B",
    "darkRed": "8B0000",
    "darkYellow": "808000",
    "green": "00FF00",
    "lightGray": "D3D3D3",
    "magenta": "FF00FF",
    "none": "FFFFFF",
    "red": "FF0000",
    "white": "FFFFFF",
    "yellow": "FFFF00",
}


def normalize_hex_color(value: str) -> str:
    candidate = value.strip().upper()
    if not candidate:
        return "#FFFFFF"
    if not candidate.startswith("#"):
        candidate = f"#{candidate}"
    if len(candidate) != 7:
        raise ValueError("Background color must be a 6-digit hex value like #FFFFFF.")
    int(candidate[1:], 16)
    return candidate


def luminance(color: Color) -> float:
    rgb = color.get_rgb()
    adjusted = []
    for channel in rgb:
        if channel <= 0.03928:
            adjusted.append(channel / 12.92)
        else:
            adjusted.append(((channel + 0.055) / 1.055) ** 2.4)
    return 0.2126 * adjusted[0] + 0.7152 * adjusted[1] + 0.0722 * adjusted[2]


def contrast_ratio(foreground: Color, background: Color) -> float:
    lum_1 = luminance(foreground)
    lum_2 = luminance(background)
    return (max(lum_1, lum_2) + 0.05) / (min(lum_1, lum_2) + 0.05)


def is_low_contrast(hex_color: str, background: str = "#FFFFFF", threshold: float = 4.5) -> bool:
    try:
        return contrast_ratio(Color(hex_color), Color(background)) < threshold
    except Exception:
        return False


def adjust_color_for_contrast(hex_color: str, background: str = "#FFFFFF", threshold: float = 4.5) -> str:
    try:
        foreground = Color(hex_color)
        background_color = Color(background)

        hue, saturation, lightness = foreground.hsl
        should_darken = luminance(background_color) > 0.5

        best_candidate = foreground
        best_ratio = contrast_ratio(foreground, background_color)

        saturation_steps = [1.0, 0.85, 0.65, 0.45, 0.25, 0.0]
        for saturation_factor in saturation_steps:
            candidate_saturation = max(0.0, min(1.0, saturation * saturation_factor))
            for index in range(1, 60):
                delta = index * 0.02
                next_lightness = max(0, lightness - delta) if should_darken else min(1, lightness + delta)
                candidate = Color(hsl=(hue, candidate_saturation, next_lightness))
                candidate_ratio = contrast_ratio(candidate, background_color)

                if candidate_ratio > best_ratio:
                    best_candidate = candidate
                    best_ratio = candidate_ratio

                if candidate_ratio >= threshold:
                    return candidate.hex_l.replace("#", "").upper()

        for fallback_hex in ["#000000", "#FFFFFF"]:
            fallback_candidate = Color(fallback_hex)
            fallback_ratio = contrast_ratio(fallback_candidate, background_color)
            if fallback_ratio > best_ratio:
                best_candidate = fallback_candidate
                best_ratio = fallback_ratio
            if fallback_ratio >= threshold:
                return fallback_candidate.hex_l.replace("#", "").upper()

        return best_candidate.hex_l.replace("#", "").upper()
    except Exception:
        return hex_color.replace("#", "").upper()


def _parse_theme_colors(temp_dir: Path) -> dict[str, str]:
    theme_path = temp_dir / "word" / "theme" / "theme1.xml"
    if not theme_path.exists():
        return {}

    try:
        tree = etree.parse(str(theme_path))
    except Exception:
        return {}

    colors: dict[str, str] = {}
    for element in tree.xpath("//a:clrScheme/*", namespaces=DRAWING_NS):
        local_name = etree.QName(element).localname
        value = element.get("lastClr") or element.get("val")
        if value:
            colors[local_name] = value.upper()
    return colors


def _apply_tint_or_shade(hex_color: str, tint: str | None, shade: str | None) -> str:
    try:
        rgb = [int(hex_color[index:index + 2], 16) for index in range(0, 6, 2)]

        if tint:
            factor = int(tint, 16) / 255
            rgb = [round(channel + (255 - channel) * factor) for channel in rgb]

        if shade:
            factor = int(shade, 16) / 255
            rgb = [round(channel * factor) for channel in rgb]

        return "".join(f"{max(0, min(channel, 255)):02X}" for channel in rgb)
    except Exception:
        return hex_color


def _resolve_theme_name(theme_name: str, theme_colors: dict[str, str]) -> str | None:
    normalized_theme = theme_name.lower()
    resolved = theme_colors.get(normalized_theme)
    if resolved is not None:
        return resolved
    alias = THEME_ALIASES.get(normalized_theme)
    if alias is not None:
        return theme_colors.get(alias)
    return None


def _resolve_color(color_element: etree._Element | None, theme_colors: dict[str, str]) -> str:
    if color_element is None:
        return "#000000"

    value = color_element.get(f"{{{NS['w']}}}val")
    theme_color = color_element.get(f"{{{NS['w']}}}themeColor")
    theme_tint = color_element.get(f"{{{NS['w']}}}themeTint")
    theme_shade = color_element.get(f"{{{NS['w']}}}themeShade")

    if value and value != "auto":
        return f"#{value}"

    if theme_color:
        normalized_theme = theme_color.lower()
        resolved = theme_colors.get(normalized_theme)
        if resolved is None:
            resolved = theme_colors.get(THEME_ALIASES.get(normalized_theme, ""), "777777")
        resolved = _apply_tint_or_shade(resolved, theme_tint, theme_shade)
        return f"#{resolved}"

    return "#000000"


def _resolve_optional_color(color_element: etree._Element | None, theme_colors: dict[str, str]) -> str | None:
    if color_element is None:
        return None
    return _resolve_color(color_element, theme_colors)


def _resolve_word_shading(shading_element: etree._Element | None, theme_colors: dict[str, str]) -> str | None:
    if shading_element is None:
        return None

    fill = shading_element.get(f"{{{NS['w']}}}fill")
    theme_fill = shading_element.get(f"{{{NS['w']}}}themeFill")
    theme_tint = shading_element.get(f"{{{NS['w']}}}themeFillTint")
    theme_shade = shading_element.get(f"{{{NS['w']}}}themeFillShade")

    if fill and fill.lower() not in {"auto", "clear", "nil"}:
        return f"#{fill}"

    if theme_fill:
        resolved = _resolve_theme_name(theme_fill, theme_colors)
        if resolved is not None:
            resolved = _apply_tint_or_shade(resolved, theme_tint, theme_shade)
            return f"#{resolved}"

    return None


def _parse_style_context(temp_dir: Path, theme_colors: dict[str, str]) -> StyleContext:
    styles_path = temp_dir / "word" / "styles.xml"
    if not styles_path.exists():
        return StyleContext(styles={}, default_color="#000000", default_shading=None)

    try:
        tree = etree.parse(str(styles_path))
    except Exception:
        return StyleContext(styles={}, default_color="#000000", default_shading=None)

    styles: dict[str, StyleInfo] = {}
    default_color = "#000000"
    default_shading = None

    default_color_el = tree.find(".//w:docDefaults/w:rPrDefault/w:rPr/w:color", NS)
    if default_color_el is not None:
        default_color = _resolve_color(default_color_el, theme_colors)

    default_shading = _resolve_word_shading(tree.find(".//w:docDefaults/w:pPrDefault/w:pPr/w:shd", NS), theme_colors)
    if default_shading is None:
        default_shading = _resolve_word_shading(tree.find(".//w:docDefaults/w:rPrDefault/w:rPr/w:shd", NS), theme_colors)

    for style in tree.xpath("//w:style", namespaces=NS):
        style_id = style.get(f"{{{NS['w']}}}styleId")
        if not style_id:
            continue

        based_on_el = style.find("w:basedOn", NS)
        based_on = based_on_el.get(f"{{{NS['w']}}}val") if based_on_el is not None else None

        run_properties = style.find("w:rPr", NS)
        color = None
        shading = None
        if run_properties is not None:
            color = _resolve_optional_color(run_properties.find("w:color", NS), theme_colors)
            shading = _resolve_word_shading(run_properties.find("w:shd", NS), theme_colors)

        style_properties = style.find("w:pPr", NS)
        if shading is None and style_properties is not None:
            shading = _resolve_word_shading(style_properties.find("w:shd", NS), theme_colors)

        table_properties = style.find("w:tblPr", NS)
        if shading is None and table_properties is not None:
            shading = _resolve_word_shading(table_properties.find("w:shd", NS), theme_colors)

        styles[style_id] = StyleInfo(
            style_type=style.get(f"{{{NS['w']}}}type"),
            based_on=based_on,
            color=color,
            shading=shading,
        )

    return StyleContext(styles=styles, default_color=default_color, default_shading=default_shading)


def _resolve_style_color(style_id: str | None, style_context: StyleContext, fallback: str, seen: set[str] | None = None) -> str:
    if not style_id:
        return fallback
    if seen is None:
        seen = set()
    if style_id in seen:
        return fallback
    seen.add(style_id)

    style = style_context.styles.get(style_id)
    if style is None:
        return fallback
    if style.color is not None:
        return style.color
    return _resolve_style_color(style.based_on, style_context, fallback, seen)


def _resolve_style_shading(style_id: str | None, style_context: StyleContext, seen: set[str] | None = None) -> str | None:
    if not style_id:
        return None
    if seen is None:
        seen = set()
    if style_id in seen:
        return None
    seen.add(style_id)

    style = style_context.styles.get(style_id)
    if style is None:
        return None
    if style.shading is not None:
        return style.shading
    return _resolve_style_shading(style.based_on, style_context, seen)


def _apply_drawing_transform(value: int, transform_name: str, transform_value: int) -> int:
    if transform_name in {"tint", "lumOff"}:
        return round(value + (255 - value) * (transform_value / 100000))
    if transform_name in {"shade", "lumMod"}:
        return round(value * (transform_value / 100000))
    return value


def _resolve_drawing_color_choice(parent: etree._Element, theme_colors: dict[str, str]) -> str | None:
    for color_name in ["srgbClr", "schemeClr", "prstClr", "sysClr"]:
        color_element = parent.find(f"a:{color_name}", EXTRA_NS)
        if color_element is None:
            continue

        if color_name == "srgbClr":
            base_value = color_element.get("val")
        elif color_name == "schemeClr":
            base_value = _resolve_theme_name(color_element.get("val", ""), theme_colors)
        else:
            base_value = color_element.get("lastClr") or color_element.get("val")

        if not base_value:
            continue

        try:
            rgb = [int(base_value[index:index + 2], 16) for index in range(0, 6, 2)]
        except Exception:
            continue

        for transform in color_element:
            local_name = etree.QName(transform).localname
            raw_value = transform.get("val")
            if raw_value is None:
                continue
            try:
                transform_value = int(raw_value)
            except ValueError:
                continue
            rgb = [_apply_drawing_transform(channel, local_name, transform_value) for channel in rgb]

        return "#" + "".join(f"{max(0, min(channel, 255)):02X}" for channel in rgb)
    return None


def _resolve_highlight(run_properties: etree._Element | None) -> str | None:
    if run_properties is None:
        return None
    highlight = run_properties.find("w:highlight", NS)
    if highlight is None:
        return None

    value = highlight.get(f"{{{NS['w']}}}val")
    if not value:
        return None
    color = HIGHLIGHT_COLORS.get(value)
    if color is None:
        return None
    return f"#{color}"


def _resolve_drawing_fill(element: etree._Element, theme_colors: dict[str, str]) -> str | None:
    solid_fill = element.find(".//a:solidFill", EXTRA_NS)
    if solid_fill is None:
        return None
    return _resolve_drawing_color_choice(solid_fill, theme_colors)


def _resolve_vml_fill(element: etree._Element) -> str | None:
    fill_color = element.get("fillcolor")
    if fill_color:
        candidate = fill_color.strip().lstrip("#")
        if len(candidate) == 6:
            return f"#{candidate.upper()}"

    style = element.get("style")
    if style:
        for part in style.split(";"):
            name, _, value = part.partition(":")
            if name.strip().lower() != "fillcolor":
                continue
            candidate = value.strip().lstrip("#")
            if len(candidate) == 6:
                return f"#{candidate.upper()}"
    return None


def _background_from_ancestors(run: etree._Element, theme_colors: dict[str, str], style_context: StyleContext, default_background: str) -> str:
    ancestors = [run, *run.iterancestors()]

    run_properties = run.find("w:rPr", NS)
    highlight_background = _resolve_highlight(run_properties)
    if highlight_background is not None:
        return highlight_background

    run_shading = _resolve_word_shading(run_properties.find("w:shd", NS) if run_properties is not None else None, theme_colors)
    if run_shading is not None:
        return run_shading

    run_style = None
    if run_properties is not None:
        run_style_el = run_properties.find("w:rStyle", NS)
        if run_style_el is not None:
            run_style = run_style_el.get(f"{{{NS['w']}}}val")

    paragraph_style = None
    table_style = None

    for ancestor in ancestors:
        local_name = etree.QName(ancestor).localname

        if local_name == "tc":
            tc_shading = _resolve_word_shading(ancestor.find("w:tcPr/w:shd", NS), theme_colors)
            if tc_shading is not None:
                return tc_shading

        if local_name == "tr":
            tr_shading = _resolve_word_shading(ancestor.find("w:trPr/w:shd", NS), theme_colors)
            if tr_shading is not None:
                return tr_shading

        if local_name == "tbl":
            tbl_shading = _resolve_word_shading(ancestor.find("w:tblPr/w:shd", NS), theme_colors)
            if tbl_shading is not None:
                return tbl_shading
            tbl_style_el = ancestor.find("w:tblPr/w:tblStyle", NS)
            if tbl_style_el is not None and table_style is None:
                table_style = tbl_style_el.get(f"{{{NS['w']}}}val")

        if local_name == "p":
            paragraph_shading = _resolve_word_shading(ancestor.find("w:pPr/w:shd", NS), theme_colors)
            if paragraph_shading is not None:
                return paragraph_shading
            paragraph_style_el = ancestor.find("w:pPr/w:pStyle", NS)
            if paragraph_style_el is not None and paragraph_style is None:
                paragraph_style = paragraph_style_el.get(f"{{{NS['w']}}}val")

        vml_fill = _resolve_vml_fill(ancestor)
        if vml_fill is not None:
            return vml_fill

        drawing_fill = _resolve_drawing_fill(ancestor, theme_colors)
        if drawing_fill is not None:
            return drawing_fill

    for style_id in [run_style, paragraph_style, table_style]:
        shading = _resolve_style_shading(style_id, style_context)
        if shading is not None:
            return shading

    return style_context.default_shading or default_background


def _set_color(run_properties: etree._Element, hex_color: str) -> None:
    for element in run_properties.findall("w:color", NS):
        run_properties.remove(element)

    new_color = etree.Element(f"{{{NS['w']}}}color")
    new_color.set(f"{{{NS['w']}}}val", hex_color)
    run_properties.append(new_color)


def _effective_run_color(run: etree._Element, run_properties: etree._Element, theme_colors: dict[str, str], style_context: StyleContext) -> str:
    direct_color = _resolve_color(run_properties.find("w:color", NS), theme_colors)
    if run_properties.find("w:color", NS) is not None:
        return direct_color

    run_style_el = run_properties.find("w:rStyle", NS)
    if run_style_el is not None:
        style_id = run_style_el.get(f"{{{NS['w']}}}val")
        return _resolve_style_color(style_id, style_context, style_context.default_color)

    for ancestor in run.iterancestors():
        local_name = etree.QName(ancestor).localname
        if local_name == "p":
            paragraph_style_el = ancestor.find("w:pPr/w:pStyle", NS)
            if paragraph_style_el is not None:
                style_id = paragraph_style_el.get(f"{{{NS['w']}}}val")
                return _resolve_style_color(style_id, style_context, style_context.default_color)
        if local_name == "tbl":
            table_style_el = ancestor.find("w:tblPr/w:tblStyle", NS)
            if table_style_el is not None:
                style_id = table_style_el.get(f"{{{NS['w']}}}val")
                return _resolve_style_color(style_id, style_context, style_context.default_color)

    return style_context.default_color


def _fix_runs(root: etree._Element, background: str, theme_colors: dict[str, str], style_context: StyleContext) -> int:
    fixed = 0
    for run in root.xpath("//w:r", namespaces=NS):
        run_properties = run.find("w:rPr", NS)
        if run_properties is None:
            run_properties = etree.Element(f"{{{NS['w']}}}rPr")
            run.insert(0, run_properties)

        hex_color = _effective_run_color(run, run_properties, theme_colors, style_context)
        effective_background = _background_from_ancestors(run, theme_colors, style_context, background)

        if is_low_contrast(hex_color, effective_background):
            new_color = adjust_color_for_contrast(hex_color, effective_background)
            if is_low_contrast(f"#{new_color}", effective_background):
                new_color = adjust_color_for_contrast("#000000", effective_background)
            _set_color(run_properties, new_color)
            fixed += 1
    return fixed


def _fix_styles(root: etree._Element, background: str, theme_colors: dict[str, str], style_context: StyleContext) -> int:
    fixed = 0
    for style in root.xpath("//w:style", namespaces=NS):
        style_id = style.get(f"{{{NS['w']}}}styleId")
        style_background = _resolve_style_shading(style_id, style_context) or background

        for run_properties in style.xpath(".//w:rPr", namespaces=NS):
            color_element = run_properties.find("w:color", NS)
            if color_element is None:
                continue
            hex_color = _resolve_color(color_element, theme_colors)

            if is_low_contrast(hex_color, style_background):
                new_color = adjust_color_for_contrast(hex_color, style_background)
                if is_low_contrast(f"#{new_color}", style_background):
                    new_color = adjust_color_for_contrast("#000000", style_background)
                _set_color(run_properties, new_color)
                fixed += 1
    return fixed


def _iter_content_parts(temp_dir: Path) -> list[Path]:
    word_dir = temp_dir / "word"
    parts = [word_dir / "document.xml"]
    parts.extend(sorted(word_dir.glob("header*.xml")))
    parts.extend(sorted(word_dir.glob("footer*.xml")))

    for name in ["footnotes.xml", "endnotes.xml", "comments.xml", "commentsExtended.xml", "commentsIds.xml"]:
        parts.append(word_dir / name)

    return [part for part in parts if part.exists()]


def _iter_style_parts(temp_dir: Path) -> list[Path]:
    word_dir = temp_dir / "word"
    candidates = [word_dir / "styles.xml", word_dir / "stylesWithEffects.xml"]
    return [part for part in candidates if part.exists()]


def build_color_correction_output_name(filename: str) -> str:
    path = Path(filename or "document.docx")
    return f"{path.stem}_color_corrected{path.suffix or '.docx'}"


def process_docx_bytes(filename: str, file_bytes: bytes, background: str = "#FFFFFF") -> ProcessedDocument:
    background = normalize_hex_color(background)
    temp_dir = Path(tempfile.mkdtemp(prefix="colorizer_"))

    try:
        source_bytes = BytesIO(file_bytes)
        with zipfile.ZipFile(source_bytes, "r") as archive:
            archive.extractall(temp_dir)

        total_fixed = 0
        theme_colors = _parse_theme_colors(temp_dir)
        style_context = _parse_style_context(temp_dir, theme_colors)

        for xml_part in _iter_content_parts(temp_dir):
            tree = etree.parse(str(xml_part))
            runs_fixed = _fix_runs(tree.getroot(), background, theme_colors, style_context)
            total_fixed += runs_fixed
            if runs_fixed:
                tree.write(str(xml_part), encoding="UTF-8", xml_declaration=True)

        for style_part in _iter_style_parts(temp_dir):
            tree = etree.parse(str(style_part))
            styles_fixed = _fix_styles(tree.getroot(), background, theme_colors, style_context)
            total_fixed += styles_fixed
            if styles_fixed:
                tree.write(str(style_part), encoding="UTF-8", xml_declaration=True)

        output_buffer = BytesIO()
        with zipfile.ZipFile(output_buffer, "w", zipfile.ZIP_DEFLATED) as archive:
            for folder, _, files in os.walk(temp_dir):
                for name in files:
                    full_path = Path(folder) / name
                    archive.write(full_path, full_path.relative_to(temp_dir))

        return ProcessedDocument(
            filename=filename,
            output_bytes=output_buffer.getvalue(),
            fixed_elements=total_fixed,
            changed=total_fixed > 0,
        )
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
