import json
import os
import re
import tempfile
import tkinter as tk
from io import BytesIO
from dataclasses import dataclass
from tkinter import filedialog, messagebox, ttk
from typing import Callable

import pandas as pd
import requests
import urllib3
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    ArrayObject,
    DictionaryObject,
    FloatObject,
    NameObject,
    NumberObject,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

LOOKUP_URL = "http://prints.spudnik.local/api/prints/format-paths"
EXCLUDED_ITEMS = {"HA0814", "HA0815", "HA0816", "HA1129", "HA0817", "984398"}

APP_FONT = "Arial"
MONO_FONT = "Arial"
PDF_FONT = "Helvetica"
PDF_BOLD_FONT = "Helvetica-Bold"
CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".drawing_compiler_studio.json")
DEFAULT_NAMING_TEMPLATES = {
    "manual_packet": "{base}_packet",
    "automated_packet": "{base}_automated_packet",
    "structure_export": "{base}_structure",
    "structure_editor": "{base}_reordered",
    "drawing_download_folder": "{base}_drawings",
}



def _register_arial_pdf_font():
    global PDF_FONT, PDF_BOLD_FONT
    candidates = [
        ("/usr/share/fonts/truetype/msttcorefonts/Arial.ttf", "/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf"),
        ("/usr/share/fonts/truetype/msttcorefonts/arial.ttf", "/usr/share/fonts/truetype/msttcorefonts/arialbd.ttf"),
        ("/usr/share/fonts/truetype/microsoft/arial.ttf", "/usr/share/fonts/truetype/microsoft/arialbd.ttf"),
        ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
    ]
    for regular, bold in candidates:
        if os.path.exists(regular):
            pdfmetrics.registerFont(TTFont("Arial", regular))
            PDF_FONT = "Arial"
            if os.path.exists(bold):
                pdfmetrics.registerFont(TTFont("Arial-Bold", bold))
                PDF_BOLD_FONT = "Arial-Bold"
            return


_register_arial_pdf_font()

# ─── Palette ────────────────────────────────────────────────────────────────
DARK_PALETTE = {
    "bg":          "#0E1117",   # page background
    "surface":     "#161B24",   # sidebar / panel
    "card":        "#1C2333",   # card / raised surface
    "card_hover":  "#222B3E",   # card hover
    "border":      "#2A3347",   # standard border
    "border_hi":   "#3D4F6E",   # highlighted border
    "accent":      "#3B82F6",   # primary blue
    "accent_dim":  "#1D4ED8",   # deeper blue
    "accent_muted":"#1E3A5F",   # accent bg (subtle)
    "green":       "#10B981",
    "green_muted": "#064E3B",
    "amber":       "#F59E0B",
    "amber_muted": "#451A03",
    "rose":        "#F43F5E",
    "rose_muted":  "#4C0519",
    "violet":      "#8B5CF6",
    "violet_muted":"#2E1065",
    "text":        "#F1F5F9",   # primary text
    "text_dim":    "#94A3B8",   # secondary text
    "text_muted":  "#475569",   # tertiary text
    "tag_bg":      "#1E293B",
    "tag_text":    "#64748B",
    "sel_bg":      "#1E3A5F",
    "sel_text":    "#93C5FD",
    "danger":      "#EF4444",
    "danger_muted":"#450A0A",
}

LIGHT_PALETTE = {
    **DARK_PALETTE,
    "bg":          "#F8FAFC",
    "surface":     "#FFFFFF",
    "card":        "#FFFFFF",
    "card_hover":  "#EEF2FF",
    "border":      "#CBD5E1",
    "border_hi":   "#94A3B8",
    "accent_muted":"#DBEAFE",
    "green_muted": "#D1FAE5",
    "amber_muted": "#FEF3C7",
    "rose_muted":  "#FFE4E6",
    "violet_muted":"#EDE9FE",
    "text":        "#0F172A",
    "text_dim":    "#334155",
    "text_muted":  "#64748B",
    "tag_bg":      "#E2E8F0",
    "tag_text":    "#475569",
    "sel_bg":      "#DBEAFE",
    "sel_text":    "#1D4ED8",
    "danger_muted":"#FEE2E2",
}

C = DARK_PALETTE.copy()

WORKFLOW_META = {
    "manual_packet":    {"color": C["accent"],  "muted": C["accent_muted"]},
    "automated_packet": {"color": C["green"],   "muted": C["green_muted"]},
    "cad_to_structure": {"color": C["amber"],   "muted": C["amber_muted"]},
    "reorder_structure":{"color": C["violet"],  "muted": C["violet_muted"]},
    "reference_download":{"color": C["rose"],   "muted": C["rose_muted"]},
    "settings":        {"color": C["accent"],  "muted": C["accent_muted"]},
}


@dataclass(frozen=True)
class Workflow:
    key: str
    title: str
    subtitle: str


WORKFLOWS = [
    Workflow("dashboard",          "Dashboard",             "Unified control center"),
    Workflow("automated_packet",   "Automated Packet",      "Download + build in one flow"),
    Workflow("reorder_structure",  "Structure Editor",      "Reorder and renumber levels"),
    Workflow("reference_download", "Drawing Downloader",    "Download drawing references"),
    Workflow("settings",           "Settings",              "Preferences and defaults"),
]


# ─── Pure-logic helpers (unchanged) ─────────────────────────────────────────

def normalize_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).strip().lower())


def find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    normalized_map = {normalize_header(col): col for col in df.columns}
    for candidate in candidates:
        key = normalize_header(candidate)
        if key in normalized_map:
            return normalized_map[key]
    return None


def parse_level_code(level: str) -> tuple:
    text = str(level).strip()
    if not text:
        return tuple()
    output = []
    for token in text.split("."):
        token = token.strip()
        if not token:
            continue
        output.append(int(token) if token.isdigit() else token)
    return tuple(output)


def build_hierarchy(entries: list[dict]) -> list[dict]:
    seen_codes: dict[tuple, int] = {}
    processed: list[dict] = []
    for entry in entries:
        code = entry["code_tuple"]
        parent_index = None
        search = code[:-1]
        while search:
            if search in seen_codes:
                parent_index = seen_codes[search]
                break
            search = search[:-1]
        indent_level = 0 if parent_index is None else processed[parent_index]["indent_level"] + 1
        new_entry = dict(entry)
        new_entry["parent_index"] = parent_index
        new_entry["indent_level"] = indent_level
        processed.append(new_entry)
        seen_codes[code] = len(processed) - 1
    return processed


def is_hydraulic_schematic_entry(description: str) -> bool:
    return str(description).strip().upper().startswith("HYDRAULIC SCHEMATIC")


def _build_index_entries(entries: list[dict]) -> list[dict]:
    grouped: dict[str, dict] = {}
    for entry in entries:
        key = entry["desc"].strip().casefold()
        if key not in grouped:
            grouped[key] = {
                "desc": entry["desc"].strip(),
                "part": "",
                "item_numbers": [],
                "indent_level": 0,
                "toc_indices": [],
            }
        part = str(entry.get("part") or "").strip()
        if part and not grouped[key]["part"]:
            grouped[key]["part"] = part
        item_number = str(entry.get("item_number") or "").strip()
        if item_number and not is_hydraulic_schematic_entry(entry["desc"]) and item_number not in grouped[key]["item_numbers"]:
            grouped[key]["item_numbers"].append(item_number)
        grouped[key]["toc_indices"].append(entry["toc_index"])
    return sorted(grouped.values(), key=lambda e: e["desc"].casefold())


def _layout_directory_entries(entries, is_index=False, desc_font_name=PDF_FONT, desc_font_size=8):
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib.units import inch
    page_size = landscape(letter)
    width, height = page_size
    margin_left = 0.55 * inch
    margin_right = 0.55 * inch
    margin_top = 0.75 * inch
    margin_bottom = 0.6 * inch
    column_gap = 0.45 * inch
    indent_step = 0.22 * inch
    row_height = 0.21 * inch
    usable_width = width - margin_left - margin_right
    column_width = (usable_width - column_gap) / 2
    column_lefts = [margin_left, margin_left + column_width + column_gap]
    rows_per_column = max(1, int((height - margin_top - margin_bottom - (0.52 * inch)) / row_height))
    rows_per_page = rows_per_column * 2
    placements = []
    row_cursor = 0
    for idx, entry in enumerate(entries):
        while True:
            per_page_row = row_cursor % rows_per_page
            page_index = row_cursor // rows_per_page
            col_index = per_page_row // rows_per_column
            row_index = per_page_row % rows_per_column
            column_left = column_lefts[col_index]
            column_right = column_left + column_width
            title_y = height - margin_top
            y = title_y - 0.45 * inch - row_index * row_height
            desc_x = column_left + (0 if is_index else entry["indent_level"] * indent_step)
            page_x = column_right - 4
            page_column_width = (1.35 * inch) if is_index else (0.62 * inch)
            item_column_width = 0.9 * inch if is_index else 1.05 * inch
            page_left_x = column_right - page_column_width
            item_x = page_left_x - 8
            item_left_x = item_x - item_column_width
            desc_right_limit = item_left_x - 8
            desc_max_width = max(0, desc_right_limit - desc_x)
            desc_lines = _wrap_text_to_width(entry["desc"], desc_font_name, desc_font_size, desc_max_width)
            part_lines = []
            if is_index:
                part_text = str(entry.get("part") or "").strip()
                if part_text:
                    part_lines = _wrap_text_to_width(part_text, desc_font_name, desc_font_size, item_column_width)
            row_span = max(len(desc_lines), len(part_lines), 1)
            if (row_cursor % rows_per_column) + row_span > rows_per_column:
                row_cursor += rows_per_column - (row_cursor % rows_per_column)
                continue
            break
        placements.append({
            "entry_index": idx, "page_index": page_index,
            "desc_x": desc_x, "item_x": item_x, "item_left_x": item_left_x,
            "page_x": page_x, "page_left_x": page_left_x, "y": y,
            "title_y": title_y, "row_span": row_span,
            "desc_lines": desc_lines, "part_lines": part_lines,
        })
        row_cursor += row_span
    total_rows = row_cursor if row_cursor > 0 else 1
    total_pages = (total_rows + rows_per_page - 1) // rows_per_page
    return page_size, placements, total_pages


def _trim_text_to_width(text, font_name, font_size, max_width):
    if max_width <= 0:
        return ""
    if pdfmetrics.stringWidth(text, font_name, font_size) <= max_width:
        return text
    ellipsis = "..."
    ellipsis_width = pdfmetrics.stringWidth(ellipsis, font_name, font_size)
    available = max_width - ellipsis_width
    if available <= 0:
        return ellipsis
    trimmed = text
    while trimmed and pdfmetrics.stringWidth(trimmed, font_name, font_size) > available:
        trimmed = trimmed[:-1]
    return f"{trimmed}{ellipsis}"


def _wrap_text_to_width(text, font_name, font_size, max_width):
    text = str(text or "")
    if max_width <= 0:
        return [text]
    if not text:
        return [""]
    lines = []
    current = ""
    for word in text.split():
        candidate = word if not current else f"{current} {word}"
        if pdfmetrics.stringWidth(candidate, font_name, font_size) <= max_width:
            current = candidate
            continue
        if current:
            lines.append(current)
            current = ""
        if pdfmetrics.stringWidth(word, font_name, font_size) <= max_width:
            current = word
            continue
        split_word = word
        while split_word:
            chunk = split_word
            while chunk and pdfmetrics.stringWidth(chunk, font_name, font_size) > max_width:
                chunk = chunk[:-1]
            if not chunk:
                lines.append(split_word[:1])
                split_word = split_word[1:]
            else:
                lines.append(chunk)
                split_word = split_word[len(chunk):]
    if current:
        lines.append(current)
    return lines or [text]


def create_directory_pdf_bytes(entries, title, page_offset_map=None, is_index=False):
    packet = BytesIO()
    page_size, placements, _ = _layout_directory_entries(entries, is_index=is_index, desc_font_name=PDF_FONT, desc_font_size=8)
    c = canvas.Canvas(packet, pagesize=page_size)
    width, _ = page_size

    def draw_header(title_y):
        c.setFont(PDF_BOLD_FONT, 16)
        c.drawString(40, title_y, title)
        c.setLineWidth(0.5)
        c.line(40, title_y - 6, width - 40, title_y - 6)
        c.setFont(PDF_FONT, 10)

    current_page = -1
    for placement in placements:
        entry = entries[placement["entry_index"]]
        if placement["page_index"] != current_page:
            if current_page != -1:
                c.showPage()
            current_page = placement["page_index"]
            draw_header(placement["title_y"])
        desc = entry["desc"]
        item_number = str(entry.get("item_number") or "").strip()
        if is_hydraulic_schematic_entry(desc):
            display_item = ""
        elif is_index:
            display_item = str(entry.get("part") or "").strip()
        else:
            display_item = item_number
        entry_index = placement["entry_index"]
        page_num = ""
        if "page_text" in entry and entry["page_text"]:
            page_num = entry["page_text"]
        elif page_offset_map is not None and page_offset_map[entry_index] is not None:
            page_num = str(page_offset_map[entry_index] + 1)
        c.setFont(PDF_FONT, 8)
        desc_line_height = 9
        for line_index, desc_line in enumerate(placement.get("desc_lines", [desc])):
            c.drawString(placement["desc_x"], placement["y"] - (line_index * desc_line_height), desc_line)
        if display_item:
            if is_index:
                item_line_height = 9
                for line_index, item_line in enumerate(placement.get("part_lines", [display_item])):
                    c.drawRightString(placement["item_x"], placement["y"] - (line_index * item_line_height), item_line)
            else:
                item_text = _trim_text_to_width(display_item, PDF_FONT, 8, placement["item_x"] - placement["item_left_x"])
                c.drawRightString(placement["item_x"], placement["y"], item_text)
        if page_num:
            c.drawRightString(placement["page_x"], placement["y"], page_num)
    if current_page == -1:
        draw_header(page_size[1] - (0.75 * 72))
    c.save()
    packet.seek(0)
    return packet, placements


def _add_internal_link_annotation(writer, from_page_index, target_page_index, rect):
    target_ref = writer.pages[target_page_index].indirect_reference
    annotation = DictionaryObject()
    annotation.update({
        NameObject("/Type"): NameObject("/Annot"),
        NameObject("/Subtype"): NameObject("/Link"),
        NameObject("/Rect"): ArrayObject([FloatObject(x) for x in rect]),
        NameObject("/Border"): ArrayObject([NumberObject(0), NumberObject(0), NumberObject(0)]),
        NameObject("/A"): DictionaryObject({
            NameObject("/S"): NameObject("/GoTo"),
            NameObject("/D"): ArrayObject([target_ref, NameObject("/Fit")]),
        }),
    })
    writer.add_annotation(from_page_index, annotation)


def add_toc_hyperlinks(writer, toc_placements, effective_page_map, line_height=12):
    for placement in toc_placements:
        target_page = effective_page_map[placement["entry_index"]]
        if target_page is None:
            continue
        _add_internal_link_annotation(
            writer=writer, from_page_index=placement["page_index"],
            target_page_index=target_page,
            rect=[placement["desc_x"] - 2, placement["y"] - 1, placement["page_x"] + 1, placement["y"] + line_height],
        )


def add_page_number_overlay(page, page_num_text, total_pages_text):
    width = float(page.mediabox.width)
    height = float(page.mediabox.height)
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(width, height))
    c.setFont(PDF_FONT, 8)
    c.drawCentredString(width / 2, 5, f"{page_num_text} / {total_pages_text}")
    c.save()
    packet.seek(0)
    overlay_pdf = PdfReader(packet)
    page.merge_page(overlay_pdf.pages[0])


def _clean_cell(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    match = re.fullmatch(r"(\d+)\.0+", text)
    if match:
        return match.group(1)
    return text


def get_indent_level(object_value) -> int:
    text = "" if object_value is None else str(object_value)
    leading_spaces = len(text) - len(text.lstrip(" "))
    return leading_spaces // 4


def is_valid_item_number(item_number) -> bool:
    text = "" if item_number is None else str(item_number).strip().upper()
    if not text.startswith(("13", "FB", "HA")):
        return False
    return text not in EXCLUDED_ITEMS


def is_skippable_nonpart_row(object_value, name_value) -> bool:
    obj = "" if object_value is None else str(object_value).strip().upper()
    name = "" if name_value is None else str(name_value).strip().upper()
    return obj in {"SECTIONS", "CONSTRAINTS"} or name in {"SECTIONS", "CONSTRAINTS"}


def read_cad_export(path: str) -> pd.DataFrame:
    _, ext = os.path.splitext(path.lower())
    if ext in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(path)
    if ext == ".csv":
        return pd.read_csv(path)
    raise ValueError("Unsupported input file type. Use .xlsx/.xlsm/.xls/.csv")


def classify_workbook(path: str) -> str:
    """Return structure or cad_export based on recognizable columns."""
    df = read_cad_export(path)
    has_structure = all(find_column(df, [name]) for name in ["Level", "Description", "Part Number"])
    if has_structure:
        return "structure"
    has_cad = all([
        find_column(df, ["Object"]),
        find_column(df, ["Name"]),
        find_column(df, ["Item Number", "ItemNumber", "Item No", "Item"]),
    ])
    if has_cad:
        return "cad_export"
    raise ValueError("File is neither a supported structure workbook nor a CAD export.")


def cad_export_to_structure_dataframe(input_path: str) -> tuple[pd.DataFrame, dict]:
    df = read_cad_export(input_path)
    object_col = find_column(df, ["Object"])
    name_col = find_column(df, ["Name"])
    item_col = find_column(df, ["Item Number", "ItemNumber", "Item No", "Item"])
    if not object_col or not name_col or not item_col:
        raise ValueError("Required columns not found: Object, Name, Item Number")
    rows = []
    for source_index, (_, row) in enumerate(df.iterrows()):
        item_number = _clean_cell(row[item_col])
        description = _clean_cell(row[name_col])
        object_value = "" if pd.isna(row[object_col]) else str(row[object_col])
        indent = get_indent_level(object_value)
        rows.append({
            "source_index": source_index, "indent": indent,
            "Description": description, "Part Number": item_number,
            "keep": False, "direct_match": is_valid_item_number(item_number),
            "skippable": is_skippable_nonpart_row(object_value, description),
        })
    keep_stack = []
    for row in rows:
        while keep_stack and keep_stack[-1]["indent"] >= row["indent"]:
            keep_stack.pop()
        if row["direct_match"]:
            row["keep"] = True
            for ancestor in keep_stack:
                if not ancestor["skippable"] and ancestor["Description"]:
                    ancestor["keep"] = True
        keep_stack.append(row)
    filtered = [row for row in rows if row["keep"]]
    if not filtered:
        raise ValueError("No matching rows found after filters.")
    counters: dict[int, int] = {}
    output_rows = []
    for row in filtered:
        indent = max(0, row["indent"])
        for k in list(counters.keys()):
            if k > indent:
                counters.pop(k, None)
        counters[indent] = counters.get(indent, 0) + 1
        parts = [str(counters.get(i, 1)) for i in range(indent + 1)]
        output_rows.append({"Level": ".".join(parts), "Description": row["Description"], "Part Number": row["Part Number"]})
    out_df = pd.DataFrame(output_rows, columns=["Level", "Description", "Part Number"])
    return out_df, {
        "source_rows": len(df), "rows_written": len(out_df),
        "mapping": {"object_col": object_col, "name_col": name_col, "item_number_col": item_col},
    }


def combine_structure_frames(frames: list[pd.DataFrame]) -> pd.DataFrame:
    models = [StructureModel.from_dataframe(frame) for frame in frames]
    combined = StructureModel()
    for model in models:
        for child in model.root.children:
            combined.root.add_child(child)
    return combined.to_dataframe()


def convert_cad_to_structure(input_path: str, output_path: str) -> dict:
    out_df, details = cad_export_to_structure_dataframe(input_path)
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    out_df.to_excel(output_path, index=False)
    return {
        "input_path": input_path, "output_path": output_path,
        "source_rows": details["source_rows"], "rows_written": details["rows_written"],
        "mapping": details["mapping"],
    }


def convert_cad_exports_to_structure(input_paths: list[str], output_path: str) -> dict:
    if not input_paths:
        raise ValueError("Choose at least one CAD export file.")
    if len({os.path.abspath(path) for path in input_paths}) != len(input_paths):
        raise ValueError("Duplicate CAD export files are not allowed.")
    frames = []
    source_rows = 0
    mappings = []
    for path in input_paths:
        frame, details = cad_export_to_structure_dataframe(path)
        frames.append(frame)
        source_rows += details["source_rows"]
        mappings.append({"input_path": path, **details["mapping"]})
    out_df = frames[0] if len(frames) == 1 else combine_structure_frames(frames)
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    out_df.to_excel(output_path, index=False)
    return {
        "input_path": "; ".join(input_paths), "input_paths": input_paths, "output_path": output_path,
        "source_rows": source_rows, "rows_written": len(out_df), "mappings": mappings,
        "mapping": mappings[0] if mappings else {},
    }


def read_structure_references(structure_path: str) -> tuple[list[str], list[str]]:
    df = read_cad_export(structure_path)
    part_col = find_column(df, ["Part Number", "Item Number", "Part", "Item"])
    url_col = find_column(df, ["File URL", "Url", "PDF URL", "Link", "Path"])
    if not part_col and not url_col:
        raise ValueError("Missing supported columns. Need Part Number and/or URL columns.")
    part_numbers: set[str] = set()
    urls: set[str] = set()
    if part_col:
        for value in df[part_col].tolist():
            if pd.isna(value):
                continue
            text = str(value).strip()
            if text:
                part_numbers.add(text)
    if url_col:
        for value in df[url_col].tolist():
            if pd.isna(value):
                continue
            text = str(value).strip()
            if text.lower().startswith(("http://", "https://")):
                urls.add(text)
    return sorted(part_numbers), sorted(urls)


def lookup_print_paths(session, part_numbers):
    if not part_numbers:
        return {}, []
    payload = {"items": part_numbers, "location": "current"}
    headers = {
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "Origin": "http://prints.spudnik.local",
        "Referer": "http://prints.spudnik.local/",
        "User-Agent": "Mozilla/5.0",
    }
    response = session.post(LOOKUP_URL, json=payload, headers=headers, timeout=60, verify=False)
    response.raise_for_status()
    data = response.json()
    found = {}
    for entry in data.get("paths", []):
        part = str(entry.get("item", "")).strip()
        url = str(entry.get("path", "")).strip()
        if part and url:
            found[part] = url
    missing = [str(v).strip() for v in data.get("notFound", []) if str(v).strip()]
    return found, sorted(set(missing))


def download_url(session, url, out_path):
    response = session.get(url, timeout=90, verify=False)
    response.raise_for_status()
    with open(out_path, "wb") as f:
        f.write(response.content)


def download_references(structure_path, output_folder, progress_callback=None):
    os.makedirs(output_folder, exist_ok=True)
    part_numbers, direct_urls = read_structure_references(structure_path)
    session = requests.Session()
    found_paths, missing_parts = lookup_print_paths(session, part_numbers)
    downloaded = []
    skipped = []
    failed = []
    total_downloads = len(found_paths) + len(direct_urls)
    completed_downloads = 0
    if progress_callback:
        progress_callback(completed_downloads, total_downloads, "Preparing downloads...")
    for part, url in found_paths.items():
        target = os.path.join(output_folder, f"{part}.pdf")
        if os.path.exists(target):
            skipped.append(part)
        else:
            try:
                download_url(session, url, target)
                downloaded.append(part)
            except Exception:
                failed.append(part)
        completed_downloads += 1
        if progress_callback:
            progress_callback(completed_downloads, total_downloads, f"Downloaded: {part}")
    for url in direct_urls:
        filename = os.path.basename(url.split("?", 1)[0]) or "downloaded_file"
        target = os.path.join(output_folder, filename)
        if os.path.exists(target):
            skipped.append(url)
        else:
            try:
                download_url(session, url, target)
                downloaded.append(url)
            except Exception:
                failed.append(url)
        completed_downloads += 1
        if progress_callback:
            progress_callback(completed_downloads, total_downloads, f"Downloaded URL: {filename}")
    return {"downloaded": downloaded, "skipped": skipped, "missing_parts": missing_parts, "failed": failed, "output_folder": output_folder}


def _find_pdf_for_part(folder, part_number):
    exact = os.path.join(folder, f"{part_number}.pdf")
    if os.path.exists(exact):
        return exact
    lower_part = part_number.lower()
    for name in os.listdir(folder):
        if not name.lower().endswith(".pdf"):
            continue
        if lower_part in name.lower():
            return os.path.join(folder, name)
    return None


def build_manual_packet(structure_path, drawings_folder, output_pdf, schematic_pdf=None, progress_callback=None):
    df = pd.read_excel(structure_path)
    level_col = find_column(df, ["Level"])
    desc_col = find_column(df, ["Description", "Name"])
    part_col = find_column(df, ["Part Number", "Item Number", "Part", "Item"])
    if not level_col or not desc_col or not part_col:
        raise ValueError("Structure file must include Level, Description, and Part Number columns")
    raw_entries = []
    for _, row in df.iterrows():
        level = "" if pd.isna(row[level_col]) else str(row[level_col]).strip()
        desc = "" if pd.isna(row[desc_col]) else str(row[desc_col]).strip()
        part = "" if pd.isna(row[part_col]) else str(row[part_col]).strip()
        if not level:
            continue
        raw_entries.append({
            "code_text": level, "code_tuple": parse_level_code(level),
            "desc": desc, "part": part, "item_number": part,
            "filename": f"{part}.pdf" if part else "",
        })
    raw_entries.sort(key=lambda entry: entry["code_tuple"])
    toc_entries = build_hierarchy(raw_entries)
    existing_entries: list[dict] = []
    missing_files: list[str] = []
    if schematic_pdf and os.path.exists(schematic_pdf):
        existing_entries.append({
            "code_text": "0", "code_tuple": (0,),
            "desc": "HYDRAULIC SCHEMATIC", "part": "", "item_number": "",
            "filename": os.path.basename(schematic_pdf),
            "parent_index": None, "indent_level": 0,
            "_source_path": schematic_pdf,
        })
    if progress_callback:
        progress_callback(0, len(toc_entries), "Scanning structure entries...")
    for idx, entry in enumerate(toc_entries, start=1):
        part = entry["part"]
        if not part:
            continue
        pdf_path = _find_pdf_for_part(drawings_folder, part)
        if not pdf_path:
            missing_files.append(f"{part}.pdf")
            if progress_callback:
                progress_callback(idx, len(toc_entries), f"Missing drawing for {part}")
            continue
        new_entry = dict(entry)
        new_entry["_source_path"] = pdf_path
        existing_entries.append(new_entry)
        if progress_callback:
            progress_callback(idx, len(toc_entries), f"Queued {part}")
    if not existing_entries:
        raise ValueError("No PDFs were added. Check your drawings folder and part numbers.")
    for toc_index, entry in enumerate(existing_entries):
        entry["toc_index"] = toc_index
    index_entries = _build_index_entries(existing_entries)
    toc_entries = existing_entries + [{"desc": "Index", "part": "", "item_number": "", "indent_level": 0}]
    toc_packet, _ = create_directory_pdf_bytes(toc_entries, "Table of Contents")
    toc_pages = len(PdfReader(toc_packet).pages)
    page_offset_map = []
    current_page = toc_pages
    for entry in existing_entries:
        reader = PdfReader(entry["_source_path"])
        page_offset_map.append(current_page)
        current_page += len(reader.pages)
    index_start_page = current_page
    toc_page_map = page_offset_map + [index_start_page]
    toc_packet, toc_placements = create_directory_pdf_bytes(toc_entries, "Table of Contents", toc_page_map)
    for entry in index_entries:
        pages = []
        for toc_index in entry["toc_indices"]:
            page = page_offset_map[toc_index]
            if page is not None and page not in pages:
                pages.append(page)
        entry["page_text"] = ", ".join(str(page + 1) for page in pages)
    index_packet, _ = create_directory_pdf_bytes(index_entries, "Index", is_index=True)
    writer = PdfWriter()
    for page in PdfReader(toc_packet).pages:
        writer.add_page(page)
    entry_start_page = {}
    for entry in existing_entries:
        reader = PdfReader(entry["_source_path"])
        start_page = len(writer.pages)
        entry_start_page[id(entry)] = start_page
        for page in reader.pages:
            writer.add_page(page)
    for page in PdfReader(index_packet).pages:
        writer.add_page(page)
    bookmark_refs = {}
    for entry in existing_entries:
        start_page = entry_start_page[id(entry)]
        parent = None
        parent_index = entry["parent_index"]
        if parent_index is not None:
            parent_entry = toc_entries[parent_index]
            if id(parent_entry) in entry_start_page:
                parent = bookmark_refs.get(id(parent_entry))
        bookmark_refs[id(entry)] = writer.add_outline_item(entry["desc"], start_page, parent=parent)
    add_toc_hyperlinks(writer, toc_placements, toc_page_map)
    total_pages = len(writer.pages)
    for i, page in enumerate(writer.pages):
        add_page_number_overlay(page, i + 1, total_pages)
    os.makedirs(os.path.dirname(output_pdf) or ".", exist_ok=True)
    with open(output_pdf, "wb") as f:
        writer.write(f)
    return {
        "output_pdf": output_pdf,
        "included_parts": len(existing_entries),
        "missing_parts": missing_files,
        "index_entries": len(index_entries),
    }


def split_path_list(paths_text: str | list[str]) -> list[str]:
    if isinstance(paths_text, list):
        return [p for p in paths_text if p]
    return [p.strip() for p in str(paths_text).split(";") if p.strip()]


def build_automated_packet(input_paths, schematic_pdf, temp_download_folder, output_pdf, progress_callback=None):
    paths = split_path_list(input_paths)
    if not paths:
        raise ValueError("Choose a CAD export or structure workbook.")
    if len({os.path.abspath(path) for path in paths}) != len(paths):
        raise ValueError("Duplicate input files are not allowed.")
    file_types = [classify_workbook(path) for path in paths]
    if "structure" in file_types and len(paths) > 1:
        raise ValueError("Use one pre-created structure workbook, or one or more CAD export files; do not mix them.")
    if file_types[0] == "structure":
        structure_path = paths[0]
        source_type = "structure"
    else:
        structure_path = default_output_path(paths[0], "_structure", ".xlsx")
        convert_cad_exports_to_structure(paths, structure_path)
        source_type = "cad_export"

    def phase_download(completed, total, message):
        if progress_callback:
            progress_callback(completed, total if total > 0 else 1, f"[Download] {message}")

    download_result = download_references(structure_path, temp_download_folder, progress_callback=phase_download)

    def phase_build(completed, total, message):
        if progress_callback:
            progress_callback(completed, total if total > 0 else 1, f"[Build] {message}")

    packet_result = build_manual_packet(
        structure_path, temp_download_folder, output_pdf,
        schematic_pdf=schematic_pdf, progress_callback=phase_build,
    )
    return {
        "output_pdf": packet_result["output_pdf"],
        "structure_path": structure_path,
        "source_type": source_type,
        "included_parts": packet_result["included_parts"],
        "missing_parts": packet_result["missing_parts"],
        "downloaded": download_result["downloaded"],
        "skipped_downloads": download_result["skipped"],
        "failed_downloads": download_result["failed"],
        "not_found": download_result["missing_parts"],
    }


def load_structure_for_reorder(path: str) -> pd.DataFrame:
    df = read_cad_export(path)
    level_col = find_column(df, ["Level"])
    desc_col = find_column(df, ["Description"])
    part_col = find_column(df, ["Part Number"])
    if not level_col or not desc_col or not part_col:
        raise ValueError("File must contain Level, Description, and Part Number columns")
    out = pd.DataFrame({
        "Level": df[level_col].fillna("").astype(str),
        "Description": df[desc_col].fillna("").astype(str),
        "Part Number": df[part_col].fillna("").astype(str),
    })
    out = out[out["Level"].str.strip() != ""].reset_index(drop=True)
    return out


def renumber_structure(df: pd.DataFrame) -> pd.DataFrame:
    counters: dict[int, int] = {}
    output = []
    for _, row in df.iterrows():
        depth = max(0, str(row["Level"]).count("."))
        for key in list(counters.keys()):
            if key > depth:
                counters.pop(key, None)
        counters[depth] = counters.get(depth, 0) + 1
        new_level = ".".join(str(counters.get(i, 1)) for i in range(depth + 1))
        output.append({"Level": new_level, "Description": row["Description"], "Part Number": row["Part Number"]})
    return pd.DataFrame(output, columns=["Level", "Description", "Part Number"])


@dataclass
class StructureNode:
    level: str
    description: str
    part_number: str
    children: list["StructureNode"]
    parent: "StructureNode | None" = None

    def __init__(self, level, description, part_number):
        self.level = level
        self.description = description
        self.part_number = part_number
        self.children = []
        self.parent = None

    def add_child(self, child):
        child.parent = self
        self.children.append(child)

    def is_ancestor_of(self, other):
        current = other.parent
        while current is not None:
            if current is self:
                return True
            current = current.parent
        return False


class StructureModel:
    def __init__(self):
        self.root = StructureNode(level="", description="ROOT", part_number="")

    @classmethod
    def from_dataframe(cls, df):
        required = ["Level", "Description", "Part Number"]
        for col in required:
            if col not in df.columns:
                raise ValueError(f"Missing required column: {col}")
        model = cls()
        by_code = {tuple(): model.root}
        rows = []
        for _, row in df.iterrows():
            level = "" if pd.isna(row["Level"]) else str(row["Level"]).strip()
            if not level:
                continue
            rows.append({
                "level": level, "code": parse_level_code(level),
                "description": "" if pd.isna(row["Description"]) else str(row["Description"]).strip(),
                "part": "" if pd.isna(row["Part Number"]) else str(row["Part Number"]).strip(),
            })
        rows.sort(key=lambda r: r["code"])
        for row in rows:
            code = row["code"]
            search = code[:-1]
            parent = None
            while search:
                if search in by_code:
                    parent = by_code[search]
                    break
                search = search[:-1]
            if parent is None:
                parent = model.root
            node = StructureNode(level=row["level"], description=row["description"], part_number=row["part"])
            parent.add_child(node)
            by_code[code] = node
        if not model.root.children:
            raise ValueError("No valid rows found in the structure file.")
        return model

    def to_dataframe(self):
        rows = []

        def walk(nodes, prefix):
            for idx, node in enumerate(nodes, start=1):
                level = ".".join(str(x) for x in (prefix + [idx]))
                rows.append({"Level": level, "Description": node.description, "Part Number": node.part_number})
                walk(node.children, prefix + [idx])

        walk(self.root.children, [])
        return pd.DataFrame(rows, columns=["Level", "Description", "Part Number"])


def render_naming_template(template, input_path, extension=""):
    base = os.path.splitext(os.path.basename(input_path))[0] if input_path else "output"
    safe_template = (template or "{base}").strip() or "{base}"
    try:
        name = safe_template.format(base=base, ext=extension.lstrip("."), extension=extension)
    except (KeyError, IndexError, ValueError):
        name = f"{base}_{safe_template}"
    name = name.strip() or base
    if extension and not name.lower().endswith(extension.lower()):
        name = f"{name}{extension}"
    return name


def default_output_path(input_path, suffix, extension, template=None):
    if not input_path:
        return ""
    name = render_naming_template(template or f"{{base}}{suffix}", input_path, extension)
    return os.path.join(os.path.dirname(input_path), name)


def summarize_list(values, limit=10):
    if not values:
        return "None"
    shown = values[:limit]
    remaining = len(values) - len(shown)
    lines = [f"  • {item}" for item in shown]
    if remaining > 0:
        lines.append(f"  ... and {remaining} more")
    return "\n".join(lines)


def validate_output_filename(filename):
    invalid_chars_pattern = r'[<>:"/\\|?*]'
    reserved_names = {
        "CON","PRN","AUX","NUL",
        "COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8","COM9",
        "LPT1","LPT2","LPT3","LPT4","LPT5","LPT6","LPT7","LPT8","LPT9",
    }
    name = os.path.basename(filename).strip()
    if not name:
        raise ValueError("Output file name cannot be blank.")
    if re.search(invalid_chars_pattern, name):
        raise ValueError('Output file name contains invalid characters: <>:"/\\|?*')
    if name.endswith(" ") or name.endswith("."):
        raise ValueError("Output file name cannot end with a space or period.")
    base_name = os.path.splitext(name)[0].upper()
    if base_name in reserved_names:
        raise ValueError(f'"{base_name}" is a reserved Windows file name.')
    return name


# ─── UI Helpers ──────────────────────────────────────────────────────────────

def hex_to_rgb(hex_color):
    h = hex_color.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


def make_hover_color(base_hex, lighten=20):
    r, g, b = hex_to_rgb(base_hex)
    r = min(255, r + lighten)
    g = min(255, g + lighten)
    b = min(255, b + lighten)
    return f"#{r:02x}{g:02x}{b:02x}"


# ─── Progress Dialog ──────────────────────────────────────────────────────────

class ProgressDialog:
    def __init__(self, parent, title):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.resizable(False, False)
        self.window.transient(parent)
        self.window.grab_set()
        self.window.configure(bg=C["bg"])

        outer = tk.Frame(self.window, bg=C["card"], bd=0, highlightthickness=1, highlightbackground=C["border"])
        outer.pack(fill="both", expand=True, padx=1, pady=1)

        frame = tk.Frame(outer, bg=C["card"], padx=24, pady=20)
        frame.pack(fill="both", expand=True)

        tk.Label(
            frame, text=title.upper(),
            bg=C["card"], fg=C["text_muted"],
            font=(MONO_FONT, 9, "bold"), anchor="w",
        ).pack(fill="x", pady=(0, 12))

        self.status_var = tk.StringVar(value="Starting…")
        tk.Label(
            frame, textvariable=self.status_var,
            bg=C["card"], fg=C["text_dim"],
            font=(APP_FONT, 10), width=52, anchor="w", wraplength=380,
        ).pack(anchor="w", pady=(0, 10))

        bar_bg = tk.Frame(frame, bg=C["border"], height=4, bd=0)
        bar_bg.pack(fill="x", pady=(0, 4))
        bar_bg.pack_propagate(False)

        self.bar_fill = tk.Frame(bar_bg, bg=C["accent"], height=4)
        self.bar_fill.place(x=0, y=0, relheight=1.0, relwidth=0.0)
        self._bar_bg = bar_bg
        self._pct = 0

        self.pct_var = tk.StringVar(value="0%")
        tk.Label(
            frame, textvariable=self.pct_var,
            bg=C["card"], fg=C["text_muted"],
            font=(MONO_FONT, 9),
        ).pack(anchor="e")

        self.window.update_idletasks()
        w, h = 440, 160
        px = parent.winfo_x() + (parent.winfo_width() - w) // 2
        py = parent.winfo_y() + (parent.winfo_height() - h) // 2
        self.window.geometry(f"{w}x{h}+{px}+{py}")

    def update(self, completed, total, message):
        pct = (completed / total) if total else 0
        self.status_var.set(message)
        self.pct_var.set(f"{int(pct * 100)}%")
        self.bar_fill.place(relwidth=pct)
        self.window.update_idletasks()
        self.window.update()

    def close(self):
        if self.window.winfo_exists():
            self.window.destroy()


# ─── Main Application ─────────────────────────────────────────────────────────

class DrawingCompilerStudio(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Drawing Compiler Studio")
        self.geometry("1260x800")
        self.minsize(1100, 680)
        self.configure(bg=C["bg"])
        self.history: list[str] = []
        self.current = "dashboard"
        self._active_nav_key = "dashboard"
        self.reorder_model: StructureModel | None = None
        self.reorder_tree: ttk.Treeview | None = None
        self.reorder_source_path: str | None = None
        self.reorder_item_lookup: dict[str, StructureNode] = {}
        self.reorder_open_state: dict[tuple[str, str], bool] = {}
        self.reorder_loaded_files: set[str] = set()
        self.reorder_toolbar_buttons: dict[str, object] = {}
        self.reorder_undo_stack: list[tuple] = []
        self.reorder_drag_item: str | None = None
        self.reorder_drag_target_item: str | None = None
        self.reorder_drop_indicator: tk.Frame | None = None
        self.reorder_drag_status_var: tk.StringVar | None = None
        self.config_data = self._load_config()
        self.theme_name = self.config_data.get("theme", "dark")
        self.default_directory = self.config_data.get("default_directory") or ""
        self.naming_templates = self._load_naming_templates()
        self.last_folder = self.config_data.get("last_folder") or self.default_directory or os.getcwd()
        self._apply_theme_palette(self.theme_name)

        self._setup_styles()
        self._build_shell()
        self.show_workflow("dashboard", add_history=False)

    # ── Settings / file dialogs ──────────────────────────────────────────────

    def _load_config(self):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def _load_naming_templates(self):
        templates = DEFAULT_NAMING_TEMPLATES.copy()
        saved = self.config_data.get("naming_templates")
        if isinstance(saved, dict):
            for key, value in saved.items():
                if key in templates and isinstance(value, str) and value.strip():
                    templates[key] = value.strip()
        return templates

    def _save_config(self):
        self.config_data.update({
            "last_folder": self.last_folder,
            "default_directory": self.default_directory,
            "theme": self.theme_name,
            "naming_templates": self.naming_templates,
        })
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(self.config_data, f, indent=2)
        except Exception:
            pass

    def _apply_theme_palette(self, theme_name):
        palette = LIGHT_PALETTE if theme_name == "light" else DARK_PALETTE
        C.clear()
        C.update(palette)
        if hasattr(self, "theme_name"):
            self.theme_name = "light" if theme_name == "light" else "dark"
        for key, meta in WORKFLOW_META.items():
            if key == "manual_packet":
                meta.update({"color": C["accent"], "muted": C["accent_muted"]})
            elif key == "automated_packet":
                meta.update({"color": C["green"], "muted": C["green_muted"]})
            elif key == "cad_to_structure":
                meta.update({"color": C["amber"], "muted": C["amber_muted"]})
            elif key == "reorder_structure":
                meta.update({"color": C["violet"], "muted": C["violet_muted"]})
            elif key == "reference_download":
                meta.update({"color": C["rose"], "muted": C["rose_muted"]})
            elif key == "settings":
                meta.update({"color": C["accent"], "muted": C["accent_muted"]})

    def _set_theme(self, theme_name):
        self._apply_theme_palette(theme_name)
        self._save_config()
        self.configure(bg=C["bg"])
        self._setup_styles()
        for child in self.winfo_children():
            child.destroy()
        self._build_shell()
        self.show_workflow(self.current, add_history=False)

    def _remember_path(self, path):
        paths = path if isinstance(path, (list, tuple)) else [path]
        for value in paths:
            if not value:
                continue
            folder = value if os.path.isdir(value) else os.path.dirname(value)
            if folder:
                self.last_folder = folder
                self._save_config()
                break
        return path

    def _askopenfilename(self, **kwargs):
        kwargs.setdefault("initialdir", self.last_folder)
        return self._remember_path(filedialog.askopenfilename(**kwargs))

    def _askopenfilenames(self, **kwargs):
        kwargs.setdefault("initialdir", self.last_folder)
        return self._remember_path(list(filedialog.askopenfilenames(**kwargs)))

    def _askdirectory(self, **kwargs):
        kwargs.setdefault("initialdir", self.last_folder)
        return self._remember_path(filedialog.askdirectory(**kwargs))

    def _asksaveasfilename(self, **kwargs):
        kwargs.setdefault("initialdir", self.last_folder)
        return self._remember_path(filedialog.asksaveasfilename(**kwargs))

    def _bind_enabled_state(self, button, required_vars):
        def update(*_):
            ready = all(var.get().strip() for var in required_vars)
            button.configure(state=("normal" if ready else "disabled"), cursor=("hand2" if ready else "arrow"))
        for var in required_vars:
            var.trace_add("write", update)
        update()

    # ── Theming ───────────────────────────────────────────────────────────────

    def _setup_styles(self):
        style = ttk.Style(self)
        style.theme_use("default")

        # Treeview
        style.configure(
            "Reorder.Treeview",
            background=C["card"],
            foreground=C["text"],
            fieldbackground=C["card"],
            borderwidth=0,
            relief="flat",
            rowheight=28,
            font=(APP_FONT, 10),
        )
        style.configure(
            "Reorder.Treeview.Heading",
            background=C["surface"],
            foreground=C["text_muted"],
            borderwidth=0,
            relief="flat",
            font=(MONO_FONT, 9, "bold"),
            padding=(8, 6),
        )
        style.map(
            "Reorder.Treeview",
            background=[("selected", C["sel_bg"])],
            foreground=[("selected", C["sel_text"])],
        )
        style.map("Reorder.Treeview.Heading", background=[("active", C["border"])])

        # Scrollbar
        style.configure(
            "Dark.Vertical.TScrollbar",
            troughcolor=C["card"],
            background=C["border"],
            borderwidth=0,
            relief="flat",
            arrowsize=0,
        )
        style.map("Dark.Vertical.TScrollbar", background=[("active", C["border_hi"])])
        style.configure(
            "Dark.Horizontal.TScrollbar",
            troughcolor=C["card"],
            background=C["border"],
            borderwidth=0,
            relief="flat",
            arrowsize=0,
        )
        style.map("Dark.Horizontal.TScrollbar", background=[("active", C["border_hi"])])

        # Progressbar (for any ttk use)
        style.configure(
            "Accent.Horizontal.TProgressbar",
            troughcolor=C["border"],
            background=C["accent"],
            borderwidth=0,
            thickness=4,
        )

    def _show_themed_dialog(self, title: str, message: str, tone: str = "info"):
        tone_map = {
            "info": (C["accent"], C["accent_muted"], "OK"),
            "warning": (C["amber"], C["amber_muted"], "Got it"),
            "error": (C["danger"], C["danger_muted"], "Close"),
        }
        accent, muted, button_text = tone_map.get(tone, tone_map["info"])
        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.configure(bg=C["bg"])

        outer = tk.Frame(dialog, bg=C["card"], bd=0, highlightthickness=1, highlightbackground=C["border"])
        outer.pack(fill="both", expand=True, padx=1, pady=1)
        frame = tk.Frame(outer, bg=C["card"], padx=20, pady=16)
        frame.pack(fill="both", expand=True)

        tk.Label(
            frame,
            text=title.upper(),
            bg=C["card"],
            fg=accent,
            font=(MONO_FONT, 10, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(0, 8))

        tk.Label(
            frame,
            text=message,
            justify="left",
            anchor="w",
            bg=C["card"],
            fg=C["text_dim"],
            font=(APP_FONT, 10),
            wraplength=560,
        ).pack(fill="x", pady=(0, 14))

        btn = tk.Button(
            frame,
            text=button_text,
            bg=muted,
            fg=accent,
            activebackground=accent,
            activeforeground="#FFFFFF",
            font=(APP_FONT, 10, "bold"),
            bd=0,
            padx=18,
            pady=8,
            cursor="hand2",
            command=dialog.destroy,
        )
        btn.pack(anchor="e")

        dialog.update_idletasks()
        w = min(640, max(460, dialog.winfo_reqwidth() + 20))
        h = min(440, max(170, dialog.winfo_reqheight() + 12))
        px = self.winfo_x() + (self.winfo_width() - w) // 2
        py = self.winfo_y() + (self.winfo_height() - h) // 2
        dialog.geometry(f"{w}x{h}+{px}+{py}")
        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        dialog.wait_window(dialog)

    # ── Shell layout ──────────────────────────────────────────────────────────

    def _build_shell(self):
        self.sidebar_frame = tk.Frame(self, bg=C["surface"], width=220)
        self.sidebar_frame.pack(side="left", fill="y")
        self.sidebar_frame.pack_propagate(False)

        # Separator line between sidebar and main
        sep = tk.Frame(self, bg=C["border"], width=1)
        sep.pack(side="left", fill="y")

        self.main_frame = tk.Frame(self, bg=C["bg"])
        self.main_frame.pack(side="left", fill="both", expand=True)

        self._build_sidebar()

    def _build_sidebar(self):
        f = self.sidebar_frame

        # ── Logo block ──
        logo_block = tk.Frame(f, bg=C["surface"], pady=20, padx=18)
        logo_block.pack(fill="x")

        logo_accent = tk.Frame(logo_block, bg=C["accent"], width=3, height=32)
        logo_accent.pack(side="left", fill="y", padx=(0, 10))

        logo_text = tk.Frame(logo_block, bg=C["surface"])
        logo_text.pack(side="left")
        tk.Label(logo_text, text="DRAWING COMPILER", bg=C["surface"], fg=C["text"],
                 font=(MONO_FONT, 11, "bold")).pack(anchor="w")
        tk.Label(logo_text, text="Studio · v2", bg=C["surface"], fg=C["text_muted"],
                 font=(MONO_FONT, 8)).pack(anchor="w")

        # Divider
        tk.Frame(f, bg=C["border"], height=1).pack(fill="x", padx=16, pady=(4, 12))

        # ── Nav label ──
        tk.Label(f, text="WORKFLOWS", bg=C["surface"], fg=C["text_muted"],
                 font=(MONO_FONT, 8, "bold"), anchor="w",
                 padx=18).pack(fill="x", pady=(0, 6))

        # ── Nav buttons ──
        self.nav_buttons: dict[str, tk.Button] = {}
        icons = {
            "dashboard":           "⊞",
            "automated_packet":    "⚙",
            "reorder_structure":   "≡",
            "reference_download":  "↓",
            "settings":            "⚙",
        }
        for wf in WORKFLOWS:
            btn = self._nav_btn(f, icons.get(wf.key, "·"), wf.title, wf.key)
            btn.pack(fill="x", padx=8, pady=2)
            self.nav_buttons[wf.key] = btn

        # ── Spacer + back ──
        spacer = tk.Frame(f, bg=C["surface"])
        spacer.pack(fill="both", expand=True)

        tk.Frame(f, bg=C["border"], height=1).pack(fill="x", padx=16, pady=(0, 10))

        self.back_btn = tk.Button(
            f, text="◀  Back",
            bg=C["surface"], fg=C["text_muted"],
            activebackground=C["card"], activeforeground=C["text"],
            font=(APP_FONT, 10), bd=0, cursor="hand2",
            relief="flat", pady=8, command=self.go_back, state="disabled",
        )
        self.back_btn.pack(fill="x", padx=8, pady=(0, 14))

    def _nav_btn(self, parent, icon, label, key):
        color = WORKFLOW_META.get(key, {}).get("color", C["text_dim"])
        btn = tk.Button(
            parent,
            text=f"  {icon}  {label}",
            bg=C["surface"],
            fg=C["text_dim"],
            activebackground=C["card"],
            activeforeground=C["text"],
            font=(APP_FONT, 10),
            bd=0, relief="flat",
            cursor="hand2",
            anchor="w",
            padx=6, pady=8,
            command=lambda k=key: self.show_workflow(k),
        )

        def on_enter(e, b=btn, c=color):
            if b.cget("state") != "disabled":
                b.configure(bg=C["card"], fg=c)

        def on_leave(e, b=btn):
            if key == self._active_nav_key:
                return
            if b["bg"] != C["border_hi"]:
                b.configure(bg=C["surface"], fg=C["text_dim"])

        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        return btn

    def _set_active_nav(self, key):
        self._active_nav_key = key
        for k, btn in self.nav_buttons.items():
            color = WORKFLOW_META.get(k, {}).get("color", C["text_dim"])
            if k == key:
                btn.configure(
                    bg=C["card"],
                    fg=color,
                    font=(APP_FONT, 10, "bold"),
                )
            else:
                btn.configure(
                    bg=C["surface"],
                    fg=C["text_dim"],
                    font=(APP_FONT, 10),
                )

    # ── Navigation ────────────────────────────────────────────────────────────

    def _clear_main(self):
        for child in self.main_frame.winfo_children():
            child.destroy()
        self.reorder_tree = None
        self.reorder_drag_item = None
        self.reorder_drag_target_item = None
        self.reorder_drag_status_var = None

    def _push(self):
        self.history.append(self.current)
        self.back_btn.configure(state="normal")

    def go_back(self):
        if not self.history:
            return
        previous = self.history.pop()
        self.show_workflow(previous, add_history=False)
        self.back_btn.configure(state="normal" if self.history else "disabled")

    def show_workflow(self, key, add_history=True):
        if add_history and key != self.current:
            self._push()
        self.current = key
        self._clear_main()
        self._set_active_nav(key)

        dispatch = {
            "dashboard":           self._dashboard,
            "automated_packet":    self._automated_packet_page,
            "reorder_structure":   self._reorder_page,
            "reference_download":  self._reference_page,
            "settings":            self._settings_page,
        }
        dispatch.get(key, self._dashboard)()

    # ── Shared UI primitives ──────────────────────────────────────────────────

    def _scrollable_main(self):
        """Return a frame that scrolls vertically inside main_frame."""
        canvas = tk.Canvas(self.main_frame, bg=C["bg"], bd=0, highlightthickness=0)
        vsb = ttk.Scrollbar(self.main_frame, orient="vertical", command=canvas.yview,
                             style="Dark.Vertical.TScrollbar")
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        inner = tk.Frame(canvas, bg=C["bg"])
        window_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(window_id, width=canvas.winfo_width())

        inner.bind("<Configure>", _on_configure)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(window_id, width=e.width))
        def _on_mousewheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        return inner

    def _page_header(self, parent, title, subtitle, accent_color=None):
        color = accent_color or C["accent"]
        header = tk.Frame(parent, bg=C["bg"], pady=0)
        header.pack(fill="x", padx=36, pady=(28, 0))

        # Accent bar
        bar = tk.Frame(header, bg=color, width=3)
        bar.pack(side="left", fill="y", padx=(0, 14))

        text_block = tk.Frame(header, bg=C["bg"])
        text_block.pack(side="left", fill="x", expand=True)

        tk.Label(
            text_block, text=title,
            bg=C["bg"], fg=C["text"],
            font=(APP_FONT, 18, "bold"), anchor="w",
        ).pack(fill="x")
        tk.Label(
            text_block, text=subtitle,
            bg=C["bg"], fg=C["text_muted"],
            font=(APP_FONT, 10), anchor="w",
        ).pack(fill="x")

        tk.Frame(parent, bg=C["border"], height=1).pack(fill="x", padx=36, pady=(16, 0))

    def _card(self, parent, padx=36, pady=(20, 0)):
        wrapper = tk.Frame(parent, bg=C["bg"])
        wrapper.pack(fill="x", padx=padx, pady=pady)
        card = tk.Frame(
            wrapper, bg=C["card"],
            bd=0, highlightthickness=1,
            highlightbackground=C["border"],
        )
        card.pack(fill="x")
        inner = tk.Frame(card, bg=C["card"], padx=24, pady=20)
        inner.pack(fill="x")
        return inner

    def _section_label(self, parent, text):
        tk.Label(
            parent, text=text.upper(),
            bg=C["card"], fg=C["text_muted"],
            font=(MONO_FONT, 8, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(0, 8))

    def _field(self, parent, label, var, browse_cmd, browse_label="Browse…", optional=False):
        """Render a labeled path input row."""
        row_frame = tk.Frame(parent, bg=C["card"])
        row_frame.pack(fill="x", pady=(0, 14))

        lbl_row = tk.Frame(row_frame, bg=C["card"])
        lbl_row.pack(fill="x", pady=(0, 4))
        tk.Label(
            lbl_row, text=label,
            bg=C["card"], fg=C["text_dim"],
            font=(APP_FONT, 9), anchor="w",
        ).pack(side="left")
        if optional:
            tk.Label(
                lbl_row, text=" optional",
                bg=C["card"], fg=C["text_muted"],
                font=(MONO_FONT, 8), anchor="w",
            ).pack(side="left")

        input_row = tk.Frame(row_frame, bg=C["card"])
        input_row.pack(fill="x")

        entry = tk.Entry(
            input_row, textvariable=var,
            bg=C["bg"], fg=C["text"],
            insertbackground=C["text"],
            selectbackground=C["accent_muted"],
            selectforeground=C["text"],
            font=(MONO_FONT, 9),
            bd=0, highlightthickness=1,
            highlightbackground=C["border"],
            highlightcolor=C["accent"],
            relief="flat",
        )
        entry.pack(side="left", fill="x", expand=True, ipady=7, ipadx=8)

        browse = self._small_btn(input_row, browse_label, browse_cmd)
        browse.pack(side="left", padx=(8, 0))

    def _small_btn(self, parent, text, command, color=None):
        bg = color or C["border"]
        btn = tk.Button(
            parent, text=text, command=command,
            bg=bg, fg=C["text_dim"],
            activebackground=C["border_hi"], activeforeground=C["text"],
            font=(APP_FONT, 9), bd=0, relief="flat", cursor="hand2",
            padx=12, pady=6,
        )
        btn.bind("<Enter>", lambda e: btn.configure(bg=C["border_hi"], fg=C["text"]))
        btn.bind("<Leave>", lambda e: btn.configure(bg=bg, fg=C["text_dim"]))
        return btn

    def _run_btn(self, parent, text, command, color=None):
        bg = color or C["accent"]
        hover = make_hover_color(bg, 25)
        btn = tk.Button(
            parent, text=f"  {text}  ", command=command,
            bg=bg, fg="#FFFFFF",
            activebackground=hover, activeforeground="#FFFFFF",
            font=(APP_FONT, 10, "bold"), bd=0, relief="flat",
            cursor="hand2", padx=16, pady=9,
        )
        btn.bind("<Enter>", lambda e: btn.configure(bg=hover))
        btn.bind("<Leave>", lambda e: btn.configure(bg=bg))
        return btn

    def _divider(self, parent):
        tk.Frame(parent, bg=C["border"], height=1).pack(fill="x", pady=16)

    # ── Dashboard ─────────────────────────────────────────────────────────────

    def _dashboard(self):
        main = self._scrollable_main()
        self._page_header(main, "Drawing Compiler Studio", "Select a workflow to get started", C["accent"])

        grid_frame = tk.Frame(main, bg=C["bg"])
        grid_frame.pack(fill="x", padx=36, pady=24)

        cards_data = [wf for wf in WORKFLOWS if wf.key != "dashboard"]
        descriptions = {
            "automated_packet":    "Ingest a CAD export or structure workbook, download drawings, and build a full packet in one flow.",
            "reorder_structure":   "Full editor: add, reorder, edit, reparent, and remove structure rows, then save.",
            "reference_download":  "Download drawings from a saved structure, CAD export, or the current Structure Editor data.",
            "settings":            "Switch themes, choose the default browse folder, and configure default output naming templates.",
        }

        for i, wf in enumerate(cards_data):
            meta = WORKFLOW_META.get(wf.key, {})
            color = meta.get("color", C["accent"])
            muted = meta.get("muted", C["accent_muted"])

            col = i % 2
            row = i // 2

            card_outer = tk.Frame(grid_frame, bg=C["card"], bd=0,
                                   highlightthickness=1, highlightbackground=C["border"])
            card_outer.grid(row=row, column=col, sticky="nsew", padx=8, pady=8)
            grid_frame.columnconfigure(col, weight=1)

            card_inner = tk.Frame(card_outer, bg=C["card"], padx=22, pady=18)
            card_inner.pack(fill="both", expand=True)

            # Color accent strip at top
            tk.Frame(card_outer, bg=color, height=3).place(x=0, y=0, relwidth=1)

            # Title row
            title_row = tk.Frame(card_inner, bg=C["card"])
            title_row.pack(fill="x", pady=(6, 6))

            tk.Label(
                title_row, text=wf.title,
                bg=C["card"], fg=C["text"],
                font=(APP_FONT, 12, "bold"), anchor="w",
            ).pack(side="left")

            # Description
            tk.Label(
                card_inner,
                text=descriptions.get(wf.key, wf.subtitle),
                bg=C["card"], fg=C["text_muted"],
                font=(APP_FONT, 9),
                anchor="w", wraplength=400, justify="left",
            ).pack(fill="x", pady=(0, 14))

            # Open button
            open_btn = tk.Button(
                card_inner, text="Open →",
                bg=muted, fg=color,
                activebackground=color, activeforeground="#fff",
                font=(APP_FONT, 9, "bold"), bd=0, relief="flat",
                cursor="hand2", padx=12, pady=6,
                command=lambda k=wf.key: self.show_workflow(k),
            )
            open_btn.bind("<Enter>", lambda e, b=open_btn, c=color: b.configure(bg=c, fg="#fff"))
            open_btn.bind("<Leave>", lambda e, b=open_btn, c=color, m=muted: b.configure(bg=m, fg=c))
            open_btn.pack(anchor="w")



    # ── Manual Packet ─────────────────────────────────────────────────────────

    def _manual_packet_page(self):
        main = self._scrollable_main()
        color = WORKFLOW_META["manual_packet"]["color"]
        self._page_header(main, "Manual Packet Builder",
                          "Merge local drawing PDFs using a structure workbook. Optionally prepend a hydraulic schematic.", color)

        card = self._card(main)
        self._section_label(card, "Inputs")

        structure_var = tk.StringVar()
        drawings_var = tk.StringVar()
        schematic_var = tk.StringVar()
        output_var = tk.StringVar()

        self._field(card, "Structure workbook (.xlsx)", structure_var,
                    lambda: self._browse_file(structure_var, output_var, "_packet", ".pdf",
                                               [("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")],
                                               "manual_packet"))
        self._field(card, "Drawings folder", drawings_var,
                    lambda: drawings_var.set(self._askdirectory() or drawings_var.get()),
                    "Browse folder…")
        self._field(card, "Hydraulic schematic PDF", schematic_var,
                    lambda: schematic_var.set(
                        self._askopenfilename(filetypes=[("PDF", "*.pdf"), ("All", "*.*")]) or schematic_var.get()),
                    optional=True)

        self._divider(card)
        self._section_label(card, "Output")
        self._field(card, "Output PDF path", output_var,
                    lambda: output_var.set(
                        self._asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")]) or output_var.get()),
                    "Save as…")

        def run():
            if not structure_var.get() or not drawings_var.get() or not output_var.get():
                messagebox.showwarning("Missing fields", "Please fill in Structure workbook, Drawings folder, and Output PDF.", parent=self)
                return
            try:
                validate_output_filename(output_var.get())
            except ValueError as exc:
                messagebox.showerror("Invalid filename", str(exc), parent=self)
                return
            progress = ProgressDialog(self, "Building Manual Packet")
            result = None
            error = None
            try:
                result = build_manual_packet(
                    structure_var.get(), drawings_var.get(), output_var.get(),
                    schematic_pdf=schematic_var.get().strip() or None,
                    progress_callback=progress.update,
                )
            except Exception as exc:
                error = str(exc)
            finally:
                progress.close()
            if error:
                self._show_themed_dialog("Build failed", error, tone="error")
                return
            self._show_themed_dialog(
                "Packet complete",
                f"Output:  {result['output_pdf']}\n"
                f"Parts included:  {result['included_parts']}\n"
                f"Parts missing:   {len(result['missing_parts'])}\n\n"
                f"Missing:\n{summarize_list(result['missing_parts'])}",
                tone="info",
            )

        self._divider(card)
        run_btn = self._run_btn(card, "Build Manual Packet", run, color)
        run_btn.pack(anchor="w")
        self._bind_enabled_state(run_btn, [structure_var, drawings_var, output_var])

    # ── Automated Packet ──────────────────────────────────────────────────────

    def _automated_packet_page(self):
        main = self._scrollable_main()
        color = WORKFLOW_META["automated_packet"]["color"]
        self._page_header(main, "Automated Packet Builder",
                          "Parse a CAD export, download drawings, then build the full packet with TOC and index.", color)

        card = self._card(main)
        self._section_label(card, "Inputs")

        cad_var = tk.StringVar()
        schematic_var = tk.StringVar()
        download_var = tk.StringVar()
        output_var = tk.StringVar()

        self._field(card, "CAD export(s) or structure workbook", cad_var,
                    lambda: self._browse_files(cad_var, output_var, "_automated_packet", ".pdf",
                                               [("Supported", "*.xlsx *.xlsm *.xls *.csv"), ("All", "*.*")],
                                               "automated_packet"))
        self._field(card, "Hydraulic schematic PDF", schematic_var,
                    lambda: schematic_var.set(
                        self._askopenfilename(filetypes=[("PDF", "*.pdf"), ("All", "*.*")]) or schematic_var.get()))
        self._field(card, "Download folder", download_var,
                    lambda: download_var.set(self._askdirectory() or download_var.get()),
                    "Browse folder…")

        self._divider(card)
        self._section_label(card, "Output")
        self._field(card, "Output PDF path", output_var,
                    lambda: output_var.set(
                        self._asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")]) or output_var.get()),
                    "Save as…")

        def run():
            if not cad_var.get() or not schematic_var.get() or not download_var.get() or not output_var.get():
                messagebox.showwarning("Missing fields", "All four fields are required.", parent=self)
                return
            try:
                validate_output_filename(output_var.get())
            except ValueError as exc:
                messagebox.showerror("Invalid filename", str(exc), parent=self)
                return
            progress = ProgressDialog(self, "Running Automated Build")
            result = None
            error = None
            try:
                result = build_automated_packet(
                    cad_var.get(), schematic_var.get(),
                    download_var.get(), output_var.get(),
                    progress_callback=progress.update,
                )
            except Exception as exc:
                error = str(exc)
            finally:
                progress.close()
            if error:
                self._show_themed_dialog("Build failed", error, tone="error")
                return
            self._show_themed_dialog(
                "Build complete",
                f"Output:             {result['output_pdf']}\n"
                f"Source type:        {result['source_type']}\n"
                f"Structure file:     {result['structure_path']}\n"
                f"Parts included:     {result['included_parts']}\n"
                f"Downloads skipped:  {len(result['skipped_downloads'])}\n"
                f"Missing in packet:  {len(result['missing_parts'])}\n"
                f"Download failures:  {len(result['failed_downloads'])}\n"
                f"Not found in lookup:{len(result['not_found'])}\n\n"
                f"Missing in packet:\n{summarize_list(result['missing_parts'])}\n\n"
                f"Lookup not found:\n{summarize_list(result['not_found'])}",
                tone="info",
            )

        self._divider(card)
        run_btn = self._run_btn(card, "Run Automated Build", run, color)
        run_btn.pack(anchor="w")
        self._bind_enabled_state(run_btn, [cad_var, schematic_var, download_var, output_var])

    # ── CAD to Structure ──────────────────────────────────────────────────────

    def _cad_page(self):
        main = self._scrollable_main()
        color = WORKFLOW_META["cad_to_structure"]["color"]
        self._page_header(main, "CAD Export to Structure",
                          "Convert a CAD export spreadsheet into the Level / Description / Part Number structure format.", color)

        card = self._card(main)
        self._section_label(card, "Input")

        input_var = tk.StringVar()
        output_var = tk.StringVar()

        self._field(card, "CAD export file(s) (.xlsx / .xlsm / .xls / .csv)", input_var,
                    lambda: self._browse_files(input_var, output_var, "_structure", ".xlsx",
                                               [("Supported", "*.xlsx *.xlsm *.xls *.csv"), ("All", "*.*")],
                                               "structure_export"))

        self._divider(card)
        self._section_label(card, "Output")
        self._field(card, "Output structure workbook (.xlsx)", output_var,
                    lambda: output_var.set(
                        self._asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")]) or output_var.get()),
                    "Save as…")

        def run():
            if not input_var.get() or not output_var.get():
                messagebox.showwarning("Missing fields", "Choose input and output files.", parent=self)
                return
            try:
                result = convert_cad_exports_to_structure(split_path_list(input_var.get()), output_var.get())
                messagebox.showinfo("Conversion complete",
                    f"Input files:  {len(result['input_paths'])}\n"
                    f"Output:       {result['output_path']}\n\n"
                    f"Rows read:    {result['source_rows']}\n"
                    f"Rows written: {result['rows_written']}", parent=self)
            except Exception as exc:
                messagebox.showerror("Conversion failed", str(exc), parent=self)

        self._divider(card)
        run_btn = self._run_btn(card, "Generate Structure", run, color)
        run_btn.pack(anchor="w")
        self._bind_enabled_state(run_btn, [input_var, output_var])

    # ── Structure Reorder ─────────────────────────────────────────────────────

    def _reorder_page(self):
        color = WORKFLOW_META["reorder_structure"]["color"]

        # Header (non-scrollable area at top)
        header_frame = tk.Frame(self.main_frame, bg=C["bg"])
        header_frame.pack(fill="x")
        self._page_header(header_frame, "Structure Editor",
                          "Add, reorder, edit, reparent, and remove rows — then save a renumbered structure workbook.", color)

        # Full toolbar: all Structure Editor actions stay visible without an overflow menu.
        toolbar = tk.Frame(self.main_frame, bg=C["bg"], padx=36, pady=12)
        toolbar.pack(fill="x")

        def tool_btn(text, cmd, accent=False, danger=False):
            if danger:
                bg, fg, hover = C["danger_muted"], C["danger"], C["danger"]
            elif accent:
                bg, fg, hover = color, "#fff", make_hover_color(color, 20)
            else:
                bg, fg, hover = C["border"], C["text_dim"], C["text"]
            btn = tk.Button(
                toolbar, text=text, command=cmd,
                bg=bg, fg=fg,
                activebackground=hover, activeforeground="#fff" if (accent or danger) else C["text"],
                font=(APP_FONT, 9), bd=0, relief="flat", cursor="hand2",
                padx=12, pady=7,
            )
            btn.bind("<Enter>", lambda e: btn.configure(bg=hover, fg="#fff" if (accent or danger) else C["text"]))
            btn.bind("<Leave>", lambda e: btn.configure(bg=bg, fg=fg))
            btn.pack(side="left", padx=(0, 6))
            return btn

        self.reorder_toolbar_buttons.clear()
        self.reorder_toolbar_buttons["add_file"] = tool_btn("Add file", self._reorder_add_file, accent=True)
        self.reorder_toolbar_buttons["add_item"] = tool_btn("Add item", self._reorder_add_item)
        self.reorder_toolbar_buttons["edit"] = tool_btn("Edit", lambda: self._reorder_edit())
        self.reorder_toolbar_buttons["up"] = tool_btn("↑", lambda: self._reorder_move(-1))
        self.reorder_toolbar_buttons["down"] = tool_btn("↓", lambda: self._reorder_move(1))
        self.reorder_toolbar_buttons["make_child"] = tool_btn("Make child", self._reorder_make_child)
        self.reorder_toolbar_buttons["promote"] = tool_btn("Promote", self._reorder_promote)
        self.reorder_toolbar_buttons["expand"] = tool_btn("Expand all", lambda: self._reorder_expand(True))
        self.reorder_toolbar_buttons["collapse"] = tool_btn("Collapse all", lambda: self._reorder_expand(False))
        self.reorder_toolbar_buttons["undo"] = tool_btn("↶", lambda: self._reorder_undo())
        self.reorder_toolbar_buttons["remove"] = tool_btn("⌫", lambda: self._reorder_remove(), danger=True)
        self.reorder_toolbar_buttons["clear"] = tool_btn("Clear editor", self._reorder_clear_editor, danger=True)
        self.reorder_toolbar_buttons["save"] = tool_btn("Save as", self._reorder_save)

        self.reorder_drag_status_var = tk.StringVar(
            value="Tip: drag rows to reorder them, drop in the middle to make a child, or near the top/bottom edge to place before/after."
        )
        tk.Label(
            toolbar, textvariable=self.reorder_drag_status_var, bg=C["bg"], fg=C["text_muted"],
            font=(APP_FONT, 9), anchor="w",
        ).pack(side="left", fill="x", expand=True, padx=(16, 0))

        # Tree area
        tree_frame = tk.Frame(self.main_frame, bg=C["bg"], padx=36)
        tree_frame.pack(fill="both", expand=True, pady=(0, 20))

        tree_border = tk.Frame(tree_frame, bg=C["border"], bd=0)
        tree_border.pack(fill="both", expand=True)

        tree_inner = tk.Frame(tree_border, bg=C["card"])
        tree_inner.pack(fill="both", expand=True, padx=1, pady=1)

        self.reorder_tree = ttk.Treeview(
            tree_inner,
            columns=("part",),
            show="tree headings",
            selectmode="browse",
            style="Reorder.Treeview",
        )
        self.reorder_tree.heading("#0", text="DESCRIPTION", anchor="w")
        self.reorder_tree.heading("part", text="PART NUMBER", anchor="w")
        self.reorder_tree.column("#0", width=720, anchor="w", minwidth=300)
        self.reorder_tree.column("part", width=200, anchor="w", minwidth=100)

        scroll = ttk.Scrollbar(tree_inner, orient="vertical",
                               command=self.reorder_tree.yview,
                               style="Dark.Vertical.TScrollbar")
        self.reorder_tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.reorder_tree.pack(side="left", fill="both", expand=True)
        self.reorder_tree.bind("<<TreeviewSelect>>", lambda _e: self._reorder_update_button_states())
        self.reorder_tree.bind("<Delete>", lambda _e: self._reorder_remove())
        self.reorder_tree.bind("<Control-z>", lambda _e: self._reorder_undo())
        self.reorder_tree.bind("<Control-Z>", lambda _e: self._reorder_undo())
        self.reorder_tree.tag_configure("drop_child", background=C["border_hi"], foreground=C["text"])
        self.reorder_tree.bind("<Return>", lambda _e: self._reorder_edit())
        self.reorder_tree.bind("<ButtonPress-1>", self._reorder_begin_drag)
        self.reorder_tree.bind("<B1-Motion>", self._reorder_drag_motion)
        self.reorder_tree.bind("<ButtonRelease-1>", self._reorder_drop)
        self._reorder_update_button_states()

    def _reorder_open(self):
        path = self._askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")])
        if not path:
            return
        try:
            df = load_structure_for_reorder(path)
            self.reorder_model = StructureModel.from_dataframe(df)
            self.reorder_source_path = path
            self.reorder_loaded_files = {os.path.abspath(path)}
            self.reorder_open_state.clear()
            self.reorder_undo_stack.clear()
            if not self._reorder_tree_available():
                self.show_workflow("reorder_structure", add_history=False)
            self._reorder_refresh()
        except Exception as exc:
            messagebox.showerror("Open failed", str(exc), parent=self)

    def _reorder_add_file(self):
        paths = self._askopenfilenames(filetypes=[("Supported", "*.xlsx *.xlsm *.xls *.csv"), ("All", "*.*")])
        if not paths:
            return
        try:
            models = []
            for path in paths:
                kind = classify_workbook(path)
                if kind == "structure":
                    models.append(StructureModel.from_dataframe(load_structure_for_reorder(path)))
                else:
                    models.append(StructureModel.from_dataframe(cad_export_to_structure_dataframe(path)[0]))

            self._reorder_append_models(models, paths)
        except Exception as exc:
            messagebox.showerror("Add file failed", str(exc), parent=self)

    def _reorder_new_file_paths(self, paths, duplicate_title="Duplicate file"):
        duplicates = [path for path in paths if os.path.abspath(path) in self.reorder_loaded_files]
        if duplicates:
            messagebox.showwarning(
                duplicate_title,
                "These files are already loaded and were skipped:\n" + "\n".join(os.path.basename(p) for p in duplicates),
                parent=self,
            )
        return [path for path in paths if os.path.abspath(path) not in self.reorder_loaded_files]

    def _reorder_append_models(self, models, source_paths):
        if not models:
            return
        self._reorder_snapshot()
        if not self.reorder_model:
            self.reorder_model = StructureModel()
        for model in models:
            for child in model.root.children:
                self.reorder_model.root.add_child(child)
        self.reorder_loaded_files.update(os.path.abspath(path) for path in source_paths)
        if not self._reorder_tree_available():
            self.show_workflow("reorder_structure", add_history=False)
        self._reorder_refresh()

    def _reorder_load_structure_files(self):
        paths = self._askopenfilenames(filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")])
        if not paths:
            return
        new_paths = self._reorder_new_file_paths(paths, duplicate_title="Duplicate structure")
        if not new_paths:
            return
        try:
            models = [StructureModel.from_dataframe(load_structure_for_reorder(path)) for path in new_paths]
            self._reorder_append_models(models, new_paths)
        except Exception as exc:
            messagebox.showerror("Load structure failed", str(exc), parent=self)

    def _reorder_load_cad_exports(self):
        paths = self._askopenfilenames(filetypes=[("Supported", "*.xlsx *.xlsm *.xls *.csv"), ("All", "*.*")])
        if not paths:
            return
        new_paths = self._reorder_new_file_paths(paths, duplicate_title="Duplicate CAD export")
        if not new_paths:
            return
        try:
            models = [StructureModel.from_dataframe(cad_export_to_structure_dataframe(path)[0]) for path in new_paths]
            self._reorder_append_models(models, new_paths)
        except Exception as exc:
            messagebox.showerror("Load CAD export failed", str(exc), parent=self)

    def _reorder_save(self):
        if not self.reorder_model:
            messagebox.showwarning("No data", "Open a structure file first.", parent=self)
            return
        template_source = self.reorder_source_path or "structure.xlsx"
        initial = render_naming_template(self.naming_templates.get("structure_editor"), template_source, ".xlsx")
        path = self._asksaveasfilename(defaultextension=".xlsx", initialfile=initial, filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            self.reorder_model.to_dataframe().to_excel(path, index=False)
            messagebox.showinfo("Saved", f"Structure saved to:\n{path}", parent=self)
        except Exception as exc:
            messagebox.showerror("Save failed", str(exc), parent=self)

    def _node_state_key(self, node):
        return (str(id(node)), node.description, node.part_number)

    def _remember_reorder_open_state(self):
        if not self._reorder_tree_available():
            return
        for item_id, node in self.reorder_item_lookup.items():
            self.reorder_open_state[self._node_state_key(node)] = bool(self.reorder_tree.item(item_id, "open"))

    def _reorder_refresh(self, select_node=None):
        if not self._reorder_tree_available():
            return
        self._remember_reorder_open_state()
        for item in self.reorder_tree.get_children():
            self.reorder_tree.delete(item)
        self.reorder_item_lookup.clear()
        if not self.reorder_model:
            self._reorder_update_button_states()
            return

        def add_nodes(parent_id, nodes):
            for node in nodes:
                item_id = self.reorder_tree.insert(
                    parent_id, "end",
                    text=f"  {node.description}",
                    values=(node.part_number,),
                    open=self.reorder_open_state.get(self._node_state_key(node), False),
                )
                self.reorder_item_lookup[item_id] = node
                add_nodes(item_id, node.children)

        add_nodes("", self.reorder_model.root.children)

        if select_node:
            for item_id, node in self.reorder_item_lookup.items():
                if node is select_node:
                    self.reorder_tree.selection_set(item_id)
                    self.reorder_tree.focus(item_id)
                    break
        self._reorder_update_button_states()

    def _reorder_snapshot(self):
        if self.reorder_model and self.reorder_model.root.children:
            self.reorder_undo_stack.append(self.reorder_model.to_dataframe())

    def _reorder_update_button_states(self):
        if not self.reorder_toolbar_buttons:
            return
        has_model = self.reorder_model is not None
        node = self._reorder_selected_node() if self._reorder_tree_available() else None
        states = {"add_file": "normal"}
        for key in ["save", "add_item", "expand", "collapse", "clear"]:
            states[key] = "normal" if has_model else "disabled"
        for key in ["edit", "remove", "make_child", "promote", "up", "down"]:
            states[key] = "normal" if node else "disabled"
        states["undo"] = "normal" if self.reorder_undo_stack else "disabled"
        if node and node.parent:
            siblings = node.parent.children
            idx = siblings.index(node)
            states["up"] = "normal" if idx > 0 else "disabled"
            states["down"] = "normal" if idx < len(siblings) - 1 else "disabled"
            states["make_child"] = "normal" if len(self.reorder_item_lookup) > 1 else "disabled"
            states["promote"] = "normal" if node.parent and node.parent.parent else "disabled"
        for key, control in self.reorder_toolbar_buttons.items():
            state = states.get(key, "normal")
            self._set_reorder_action_state(control, state)

    def _set_reorder_action_state(self, control, state):
        if isinstance(control, tuple):
            menu, index = control
            try:
                menu.entryconfigure(index, state=state)
            except tk.TclError:
                pass
        else:
            control.configure(state=state, cursor=("hand2" if state == "normal" else "arrow"))

    def _reorder_selected_node(self):
        if not self._reorder_tree_available():
            return None
        sel = self.reorder_tree.selection()
        return self.reorder_item_lookup.get(sel[0]) if sel else None

    def _reorder_tree_available(self):
        return bool(self.reorder_tree and self.reorder_tree.winfo_exists())

    def _reorder_item_for_node(self, wanted_node):
        for item_id, node in self.reorder_item_lookup.items():
            if node is wanted_node:
                return item_id
        return ""

    def _reorder_clear_drop_target(self):
        if self.reorder_tree and self.reorder_drag_target_item:
            try:
                self.reorder_tree.item(self.reorder_drag_target_item, tags=())
            except tk.TclError:
                pass
        if self.reorder_drop_indicator and self.reorder_drop_indicator.winfo_exists():
            self.reorder_drop_indicator.place_forget()
        self.reorder_drag_target_item = None

    def _reorder_show_drop_indicator(self, target_item, mode):
        if not self._reorder_tree_available() or not target_item:
            return
        if mode == "inside":
            self.reorder_drag_target_item = target_item
            self.reorder_tree.item(target_item, tags=("drop_child",))
            return
        bbox = self.reorder_tree.bbox(target_item)
        if not bbox:
            return
        if self.reorder_drop_indicator is None or not self.reorder_drop_indicator.winfo_exists():
            self.reorder_drop_indicator = tk.Frame(self.reorder_tree, bg=C["accent"], height=2)
        x, y, width, height = bbox
        line_y = y if mode == "before" else y + height - 2
        self.reorder_drop_indicator.place(x=x, y=line_y, width=width, height=2)
        self.reorder_drag_target_item = target_item

    def _reorder_set_drag_status(self, message):
        if self.reorder_drag_status_var is not None:
            self.reorder_drag_status_var.set(message)

    def _reorder_begin_drag(self, event):
        if not self._reorder_tree_available():
            return
        if self.reorder_tree.identify_region(event.x, event.y) not in {"tree", "cell"}:
            return
        item_id = self.reorder_tree.identify_row(event.y)
        if not item_id or item_id not in self.reorder_item_lookup:
            return
        self.reorder_drag_item = item_id
        self.reorder_tree.selection_set(item_id)
        self.reorder_tree.focus(item_id)
        node = self.reorder_item_lookup[item_id]
        self._reorder_set_drag_status(f"Dragging '{node.description}'. Drop on a row to make it a child, or near an edge to place before/after.")
        self._reorder_update_button_states()

    def _reorder_drag_motion(self, event):
        if not self.reorder_drag_item or not self._reorder_tree_available():
            return
        target_item = self.reorder_tree.identify_row(event.y)
        plan = self._reorder_drop_plan(target_item, event.y)
        self._reorder_clear_drop_target()
        if not plan:
            self.reorder_tree.configure(cursor="X_cursor")
            self._reorder_set_drag_status("Drop not allowed here.")
            return
        _parent, _index, mode, target_node = plan
        if target_item:
            self._reorder_show_drop_indicator(target_item, mode)
        self.reorder_tree.configure(cursor="hand2")
        if mode == "top":
            message = "Drop to move to the bottom of the top level."
        elif mode == "inside":
            message = f"Drop to make it a child of '{target_node.description}'."
        elif mode == "before":
            message = f"Drop to place it before '{target_node.description}'."
        else:
            message = f"Drop to place it after '{target_node.description}'."
        self._reorder_set_drag_status(message)

    def _reorder_drop(self, event):
        if not self.reorder_drag_item or not self._reorder_tree_available():
            self.reorder_drag_item = None
            return
        target_item = self.reorder_tree.identify_row(event.y)
        plan = self._reorder_drop_plan(target_item, event.y)
        drag_item = self.reorder_drag_item
        self.reorder_drag_item = None
        self._reorder_clear_drop_target()
        self.reorder_tree.configure(cursor="")
        if not plan:
            self._reorder_set_drag_status("Drag canceled: that drop target is not valid.")
            return
        node = self.reorder_item_lookup.get(drag_item)
        if not node or not node.parent:
            self._reorder_set_drag_status("Drag canceled: selected row is no longer available.")
            return
        parent, index, mode, target_node = plan
        if parent is node.parent:
            old_index = node.parent.children.index(node)
            if old_index < index:
                index -= 1
            if old_index == index:
                self._reorder_set_drag_status("No change: item was dropped in its current position.")
                return
        self._reorder_snapshot()
        node.parent.children.remove(node)
        node.parent = parent
        parent.children.insert(index, node)
        self._reorder_refresh(node)
        if mode == "top":
            message = f"Moved '{node.description}' to the top level."
        elif mode == "inside":
            message = f"Moved '{node.description}' under '{target_node.description}'."
        elif mode == "before":
            message = f"Moved '{node.description}' before '{target_node.description}'."
        else:
            message = f"Moved '{node.description}' after '{target_node.description}'."
        self._reorder_set_drag_status(message)

    def _reorder_drop_plan(self, target_item, y):
        if not self.reorder_drag_item or not self.reorder_model:
            return None
        dragged_node = self.reorder_item_lookup.get(self.reorder_drag_item)
        if not dragged_node or not dragged_node.parent:
            return None
        if not target_item:
            return self.reorder_model.root, len(self.reorder_model.root.children), "top", self.reorder_model.root
        target_node = self.reorder_item_lookup.get(target_item)
        if not target_node:
            return None
        if target_node is dragged_node or dragged_node.is_ancestor_of(target_node):
            return None
        bbox = self.reorder_tree.bbox(target_item) if self._reorder_tree_available() else ""
        if not bbox:
            return None
        _x, row_y, _w, row_h = bbox
        relative_y = (y - row_y) / max(row_h, 1)
        if relative_y < 0.25:
            siblings = target_node.parent.children if target_node.parent else self.reorder_model.root.children
            return target_node.parent or self.reorder_model.root, siblings.index(target_node), "before", target_node
        if relative_y > 0.75:
            siblings = target_node.parent.children if target_node.parent else self.reorder_model.root.children
            return target_node.parent or self.reorder_model.root, siblings.index(target_node) + 1, "after", target_node
        return target_node, len(target_node.children), "inside", target_node

    def _reorder_prompt(self, title, desc="", part=""):
        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.configure(bg=C["bg"])
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)

        outer = tk.Frame(dialog, bg=C["card"], bd=0, highlightthickness=1, highlightbackground=C["border"])
        outer.pack(fill="both", expand=True, padx=1, pady=1)
        frame = tk.Frame(outer, bg=C["card"], padx=24, pady=20)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text=title.upper(), bg=C["card"], fg=C["text_muted"],
                 font=(MONO_FONT, 9, "bold")).pack(anchor="w", pady=(0, 14))

        tk.Label(frame, text="Description", bg=C["card"], fg=C["text_dim"],
                 font=(APP_FONT, 9)).pack(anchor="w")
        desc_var = tk.StringVar(value=desc)
        tk.Entry(frame, textvariable=desc_var, bg=C["bg"], fg=C["text"],
                 insertbackground=C["text"], font=(MONO_FONT, 10),
                 bd=0, highlightthickness=1, highlightbackground=C["border"],
                 highlightcolor=C["accent"], relief="flat", width=52
                 ).pack(fill="x", pady=(4, 12), ipady=6, ipadx=8)

        tk.Label(frame, text="Part Number", bg=C["card"], fg=C["text_dim"],
                 font=(APP_FONT, 9)).pack(anchor="w")
        part_var = tk.StringVar(value=part)
        tk.Entry(frame, textvariable=part_var, bg=C["bg"], fg=C["text"],
                 insertbackground=C["text"], font=(MONO_FONT, 10),
                 bd=0, highlightthickness=1, highlightbackground=C["border"],
                 highlightcolor=C["accent"], relief="flat", width=52
                 ).pack(fill="x", pady=(4, 0), ipady=6, ipadx=8)

        result = [None]

        def on_ok():
            d = desc_var.get().strip()
            if not d:
                messagebox.showwarning("Required", "Description cannot be blank.", parent=dialog)
                return
            result[0] = (d, part_var.get().strip())
            dialog.destroy()

        btns = tk.Frame(frame, bg=C["card"])
        btns.pack(anchor="e", pady=(16, 0))

        cancel = tk.Button(btns, text="Cancel", command=dialog.destroy,
                           bg=C["card"], fg=C["text_dim"],
                           activebackground=C["border"], activeforeground=C["text"],
                           font=(APP_FONT, 9), bd=0, relief="flat", cursor="hand2",
                           padx=12, pady=6)
        cancel.pack(side="right")

        ok_color = WORKFLOW_META["reorder_structure"]["color"]
        ok = self._run_btn(btns, "OK", on_ok, ok_color)
        ok.pack(side="right", padx=(0, 8))

        w, h = 440, 240
        dialog.geometry(f"{w}x{h}+{self.winfo_x() + (self.winfo_width()-w)//2}+{self.winfo_y() + (self.winfo_height()-h)//2}")
        self.wait_window(dialog)
        return result[0]

    def _reorder_parent_dialog(self, title, prompt, excluded_node=None, allow_top=True):
        if not self.reorder_model:
            return None
        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.configure(bg=C["bg"])
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(True, True)

        outer = tk.Frame(dialog, bg=C["card"], bd=0, highlightthickness=1, highlightbackground=C["border"])
        outer.pack(fill="both", expand=True, padx=1, pady=1)
        frame = tk.Frame(outer, bg=C["card"], padx=20, pady=16)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text=title.upper(), bg=C["card"], fg=C["text_muted"],
                 font=(MONO_FONT, 9, "bold")).pack(anchor="w", pady=(0, 8))
        tk.Label(frame, text=prompt, bg=C["card"], fg=C["text_dim"],
                 font=(APP_FONT, 10), wraplength=560, justify="left").pack(fill="x", pady=(0, 12))

        tree_frame = tk.Frame(frame, bg=C["border"], bd=0)
        tree_frame.pack(fill="both", expand=True)
        inner = tk.Frame(tree_frame, bg=C["card"])
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        parent_tree = ttk.Treeview(inner, columns=("part",), show="tree headings", selectmode="browse", style="Reorder.Treeview")
        parent_tree.heading("#0", text="PARENT", anchor="w")
        parent_tree.heading("part", text="PART NUMBER", anchor="w")
        parent_tree.column("#0", width=420, minwidth=220, anchor="w")
        parent_tree.column("part", width=160, minwidth=80, anchor="w")
        scroll = ttk.Scrollbar(inner, orient="vertical", command=parent_tree.yview, style="Dark.Vertical.TScrollbar")
        parent_tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        parent_tree.pack(side="left", fill="both", expand=True)

        lookup = {}
        if allow_top:
            top_id = parent_tree.insert("", "end", text="  Top level", values=("",), open=True)
            lookup[top_id] = self.reorder_model.root

        def allowed(candidate):
            if excluded_node is None:
                return True
            return candidate is not excluded_node and not excluded_node.is_ancestor_of(candidate)

        def add_nodes(parent_id, nodes):
            for node in nodes:
                if not allowed(node):
                    continue
                item_id = parent_tree.insert(parent_id, "end", text=f"  {node.description}", values=(node.part_number,), open=True)
                lookup[item_id] = node
                add_nodes(item_id, node.children)

        add_nodes("", self.reorder_model.root.children)
        if lookup:
            first = next(iter(lookup))
            parent_tree.selection_set(first)
            parent_tree.focus(first)

        result = [None]

        def on_ok():
            sel = parent_tree.selection()
            if not sel:
                messagebox.showwarning("Required", "Select a parent or Top level.", parent=dialog)
                return
            result[0] = lookup.get(sel[0])
            dialog.destroy()

        btns = tk.Frame(frame, bg=C["card"])
        btns.pack(anchor="e", pady=(14, 0))
        tk.Button(btns, text="Cancel", command=dialog.destroy, bg=C["card"], fg=C["text_dim"],
                  activebackground=C["border"], activeforeground=C["text"], font=(APP_FONT, 9),
                  bd=0, relief="flat", cursor="hand2", padx=12, pady=6).pack(side="right")
        ok = self._run_btn(btns, "OK", on_ok, WORKFLOW_META["reorder_structure"]["color"])
        ok.pack(side="right", padx=(0, 8))
        parent_tree.bind("<Double-1>", lambda _e: on_ok())

        w, h = 640, 520
        dialog.geometry(f"{w}x{h}+{self.winfo_x() + (self.winfo_width()-w)//2}+{self.winfo_y() + (self.winfo_height()-h)//2}")
        self.wait_window(dialog)
        return result[0]

    def _reorder_add_item(self):
        if not self.reorder_model:
            return
        values = self._reorder_prompt("Add Item")
        if not values:
            return
        parent = self._reorder_parent_dialog(
            "Select Parent",
            "Select the parent for the new item, or choose Top level to add it as a top-level assembly.",
            allow_top=True,
        )
        if parent is None:
            return
        self._reorder_snapshot()
        new_node = StructureNode("", values[0], values[1])
        parent.add_child(new_node)
        self._reorder_refresh(new_node)

    def _reorder_make_child(self):
        node = self._reorder_selected_node()
        if not node or not node.parent:
            return
        parent = self._reorder_parent_dialog(
            "Make Child",
            f"Select the new parent for '{node.description}'. Descendants and the selected item itself are hidden to prevent cycles.",
            excluded_node=node,
            allow_top=True,
        )
        if parent is None or parent is node.parent:
            return
        self._reorder_snapshot()
        node.parent.children.remove(node)
        parent.add_child(node)
        self._reorder_refresh(node)

    def _reorder_add_top(self):
        if not self.reorder_model:
            return
        values = self._reorder_prompt("Add Top Level")
        if not values:
            return
        self._reorder_snapshot()
        node = StructureNode("", values[0], values[1])
        self.reorder_model.root.add_child(node)
        self._reorder_refresh(node)

    def _reorder_add_child(self):
        node = self._reorder_selected_node()
        if not node:
            return
        values = self._reorder_prompt("Add Child")
        if not values:
            return
        self._reorder_snapshot()
        new_node = StructureNode("", values[0], values[1])
        node.add_child(new_node)
        self._reorder_refresh(new_node)

    def _reorder_add_sibling(self):
        node = self._reorder_selected_node()
        if not node or not node.parent:
            return
        values = self._reorder_prompt("Add Sibling")
        if not values:
            return
        self._reorder_snapshot()
        siblings = node.parent.children
        idx = siblings.index(node) + 1
        new_node = StructureNode("", values[0], values[1])
        new_node.parent = node.parent
        siblings.insert(idx, new_node)
        self._reorder_refresh(new_node)

    def _reorder_edit(self):
        node = self._reorder_selected_node()
        if not node:
            return
        values = self._reorder_prompt("Edit Item", node.description, node.part_number)
        if not values:
            return
        self._reorder_snapshot()
        node.description, node.part_number = values
        self._reorder_refresh(node)

    def _reorder_move(self, delta):
        node = self._reorder_selected_node()
        if not node or not node.parent:
            return
        siblings = node.parent.children
        idx = siblings.index(node)
        new_idx = idx + delta
        if new_idx < 0 or new_idx >= len(siblings):
            return
        self._reorder_snapshot()
        siblings[idx], siblings[new_idx] = siblings[new_idx], siblings[idx]
        self._reorder_refresh(node)

    def _reorder_make_child_of_previous(self):
        node = self._reorder_selected_node()
        if not node or not node.parent:
            return
        siblings = node.parent.children
        idx = siblings.index(node)
        if idx <= 0:
            return
        new_parent = siblings[idx - 1]
        if node.is_ancestor_of(new_parent):
            return
        self._reorder_snapshot()
        siblings.pop(idx)
        new_parent.add_child(node)
        self._reorder_refresh(node)

    def _reorder_promote(self):
        node = self._reorder_selected_node()
        if not node or not node.parent or not node.parent.parent:
            return
        old_parent = node.parent
        grandparent = old_parent.parent
        self._reorder_snapshot()
        old_parent.children.remove(node)
        insert_at = grandparent.children.index(old_parent) + 1
        node.parent = grandparent
        grandparent.children.insert(insert_at, node)
        self._reorder_refresh(node)

    def _reorder_remove(self):
        node = self._reorder_selected_node()
        if not node or not node.parent:
            return
        siblings = node.parent.children
        idx = siblings.index(node)
        self._reorder_snapshot()
        siblings.pop(idx)
        self._reorder_refresh()

    def _reorder_clear_editor(self):
        if not self.reorder_model:
            return
        self._reorder_snapshot()
        self.reorder_model = StructureModel()
        self.reorder_source_path = None
        self.reorder_loaded_files.clear()
        self.reorder_open_state.clear()
        self._reorder_refresh()
        self._reorder_set_drag_status("Editor cleared. Use Undo to restore the previous structure.")

    def _reorder_undo(self):
        if not self.reorder_undo_stack:
            return
        df = self.reorder_undo_stack.pop()
        self.reorder_model = StructureModel.from_dataframe(df)
        self._reorder_refresh()

    def _reorder_expand(self, expand):
        if not self._reorder_tree_available():
            return
        def toggle(item, force_open=False):
            self.reorder_tree.item(item, open=(True if force_open else expand))
            for child in self.reorder_tree.get_children(item):
                toggle(child)
        for item in self.reorder_tree.get_children():
            toggle(item, force_open=not expand)

    def _current_reorder_structure_path_for_download(self):
        if not self.reorder_model or not self.reorder_model.root.children:
            raise ValueError("No Structure Editor data is loaded.")
        tmp = tempfile.NamedTemporaryFile(prefix="structure_editor_", suffix=".xlsx", delete=False)
        tmp.close()
        self.reorder_model.to_dataframe().to_excel(tmp.name, index=False)
        return tmp.name

    def _prepare_reference_input(self, input_value):
        value = input_value.strip()
        if value == "__CURRENT_STRUCTURE_EDITOR__":
            return self._current_reorder_structure_path_for_download(), "current Structure Editor data"
        kind = classify_workbook(value)
        if kind == "structure":
            return value, "structure workbook"
        frame, _details = cad_export_to_structure_dataframe(value)
        tmp = tempfile.NamedTemporaryFile(prefix="cad_export_structure_", suffix=".xlsx", delete=False)
        tmp.close()
        frame.to_excel(tmp.name, index=False)
        return tmp.name, "CAD export"

    # ── Reference Download ────────────────────────────────────────────────────

    def _reference_page(self):
        main = self._scrollable_main()
        color = WORKFLOW_META["reference_download"]["color"]
        self._page_header(main, "Drawing Downloader",
                          "Download drawings from a structure workbook, CAD export, or current Structure Editor data.", color)

        card = self._card(main)
        self._section_label(card, "Inputs")

        structure_var = tk.StringVar()
        output_var = tk.StringVar()

        use_editor_var = tk.BooleanVar(value=False)
        editor_available = bool(self.reorder_model and self.reorder_model.root.children)

        row_frame = tk.Frame(card, bg=C["card"])
        row_frame.pack(fill="x", pady=(0, 14))

        input_row = tk.Frame(row_frame, bg=C["card"])

        def browse_reference_input():
            path = self._askopenfilename(filetypes=[("Supported", "*.xlsx *.xlsm *.xls *.csv"), ("All", "*.*")])
            if path:
                structure_var.set(path)
                if not output_var.get().strip():
                    output_var.set(default_output_path(path, "_drawings", "", self.naming_templates.get("drawing_download_folder")))

        def update_reference_source_state(*_):
            using_editor = use_editor_var.get()
            if using_editor:
                structure_var.set("__CURRENT_STRUCTURE_EDITOR__")
            elif structure_var.get() == "__CURRENT_STRUCTURE_EDITOR__":
                structure_var.set("")
            state = "disabled" if using_editor else "normal"
            source_entry.configure(
                state=state,
                disabledbackground=C["border"],
                disabledforeground=C["text_muted"],
            )
            browse.configure(state=state, cursor=("arrow" if using_editor else "hand2"))

        editor_check = tk.Checkbutton(
            row_frame, text="Use Structure Editor structure", variable=use_editor_var,
            command=update_reference_source_state,
            bg=C["card"], fg=C["text_dim"], selectcolor=C["surface"],
            activebackground=C["card"], activeforeground=C["text"],
            disabledforeground=C["text_muted"], font=(APP_FONT, 10),
            bd=0, highlightthickness=0, anchor="w",
            state=("normal" if editor_available else "disabled"),
        )
        editor_check.pack(fill="x", pady=(0, 8))
        tk.Label(row_frame, text="Structure workbook or CAD export", bg=C["card"], fg=C["text_dim"],
                 font=(APP_FONT, 9), anchor="w").pack(fill="x", pady=(0, 4))
        input_row.pack(fill="x")
        source_entry = tk.Entry(
            input_row, textvariable=structure_var, bg=C["bg"], fg=C["text"],
            insertbackground=C["text"], selectbackground=C["accent_muted"],
            selectforeground=C["text"], font=(MONO_FONT, 9), bd=0,
            highlightthickness=1, highlightbackground=C["border"],
            highlightcolor=C["accent"], relief="flat",
        )
        source_entry.pack(side="left", fill="x", expand=True, ipady=7, ipadx=8)
        browse = self._small_btn(input_row, "Browse…", browse_reference_input)
        browse.pack(side="left", padx=(8, 0))
        update_reference_source_state()

        self._divider(card)
        self._section_label(card, "Output")
        self._field(card, "Download folder", output_var,
                    lambda: output_var.set(self._askdirectory() or output_var.get()),
                    "Browse folder…")

        def run():
            if not structure_var.get() or not output_var.get():
                messagebox.showwarning("Missing fields", "Choose a structure source and output folder.", parent=self)
                return
            progress = ProgressDialog(self, "Downloading Drawings")
            result = None
            error = None
            try:
                structure_path, input_kind = self._prepare_reference_input(structure_var.get())
                result = download_references(structure_path, output_var.get(), progress_callback=progress.update)
                result["input_kind"] = input_kind
            except Exception as exc:
                error = str(exc)
            finally:
                progress.close()
            if error:
                self._show_themed_dialog("Download failed", error, tone="error")
                return
            self._show_themed_dialog(
                "Download complete",
                f"Input:       {result.get('input_kind', 'structure workbook')}\n"
                f"Downloaded:  {len(result['downloaded'])}\n"
                f"Skipped:     {len(result['skipped'])}\n"
                f"Not found:   {len(result['missing_parts'])}\n"
                f"Failed:      {len(result['failed'])}\n\n"
                f"Not found:\n{summarize_list(result['missing_parts'])}\n\n"
                f"Failed:\n{summarize_list(result['failed'])}",
                tone="info",
            )

        self._divider(card)
        run_btn = self._run_btn(card, "Download Drawings", run, color)
        run_btn.pack(anchor="w")
        self._bind_enabled_state(run_btn, [structure_var, output_var])

    # ── Settings ───────────────────────────────────────────────────────────────

    def _settings_page(self):
        main = self._scrollable_main()
        color = WORKFLOW_META["settings"]["color"]
        self._page_header(main, "Settings", "Application preferences and defaults", color)

        card = self._card(main)
        self._section_label(card, "Appearance")
        theme_var = tk.StringVar(value=self.theme_name)
        theme_row = tk.Frame(card, bg=C["card"])
        theme_row.pack(fill="x", pady=(0, 14))
        tk.Label(theme_row, text="Theme", bg=C["card"], fg=C["text_dim"],
                 font=(APP_FONT, 9), anchor="w").pack(fill="x", pady=(0, 8))
        for value, label in [("dark", "Dark mode"), ("light", "Light mode")]:
            rb = tk.Radiobutton(
                theme_row, text=label, variable=theme_var, value=value,
                command=lambda v=value: self._set_theme(v),
                bg=C["card"], fg=C["text_dim"], selectcolor=C["surface"],
                activebackground=C["card"], activeforeground=C["text"],
                font=(APP_FONT, 10), bd=0, highlightthickness=0,
            )
            rb.pack(side="left", padx=(0, 18))

        self._divider(card)
        self._section_label(card, "Default directory")
        directory_var = tk.StringVar(value=self.default_directory)
        self._field(
            card,
            "Folder used first when browsing for files",
            directory_var,
            lambda: directory_var.set(self._askdirectory() or directory_var.get()),
            "Browse folder…",
        )

        self._divider(card)
        self._section_label(card, "Default naming templates")
        tk.Label(
            card,
            text="Use {base} for the selected input filename, {ext} for the extension without a dot, and {extension} for the extension with a dot.",
            bg=C["card"], fg=C["text_muted"], font=(APP_FONT, 9), justify="left", anchor="w", wraplength=820,
        ).pack(fill="x", pady=(0, 10))

        template_vars = {}
        template_labels = [
            ("manual_packet", "Manual packet PDF"),
            ("automated_packet", "Automated packet PDF"),
            ("structure_export", "CAD-to-structure workbook"),
            ("structure_editor", "Structure Editor save workbook"),
            ("drawing_download_folder", "Drawing download folder"),
        ]
        for key, label in template_labels:
            template_vars[key] = tk.StringVar(value=self.naming_templates.get(key, DEFAULT_NAMING_TEMPLATES[key]))
            tk.Label(card, text=label, bg=C["card"], fg=C["text_dim"],
                     font=(APP_FONT, 9), anchor="w").pack(fill="x", pady=(0, 4))
            tk.Entry(
                card, textvariable=template_vars[key], bg=C["bg"], fg=C["text"],
                insertbackground=C["text"], selectbackground=C["accent_muted"],
                selectforeground=C["text"], font=(MONO_FONT, 9), bd=0,
                highlightthickness=1, highlightbackground=C["border"],
                highlightcolor=C["accent"], relief="flat",
            ).pack(fill="x", pady=(0, 12), ipady=7, ipadx=8)

        btn_row = tk.Frame(card, bg=C["card"])
        btn_row.pack(fill="x")

        def save_settings():
            path = directory_var.get().strip()
            if path and not os.path.isdir(path):
                messagebox.showerror("Invalid directory", "Choose an existing folder.", parent=self)
                return
            cleaned_templates = {}
            for key, var in template_vars.items():
                value = var.get().strip()
                if not value:
                    messagebox.showerror("Invalid template", "Naming templates cannot be blank.", parent=self)
                    return
                cleaned_templates[key] = value
            self.default_directory = path
            self.last_folder = path or self.last_folder
            self.naming_templates = cleaned_templates
            self._save_config()
            self._show_themed_dialog("Settings saved", "Your preferences and output naming templates have been saved.", tone="info")

        save_btn = self._run_btn(btn_row, "Save Settings", save_settings, color)
        save_btn.pack(anchor="w")

    # ── Shared file browse helper ─────────────────────────────────────────────

    def _browse_file(self, input_var, output_var, suffix, ext, filetypes, template_key=None):
        path = self._askopenfilename(filetypes=filetypes)
        if not path:
            return
        input_var.set(path)
        if not output_var.get().strip():
            output_var.set(default_output_path(path, suffix, ext, self.naming_templates.get(template_key)))

    def _browse_files(self, input_var, output_var, suffix, ext, filetypes, template_key=None):
        paths = self._askopenfilenames(filetypes=filetypes)
        if not paths:
            return
        input_var.set("; ".join(paths))
        if not output_var.get().strip():
            output_var.set(default_output_path(paths[0], suffix, ext, self.naming_templates.get(template_key)))


def main():
    app = DrawingCompilerStudio()
    app.mainloop()


if __name__ == "__main__":
    main()
