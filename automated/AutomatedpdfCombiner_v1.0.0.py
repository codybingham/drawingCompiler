import json
import logging
import os
import re
import sys
from io import BytesIO
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk

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
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Suppress only for this session; scoped to avoid global side-effects in larger processes
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

logging.basicConfig(level=logging.WARNING, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

APP_VERSION = "1.0.0"

LOOKUP_URL = "http://prints.spudnik.local/api/prints/format-paths"

EXCLUDED_ITEMS = {
    "HA0814",
    "HA0815",
    "HA0816",
    "HA1129",
    "HA0817",
    "984398",
}

# ---------------------------------------------------------------------------
# Config persistence (remembers last-used folder across runs)
# ---------------------------------------------------------------------------

def _config_path():
    """Return path to config file next to the executable or script."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "drawing_compiler_config.json")


def load_config():
    try:
        with open(_config_path(), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(config):
    try:
        with open(_config_path(), "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)
    except Exception as exc:
        logger.warning("Could not save config: %s", exc)


# ---------------------------------------------------------------------------
# Helper – centre any Toplevel on screen
# ---------------------------------------------------------------------------

def center_toplevel(window):
    window.update_idletasks()
    w = window.winfo_width()
    h = window.winfo_height()
    sw = window.winfo_screenwidth()
    sh = window.winfo_screenheight()
    x = int((sw - w) / 2)
    y = int((sh - h) / 2)
    window.geometry(f"{w}x{h}+{x}+{y}")


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

class Node:
    def __init__(self, description, part_number, indent, source_index):
        self.description = description
        self.part_number = part_number
        self.indent = indent
        self.source_index = source_index
        self.children = []
        self.parent = None

    def add_child(self, child):
        child.parent = self
        self.children.append(child)


class ExportSection:
    def __init__(self, file_path, root_node):
        self.file_path = file_path
        self.root_node = root_node

    @property
    def file_name(self):
        return os.path.basename(self.file_path)


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

def normalize_header(value):
    return re.sub(r"[^a-z0-9]", "", str(value).strip().lower())


def find_column(df, candidates):
    normalized_map = {normalize_header(col): col for col in df.columns}
    for candidate in candidates:
        key = normalize_header(candidate)
        if key in normalized_map:
            return normalized_map[key]
    raise KeyError(f"Could not find any of these columns: {candidates}")


def get_indent_level(object_value):
    text = "" if object_value is None else str(object_value)
    leading_spaces = len(text) - len(text.lstrip(" "))
    return leading_spaces // 4


def is_valid_item_number(item_number):
    text = "" if item_number is None else str(item_number).strip().upper()

    if not text.startswith(("13", "FB", "HA")):
        return False

    if text in EXCLUDED_ITEMS:
        return False

    return True


def is_skippable_nonpart_row(object_value, name_value):
    # Both columns are checked independently because either one may carry the
    # sentinel value depending on how the CAD export was structured.
    obj = "" if object_value is None else str(object_value).strip().upper()
    name = "" if name_value is None else str(name_value).strip().upper()

    if obj in {"SECTIONS", "CONSTRAINTS"}:
        return True

    if name in {"SECTIONS", "CONSTRAINTS"}:
        return True

    return False


def validate_output_filename(filename):
    invalid_chars_pattern = r'[<>:"/\\|?*]'
    reserved_names = {
        "CON", "PRN", "AUX", "NUL",
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
    }

    if not filename or not filename.strip():
        raise ValueError("Output file name cannot be blank.")

    filename = filename.strip()

    if re.search(invalid_chars_pattern, filename):
        raise ValueError('Output file name contains invalid characters: <>:"/\\|?*')

    if filename.endswith(" ") or filename.endswith("."):
        raise ValueError("Output file name cannot end with a space or period.")

    base_name = os.path.splitext(filename)[0].upper()
    if base_name in reserved_names:
        raise ValueError(f'"{base_name}" is a reserved Windows file name.')

    return filename


def get_schematic_label(schematic_file_path):
    return "HYDRAULIC SCHEMATIC"


def is_hydraulic_schematic_entry(description):
    return str(description).strip().upper().startswith("HYDRAULIC SCHEMATIC")


# ---------------------------------------------------------------------------
# BOM parsing
# ---------------------------------------------------------------------------

def build_nodes_from_rows(filtered_rows):
    root = Node(description="ROOT", part_number="", indent=-1, source_index=-1)
    stack = [root]

    for row in filtered_rows:
        node = Node(
            description=row["Description"],
            part_number=row["Part Number"],
            indent=row["indent"],
            source_index=row["source_index"],
        )

        while stack and stack[-1].indent >= node.indent:
            stack.pop()

        parent = stack[-1]
        parent.add_child(node)
        stack.append(node)

    return root


def collect_preserved_rows(df, object_col, name_col, item_col):
    rows = []

    for source_index, (_, row) in enumerate(df.iterrows()):
        item_number = "" if pd.isna(row[item_col]) else str(row[item_col]).strip()
        description = "" if pd.isna(row[name_col]) else str(row[name_col]).strip()
        object_value = "" if pd.isna(row[object_col]) else str(row[object_col])
        indent = get_indent_level(object_value)

        rows.append(
            {
                "source_index": source_index,
                "indent": indent,
                "Description": description,
                "Part Number": item_number,
                "Object": object_value,
                "keep": False,
                "direct_match": is_valid_item_number(item_number),
                "skippable": is_skippable_nonpart_row(object_value, description),
            }
        )

    keep_stack = []

    for row in rows:
        indent = row["indent"]

        while keep_stack and keep_stack[-1]["indent"] >= indent:
            keep_stack.pop()

        if row["direct_match"]:
            row["keep"] = True
            for ancestor in keep_stack:
                if not ancestor["skippable"] and ancestor["Description"]:
                    ancestor["keep"] = True

        keep_stack.append(row)

    filtered_rows = []
    for row in rows:
        if not row["keep"]:
            continue

        filtered_rows.append(
            {
                "indent": row["indent"],
                "Description": row["Description"],
                "Part Number": row["Part Number"],
                "source_index": row["source_index"],
            }
        )

    return filtered_rows


def load_export_section(file_path):
    df = pd.read_excel(file_path)

    object_col = find_column(df, ["Object"])
    name_col = find_column(df, ["Name"])
    item_col = find_column(df, ["Item Number", "ItemNumber", "Item No", "Item"])

    filtered_rows = collect_preserved_rows(df, object_col, name_col, item_col)

    if not filtered_rows:
        raise ValueError(
            f"No matching rows found in {os.path.basename(file_path)}. "
            "No item numbers started with 13, FB, or HA after exclusions."
        )

    structure_root = build_nodes_from_rows(filtered_rows)
    return ExportSection(file_path, structure_root)


def flatten_sections_to_rows(export_sections, schematic_file_path):
    rows = [
        {
            "Level": "1",
            "Description": get_schematic_label(schematic_file_path),
            "Part Number": "",
            "IsSection": False,
        }
    ]

    next_top_level = 2

    def walk(nodes, prefix):
        for idx, node in enumerate(nodes, start=1):
            current = prefix + [idx]
            rows.append(
                {
                    "Level": ".".join(str(x) for x in current),
                    "Description": node.description,
                    "Part Number": node.part_number,
                    "IsSection": False,
                }
            )
            walk(node.children, current)

    for section in export_sections:
        top_nodes = section.root_node.children

        for top_node in top_nodes:
            rows.append(
                {
                    "Level": str(next_top_level),
                    "Description": top_node.description,
                    "Part Number": top_node.part_number,
                    "IsSection": False,
                }
            )

            walk(top_node.children, [next_top_level])
            next_top_level += 1

    return rows


# ---------------------------------------------------------------------------
# Hierarchy / TOC helpers
# ---------------------------------------------------------------------------

def parse_level_code(value):
    text = str(value).strip()
    if not text:
        return tuple()

    parts = []
    for token in text.split("."):
        token = token.strip()
        if token == "":
            continue
        try:
            parts.append(int(token))
        except ValueError:
            parts.append(token)

    return tuple(parts)


def build_hierarchy(entries):
    seen_codes = {}
    processed = []

    for entry in entries:
        code = entry["code_tuple"]

        parent_index = None
        search = code[:-1]

        while search:
            if search in seen_codes:
                parent_index = seen_codes[search]
                break
            search = search[:-1]

        if parent_index is None:
            indent_level = 0
        else:
            indent_level = processed[parent_index]["indent_level"] + 1

        new_entry = dict(entry)
        new_entry["parent_index"] = parent_index
        new_entry["indent_level"] = indent_level

        processed.append(new_entry)
        seen_codes[code] = len(processed) - 1

    return processed


def _get_toc_fonts():
    """Defer font registration until actually needed."""
    candidates = [
        (r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\arialbd.ttf"),
        (r"C:\Windows\Fonts\Arial.ttf", r"C:\Windows\Fonts\Arialbd.ttf"),
        ("/usr/share/fonts/truetype/msttcorefonts/Arial.ttf", "/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf"),
        ("/usr/share/fonts/truetype/msttcorefonts/arial.ttf", "/usr/share/fonts/truetype/msttcorefonts/arialbd.ttf"),
        ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf"),
    ]

    regular_name = "Arial"
    bold_name = "Arial-Bold"

    for regular_path, bold_path in candidates:
        if os.path.exists(regular_path) and os.path.exists(bold_path):
            try:
                if regular_name not in pdfmetrics.getRegisteredFontNames():
                    pdfmetrics.registerFont(TTFont(regular_name, regular_path))
                if bold_name not in pdfmetrics.getRegisteredFontNames():
                    pdfmetrics.registerFont(TTFont(bold_name, bold_path))
                return regular_name, bold_name
            except Exception as exc:
                logger.warning("Could not register font %s: %s", regular_name, exc)

    return "Helvetica", "Helvetica-Bold"


def _build_index_entries(toc_entries):
    grouped = {}
    for entry in toc_entries:
        key = entry["desc"].strip().casefold()
        if key not in grouped:
            grouped[key] = {
                "desc": entry["desc"].strip(),
                "part": "",
                "item_numbers": [],
                "indent_level": 0,
                "toc_indices": [],
            }
        part = (entry.get("part") or "").strip()
        if part and not grouped[key]["part"]:
            grouped[key]["part"] = part
        item_number = (entry.get("item_number") or "").strip()
        if item_number and not is_hydraulic_schematic_entry(entry["desc"]) and item_number not in grouped[key]["item_numbers"]:
            grouped[key]["item_numbers"].append(item_number)
        grouped[key]["toc_indices"].append(entry["toc_index"])

    return sorted(grouped.values(), key=lambda entry: entry["desc"].casefold())


def _layout_directory_entries(entries, is_index=False, desc_font_name="Helvetica", desc_font_size=8):
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
            desc_lines = _wrap_text_to_width(
                entry["desc"],
                desc_font_name,
                desc_font_size,
                desc_max_width,
            )

            part_lines = []
            if is_index:
                part_text = str(entry.get("part") or "").strip()
                if part_text:
                    part_lines = _wrap_text_to_width(
                        part_text,
                        desc_font_name,
                        desc_font_size,
                        item_column_width,
                    )

            row_span = max(len(desc_lines), len(part_lines), 1)

            if (row_cursor % rows_per_column) + row_span > rows_per_column:
                row_cursor += rows_per_column - (row_cursor % rows_per_column)
                continue
            break

        placements.append(
            {
                "entry_index": idx,
                "page_index": page_index,
                "column_left": column_left,
                "desc_x": desc_x,
                "item_x": item_x,
                "item_left_x": item_left_x,
                "page_x": page_x,
                "page_left_x": page_left_x,
                "y": y,
                "title_y": title_y,
                "row_span": row_span,
                "desc_lines": desc_lines,
                "part_lines": part_lines,
            }
        )
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
    toc_font_regular, toc_font_bold = _get_toc_fonts()
    desc_font_size = 8

    packet = BytesIO()
    page_size, placements, total_pages = _layout_directory_entries(
        entries,
        is_index=is_index,
        desc_font_name=toc_font_regular,
        desc_font_size=desc_font_size,
    )
    c = canvas.Canvas(packet, pagesize=page_size)
    width, _ = page_size

    def draw_header(title_y):
        c.setFont(toc_font_bold, 16)
        c.drawString(40, title_y, title)

        c.setLineWidth(0.5)
        c.line(40, title_y - 6, width - 40, title_y - 6)
        c.setFont(toc_font_regular, 10)

    current_page = -1
    for placement in placements:
        entry = entries[placement["entry_index"]]

        if placement["page_index"] != current_page:
            if current_page != -1:
                c.showPage()
            current_page = placement["page_index"]
            draw_header(placement["title_y"])

        desc = entry["desc"]
        item_number = (entry.get("item_number") or "").strip()

        if is_hydraulic_schematic_entry(desc):
            display_item = ""
        elif is_index:
            display_item = (entry.get("part") or "").strip()
        else:
            display_item = item_number

        entry_index = placement["entry_index"]
        page_num = ""
        if "page_text" in entry and entry["page_text"]:
            page_num = entry["page_text"]
        elif page_offset_map is not None and page_offset_map[entry_index] is not None:
            page_num = str(page_offset_map[entry_index] + 1)

        c.setFont(toc_font_regular, desc_font_size)
        if is_index:
            desc_line_height = 9
            for line_index, desc_line in enumerate(placement.get("desc_lines", [desc])):
                c.drawString(placement["desc_x"], placement["y"] - (line_index * desc_line_height), desc_line)
        else:
            desc_line_height = 9
            for line_index, desc_line in enumerate(placement.get("desc_lines", [desc])):
                c.drawString(placement["desc_x"], placement["y"] - (line_index * desc_line_height), desc_line)

        if display_item:
            if is_index:
                item_line_height = 9
                for line_index, item_line in enumerate(placement.get("part_lines", [display_item])):
                    c.drawRightString(placement["item_x"], placement["y"] - (line_index * item_line_height), item_line)
            else:
                item_text = _trim_text_to_width(
                    display_item,
                    toc_font_regular,
                    8,
                    placement["item_x"] - placement["item_left_x"],
                )
                c.drawRightString(placement["item_x"], placement["y"], item_text)

        if page_num:
            page_text = _trim_text_to_width(
                page_num,
                toc_font_regular,
                8,
                placement["page_x"] - placement["page_left_x"],
            )
            c.drawRightString(placement["page_x"], placement["y"], page_text)

    if current_page == -1:
        draw_header(page_size[1] - (0.75 * 72))

    c.save()
    packet.seek(0)
    return packet, placements, total_pages


def _add_internal_link_annotation(writer, from_page_index, target_page_index, rect):
    target_ref = writer.pages[target_page_index].indirect_reference
    annotation = DictionaryObject()
    annotation.update(
        {
            NameObject("/Type"): NameObject("/Annot"),
            NameObject("/Subtype"): NameObject("/Link"),
            NameObject("/Rect"): ArrayObject([FloatObject(x) for x in rect]),
            NameObject("/Border"): ArrayObject([NumberObject(0), NumberObject(0), NumberObject(0)]),
            NameObject("/A"): DictionaryObject(
                {
                    NameObject("/S"): NameObject("/GoTo"),
                    NameObject("/D"): ArrayObject([target_ref, NameObject("/Fit")]),
                }
            ),
        }
    )
    writer.add_annotation(from_page_index, annotation)


def add_toc_hyperlinks(writer, toc_placements, effective_page_map, source_page_offset=0, line_height=12):
    for placement in toc_placements:
        entry_index = placement["entry_index"]
        target_page = effective_page_map[entry_index]
        if target_page is None:
            continue

        _add_internal_link_annotation(
            writer,
            from_page_index=placement["page_index"] + source_page_offset,
            target_page_index=target_page,
            rect=[
                placement["desc_x"] - 2,
                placement["y"] - 1,
                placement["page_x"] + 1,
                placement["y"] + line_height,
            ],
        )


def add_index_hyperlinks(writer, index_placements, index_target_map, source_page_offset=0, line_height=12):
    for placement in index_placements:
        entry_index = placement["entry_index"]
        target_page = index_target_map[entry_index]
        if target_page is None:
            continue

        _add_internal_link_annotation(
            writer,
            from_page_index=placement["page_index"] + source_page_offset,
            target_page_index=target_page,
            rect=[
                placement["desc_x"] - 2,
                placement["y"] - 1,
                placement["page_x"] + 1,
                placement["y"] + line_height,
            ],
        )


def add_page_number_overlay(page, page_num_text, total_pages_text):
    width = float(page.mediabox.width)
    height = float(page.mediabox.height)

    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(width, height))
    c.setFont("Helvetica", 8)
    c.drawCentredString(width / 2, 5, f"{page_num_text} / {total_pages_text}")
    c.save()

    packet.seek(0)
    overlay_pdf = PdfReader(packet)
    overlay_page = overlay_pdf.pages[0]
    page.merge_page(overlay_page)


# ---------------------------------------------------------------------------
# Download
# ---------------------------------------------------------------------------

def download_missing_drawings(items, output_folder, progress_callback=None):
    """
    Download PDFs for each item in *items* that is not already cached locally.

    progress_callback(current, total, label) is called after each item attempt
    so the UI can update a progress bar.
    """
    session = requests.Session()

    # --- skip items already on disk ---
    items_to_fetch = []
    skipped = []
    for item in items:
        out_path = os.path.join(output_folder, f"{item}.pdf")
        if os.path.exists(out_path):
            skipped.append(item)
        else:
            items_to_fetch.append(item)

    if progress_callback:
        progress_callback(0, len(items_to_fetch), "Looking up drawings on server...")

    if not items_to_fetch:
        return skipped, []

    payload = {
        "items": items_to_fetch,
        "location": "current",
    }

    headers = {
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "Origin": "http://prints.spudnik.local",
        "Referer": "http://prints.spudnik.local/",
        "User-Agent": "Mozilla/5.0",
    }

    r = session.post(LOOKUP_URL, json=payload, headers=headers, timeout=60, verify=False)
    r.raise_for_status()
    data = r.json()

    found = data.get("paths", [])
    not_found = list(data.get("notFound", []))
    downloaded = list(skipped)  # already-cached items count as downloaded

    total = len(found)
    for idx, entry in enumerate(found):
        item = str(entry["item"]).strip()
        pdf_url = entry["path"]
        out_path = os.path.join(output_folder, f"{item}.pdf")

        if progress_callback:
            progress_callback(idx, total, f"Downloading {item}  ({idx + 1} / {total})")

        try:
            pdf_r = session.get(pdf_url, timeout=60, verify=False)
            pdf_r.raise_for_status()

            with open(out_path, "wb") as f:
                f.write(pdf_r.content)

            downloaded.append(item)
        except Exception as exc:
            logger.warning("Failed to download %s from %s: %s", item, pdf_url, exc)
            not_found.append(item)

    if progress_callback:
        progress_callback(total, total, "Download complete.")

    return downloaded, sorted(set(not_found))


# ---------------------------------------------------------------------------
# Page-map (iterative, safe for deep hierarchies)
# ---------------------------------------------------------------------------

def build_effective_page_map(toc_entries, direct_page_map):
    children_map = {i: [] for i in range(len(toc_entries))}
    for i, entry in enumerate(toc_entries):
        parent_index = entry.get("parent_index")
        if parent_index is not None:
            children_map[parent_index].append(i)

    effective = list(direct_page_map)  # copy; None entries will be filled in

    # Post-order iterative traversal so each parent is resolved after its children
    # Identify roots (no parent)
    roots = [i for i, e in enumerate(toc_entries) if e.get("parent_index") is None]

    visit_order = []
    stack = list(roots)
    while stack:
        idx = stack.pop()
        visit_order.append(idx)
        stack.extend(children_map[idx])

    # Process in reverse (children before parents)
    for idx in reversed(visit_order):
        if effective[idx] is not None:
            continue
        for child_index in children_map[idx]:
            if effective[child_index] is not None:
                effective[idx] = effective[child_index]
                break

    return effective


def build_page_maps(toc_entries, drawings_folder, toc_pages):
    direct_page_map = [None] * len(toc_entries)
    current_page = toc_pages

    for i, entry in enumerate(toc_entries):
        filename = entry.get("filename")
        if not filename:
            continue

        fpath = os.path.join(drawings_folder, filename)
        if os.path.exists(fpath):
            reader = PdfReader(fpath)
            direct_page_map[i] = current_page
            current_page += len(reader.pages)

    return direct_page_map, build_effective_page_map(toc_entries, direct_page_map), current_page


# ---------------------------------------------------------------------------
# Reorder dialog
# ---------------------------------------------------------------------------

class MultiExportReorderDialog:
    def __init__(self, parent, export_sections, schematic_file_path):
        self.parent = parent
        self.export_sections = export_sections[:]
        self.schematic_file_path = schematic_file_path
        self.result = False

        # Undo stack: each entry is a callable that restores the previous state
        self._undo_stack = []

        self.window = tk.Toplevel(parent)
        self.window.title(f"Structure Reorder - v{APP_VERSION}")
        self.window.geometry("1100x700")
        self.window.minsize(850, 550)
        self.window.grab_set()

        self.tree = None
        self.node_lookup = {}
        self.section_lookup = {}
        self.item_lookup = {}

        self.build_ui()
        self.populate_tree()
        self.update_buttons()

        self.window.update_idletasks()
        center_toplevel(self.window)
        self.window.lift()
        self.window.focus_force()
        self.window.attributes("-topmost", True)
        self.window.after(200, lambda: self.window.attributes("-topmost", False))

        self.window.protocol("WM_DELETE_WINDOW", self.on_cancel)

        # Keyboard shortcut: Delete key triggers Remove Level
        self.window.bind("<Delete>", lambda event: self.on_remove_level())
        self.tree.bind("<Delete>", lambda event: self.on_remove_level())

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def build_ui(self):
        main = ttk.Frame(self.window, padding=12)
        main.pack(fill="both", expand=True)

        instructions = (
            "Reorder siblings with Move Up / Move Down.\n"
            "Children stay attached to their parent.\n"
            "Use Add CAD Export to append another CAD export as a new top-level section.\n"
            "Press Delete to remove the selected level.  Press Ctrl+Z to undo a removal."
        )
        ttk.Label(main, text=instructions, justify="left").pack(anchor="w", pady=(0, 10))

        tree_frame = ttk.Frame(main)
        tree_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(
            tree_frame,
            columns=("part", "source"),
            show="tree headings",
            selectmode="browse",
        )
        self.tree.heading("#0", text="Description")
        self.tree.heading("part", text="Part Number")
        self.tree.heading("source", text="Source")
        self.tree.column("#0", width=550, anchor="w")
        self.tree.column("part", width=160, anchor="w")
        self.tree.column("source", width=220, anchor="w")

        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

        self.tree.bind("<<TreeviewSelect>>", lambda event: self.update_buttons())

        button_row = ttk.Frame(main)
        button_row.pack(fill="x", pady=(10, 0))

        self.move_up_btn = ttk.Button(button_row, text="Move Up", command=self.on_move_up)
        self.move_down_btn = ttk.Button(button_row, text="Move Down", command=self.on_move_down)
        self.remove_btn = ttk.Button(button_row, text="Remove Level", command=self.on_remove_level)
        self.undo_btn = ttk.Button(button_row, text="Undo Remove", command=self.on_undo_remove)
        self.add_export_btn = ttk.Button(button_row, text="Add CAD Export", command=self.on_add_export)
        self.expand_btn = ttk.Button(button_row, text="Expand All", command=self.expand_all)
        self.collapse_btn = ttk.Button(button_row, text="Collapse All", command=self.collapse_all)
        self.continue_btn = ttk.Button(button_row, text="Continue", command=self.on_continue)
        self.cancel_btn = ttk.Button(button_row, text="Cancel", command=self.on_cancel)

        self.move_up_btn.pack(side="left")
        self.move_down_btn.pack(side="left", padx=(8, 0))
        self.remove_btn.pack(side="left", padx=(8, 0))
        self.undo_btn.pack(side="left", padx=(8, 0))
        self.add_export_btn.pack(side="left", padx=(20, 0))
        self.expand_btn.pack(side="left", padx=(20, 0))
        self.collapse_btn.pack(side="left", padx=(8, 0))
        self.cancel_btn.pack(side="right")
        self.continue_btn.pack(side="right", padx=(0, 8))

        # Ctrl+Z for undo
        self.window.bind("<Control-z>", lambda event: self.on_undo_remove())

    # ------------------------------------------------------------------
    # Tree population
    # ------------------------------------------------------------------

    def populate_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.node_lookup.clear()
        self.section_lookup.clear()
        self.item_lookup.clear()

        root_id = self.tree.insert(
            "",
            "end",
            text=get_schematic_label(self.schematic_file_path),
            values=("", "Selected schematic"),
            open=True,
        )
        self.item_lookup[root_id] = ("root", None)

        for section in self.export_sections:
            for top_node in section.root_node.children:
                section_id = self.tree.insert(
                    root_id,
                    "end",
                    text=top_node.description,
                    values=(top_node.part_number, os.path.basename(section.file_path)),
                    open=True,
                )
                self.section_lookup[section_id] = (section, top_node)
                self.item_lookup[section_id] = ("section_root", (section, top_node))

                self._add_children(section_id, top_node.children, section)

    def _add_children(self, parent_item_id, nodes, section):
        for node in nodes:
            item_id = self.tree.insert(
                parent_item_id,
                "end",
                text=node.description,
                values=(node.part_number, os.path.basename(section.file_path)),
                open=True,
            )
            self.node_lookup[item_id] = node
            self.item_lookup[item_id] = ("node", node)
            self._add_children(item_id, node.children, section)

    # ------------------------------------------------------------------
    # Selection helpers
    # ------------------------------------------------------------------

    def get_selected_item_id(self):
        selection = self.tree.selection()
        if not selection:
            return None
        return selection[0]

    def get_selected_payload(self):
        item_id = self.get_selected_item_id()
        if not item_id:
            return None
        return self.item_lookup.get(item_id)

    def update_buttons(self):
        payload = self.get_selected_payload()

        if payload is None:
            self.move_up_btn.config(state="disabled")
            self.move_down_btn.config(state="disabled")
            self.remove_btn.config(state="disabled")
            self.undo_btn.config(state="normal" if self._undo_stack else "disabled")
            return

        kind, obj = payload

        if kind == "root":
            self.move_up_btn.config(state="disabled")
            self.move_down_btn.config(state="disabled")
            self.remove_btn.config(state="disabled")
            self.undo_btn.config(state="normal" if self._undo_stack else "disabled")
            return

        if kind == "section_root":
            section, top_node = obj
            top_nodes = section.root_node.children
            idx = top_nodes.index(top_node)

            section_idx = self.export_sections.index(section)
            can_move_up = idx > 0 or section_idx > 0
            can_move_down = idx < len(top_nodes) - 1 or section_idx < len(self.export_sections) - 1

            self.move_up_btn.config(state="normal" if can_move_up else "disabled")
            self.move_down_btn.config(state="normal" if can_move_down else "disabled")
            self.remove_btn.config(state="normal")
            self.undo_btn.config(state="normal" if self._undo_stack else "disabled")
            return

        if kind == "node":
            siblings = obj.parent.children if obj.parent else []
            idx = siblings.index(obj) if obj in siblings else -1
            self.move_up_btn.config(state="normal" if idx > 0 else "disabled")
            self.move_down_btn.config(state="normal" if 0 <= idx < len(siblings) - 1 else "disabled")
            self.remove_btn.config(state="normal")
            self.undo_btn.config(state="normal" if self._undo_stack else "disabled")
            return

    # ------------------------------------------------------------------
    # Open/close state preservation across refresh
    # ------------------------------------------------------------------

    def capture_open_state(self):
        open_state = {}

        def recurse(item_id):
            payload = self.item_lookup.get(item_id)
            if payload is not None:
                kind, obj = payload
                key = (kind, id(obj) if obj is not None else 0)
                open_state[key] = bool(self.tree.item(item_id, "open"))
            for child in self.tree.get_children(item_id):
                recurse(child)

        for top in self.tree.get_children():
            recurse(top)

        return open_state

    def restore_open_state(self, open_state):
        def recurse(item_id):
            payload = self.item_lookup.get(item_id)
            if payload is not None:
                kind, obj = payload
                key = (kind, id(obj) if obj is not None else 0)
                if key in open_state:
                    self.tree.item(item_id, open=open_state[key])
            for child in self.tree.get_children(item_id):
                recurse(child)

        for top in self.tree.get_children():
            self.tree.item(top, open=True)
            recurse(top)

    def refresh(self, selected_target=None):
        open_state = self.capture_open_state()
        self.populate_tree()
        self.restore_open_state(open_state)

        if selected_target is not None:
            target_kind, target_obj = selected_target
            for item_id, payload in self.item_lookup.items():
                kind, obj = payload
                if kind == target_kind and obj is target_obj:
                    self.tree.selection_set(item_id)
                    self.tree.focus(item_id)
                    self.tree.see(item_id)
                    break

        self.update_buttons()

    # ------------------------------------------------------------------
    # Move operations
    # ------------------------------------------------------------------

    def move_top_node_up(self, section, top_node):
        top_nodes = section.root_node.children
        idx = top_nodes.index(top_node)
        if idx > 0:
            top_nodes[idx - 1], top_nodes[idx] = top_nodes[idx], top_nodes[idx - 1]
            return section, top_node

        section_idx = self.export_sections.index(section)
        if section_idx <= 0:
            return section, top_node

        prev_section = self.export_sections[section_idx - 1]
        top_nodes.pop(idx)
        prev_section.root_node.children.append(top_node)

        if len(section.root_node.children) == 0:
            self.export_sections.pop(section_idx)

        return prev_section, top_node

    def move_top_node_down(self, section, top_node):
        top_nodes = section.root_node.children
        idx = top_nodes.index(top_node)
        if idx < len(top_nodes) - 1:
            top_nodes[idx + 1], top_nodes[idx] = top_nodes[idx], top_nodes[idx + 1]
            return section, top_node

        section_idx = self.export_sections.index(section)
        if section_idx >= len(self.export_sections) - 1:
            return section, top_node

        next_section = self.export_sections[section_idx + 1]
        top_nodes.pop(idx)
        next_section.root_node.children.insert(0, top_node)

        if len(section.root_node.children) == 0:
            self.export_sections.pop(section_idx)

        return next_section, top_node

    def on_move_up(self):
        payload = self.get_selected_payload()
        if payload is None:
            return

        kind, obj = payload

        if kind == "section_root":
            section, top_node = obj
            new_section, same_node = self.move_top_node_up(section, top_node)
            self.refresh(("section_root", (new_section, same_node)))
            return

        if kind == "node":
            siblings = obj.parent.children if obj.parent else []
            idx = siblings.index(obj) if obj in siblings else -1
            if idx <= 0:
                return
            siblings[idx - 1], siblings[idx] = siblings[idx], siblings[idx - 1]
            self.refresh(("node", obj))

    def on_move_down(self):
        payload = self.get_selected_payload()
        if payload is None:
            return

        kind, obj = payload

        if kind == "section_root":
            section, top_node = obj
            new_section, same_node = self.move_top_node_down(section, top_node)
            self.refresh(("section_root", (new_section, same_node)))
            return

        if kind == "node":
            siblings = obj.parent.children if obj.parent else []
            idx = siblings.index(obj) if obj in siblings else -1
            if idx < 0 or idx >= len(siblings) - 1:
                return
            siblings[idx + 1], siblings[idx] = siblings[idx], siblings[idx + 1]
            self.refresh(("node", obj))

    # ------------------------------------------------------------------
    # Remove with undo support
    # ------------------------------------------------------------------

    def on_remove_level(self):
        payload = self.get_selected_payload()
        if payload is None:
            return

        kind, obj = payload

        if kind == "root":
            return

        if kind == "section_root":
            section, top_node = obj
            answer = messagebox.askyesno(
                "Remove Level",
                f"Remove '{top_node.description}' and all of its children?",
                parent=self.window,
            )
            if not answer:
                return

            # Capture enough info to undo this removal
            node_idx = section.root_node.children.index(top_node) if top_node in section.root_node.children else None
            section_idx = self.export_sections.index(section) if section in self.export_sections else None

            if top_node in section.root_node.children:
                section.root_node.children.remove(top_node)

            section_was_removed = False
            if len(section.root_node.children) == 0 and section in self.export_sections:
                self.export_sections.remove(section)
                section_was_removed = True

            # Build undo closure
            def undo_section_root(
                _section=section,
                _top_node=top_node,
                _node_idx=node_idx,
                _section_idx=section_idx,
                _section_was_removed=section_was_removed,
            ):
                if _section_was_removed and _section not in self.export_sections:
                    if _section_idx is not None:
                        self.export_sections.insert(_section_idx, _section)
                    else:
                        self.export_sections.append(_section)
                if _top_node not in _section.root_node.children:
                    if _node_idx is not None:
                        _section.root_node.children.insert(_node_idx, _top_node)
                    else:
                        _section.root_node.children.append(_top_node)
                self.refresh(("section_root", (_section, _top_node)))

            self._undo_stack.append(undo_section_root)
            self.refresh()
            return

        if kind == "node":
            node = obj
            answer = messagebox.askyesno(
                "Remove Level",
                f"Remove '{node.description}' and all of its children?",
                parent=self.window,
            )
            if not answer:
                return

            parent = node.parent
            node_idx = parent.children.index(node) if parent and node in parent.children else None

            if parent and node in parent.children:
                parent.children.remove(node)

            def undo_node(_node=node, _parent=parent, _node_idx=node_idx):
                if _parent and _node not in _parent.children:
                    if _node_idx is not None:
                        _parent.children.insert(_node_idx, _node)
                    else:
                        _parent.children.append(_node)
                self.refresh(("node", _node))

            self._undo_stack.append(undo_node)
            self.refresh()

    def on_undo_remove(self):
        if not self._undo_stack:
            return
        undo_fn = self._undo_stack.pop()
        undo_fn()
        self.update_buttons()

    # ------------------------------------------------------------------
    # Add export
    # ------------------------------------------------------------------

    def on_add_export(self):
        config = load_config()
        initial_dir = config.get("last_folder") or None

        file_path = filedialog.askopenfilename(
            title="Select Additional CAD Export Excel File",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
            initialdir=initial_dir,
            parent=self.window,
        )
        if not file_path:
            return

        config["last_folder"] = os.path.dirname(file_path)
        save_config(config)

        existing_paths = {os.path.normcase(os.path.abspath(s.file_path)) for s in self.export_sections}
        if os.path.normcase(os.path.abspath(file_path)) in existing_paths:
            messagebox.showwarning("Already Added", f"{os.path.basename(file_path)} is already in the list.", parent=self.window)
            return

        try:
            section = load_export_section(file_path)
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self.window)
            return

        self.export_sections.append(section)

        selected_target = None
        if section.root_node.children:
            selected_target = ("section_root", (section, section.root_node.children[0]))

        self.refresh(selected_target)

    # ------------------------------------------------------------------
    # Expand / collapse
    # ------------------------------------------------------------------

    def expand_all(self):
        def recurse(item_id):
            self.tree.item(item_id, open=True)
            for child in self.tree.get_children(item_id):
                recurse(child)

        for top in self.tree.get_children():
            recurse(top)

    def collapse_all(self):
        def recurse(item_id):
            self.tree.item(item_id, open=False)
            for child in self.tree.get_children(item_id):
                recurse(child)

        for top in self.tree.get_children():
            # Keep the very top node open so the tree is still navigable
            self.tree.item(top, open=True)
            for child in self.tree.get_children(top):
                recurse(child)

    # ------------------------------------------------------------------
    # Continue / cancel
    # ------------------------------------------------------------------

    def on_continue(self):
        if not self.export_sections or all(len(section.root_node.children) == 0 for section in self.export_sections):
            messagebox.showerror("Error", "At least one assembly must remain.", parent=self.window)
            return

        self.result = True
        self.window.destroy()

    def on_cancel(self):
        self.result = False
        self.window.destroy()


# ---------------------------------------------------------------------------
# Main workflow
# ---------------------------------------------------------------------------

def main():
    root = tk.Tk()
    root.withdraw()

    config = load_config()
    initial_dir = config.get("last_folder") or None

    first_cad_export = filedialog.askopenfilename(
        title=f"Automated Drawing Packet Builder v{APP_VERSION} - Select Initial CAD Export Excel File",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        initialdir=initial_dir,
    )
    if not first_cad_export:
        raise Exception("No CAD export Excel file selected.")

    config["last_folder"] = os.path.dirname(first_cad_export)
    save_config(config)

    schematic_file = filedialog.askopenfilename(
        title="Select Hydraulic Schematic PDF",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        initialdir=config.get("last_folder"),
    )
    if not schematic_file:
        raise Exception("No schematic PDF selected.")

    output_name = simpledialog.askstring(
        "Output File Name",
        "Enter compiled PDF file name:",
        initialvalue="CombinedDrawings.pdf",
        parent=root,
    )
    if output_name is None:
        raise Exception("Output file name entry cancelled.")

    output_name = output_name.strip()
    if not output_name:
        output_name = "CombinedDrawings.pdf"
    if not output_name.lower().endswith(".pdf"):
        output_name += ".pdf"

    output_name = validate_output_filename(output_name)

    base_folder = os.path.dirname(first_cad_export)
    output_pdf = os.path.join(base_folder, output_name)

    drawings_folder = os.path.join(base_folder, "_downloaded_drawings")
    os.makedirs(drawings_folder, exist_ok=True)

    # --- Loading spinner ---
    loading = tk.Toplevel(root)
    loading.title("Loading")
    loading.geometry("400x90")
    loading.resizable(False, False)

    loading_label = ttk.Label(
        loading,
        text="Reading initial CAD export...",
        padding=20,
        justify="center",
    )
    loading_label.pack(expand=True)

    center_toplevel(loading)
    loading.lift()
    loading.attributes("-topmost", True)
    loading.after(200, lambda: loading.attributes("-topmost", False))
    loading.update()

    first_section = load_export_section(first_cad_export)
    loading.destroy()

    dialog = MultiExportReorderDialog(root, [first_section], schematic_file)
    root.wait_window(dialog.window)

    if not dialog.result:
        raise Exception("Operation cancelled.")

    export_sections = dialog.export_sections

    all_structure_rows = flatten_sections_to_rows(export_sections, schematic_file)
    structure_df = pd.DataFrame(all_structure_rows)

    part_col = find_column(structure_df, ["Part Number", "PartNumber", "Part #", "Part"])
    desc_col = find_column(structure_df, ["Description", "Desc"])
    level_col = find_column(structure_df, ["Level", "Level Code", "Structure Level"])

    items_to_download = sorted(
        {
            str(x).strip()
            for x in structure_df[part_col]
            if pd.notna(x) and str(x).strip()
        }
    )

    # --- Download progress window ---
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()

    progress_win = tk.Toplevel(root)
    progress_win.title("Downloading")
    progress_win.geometry("460x130")
    progress_win.resizable(False, False)

    progress_frame = ttk.Frame(progress_win, padding=16)
    progress_frame.pack(fill="both", expand=True)

    progress_label = ttk.Label(
        progress_frame,
        text="Looking up and downloading drawings...",
        justify="center",
        anchor="center",
    )
    progress_label.pack(fill="x")

    progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate", length=400)
    progress_bar.pack(pady=(10, 0))

    center_toplevel(progress_win)
    progress_win.lift()
    progress_win.attributes("-topmost", True)
    progress_win.after(200, lambda: progress_win.attributes("-topmost", False))
    progress_win.update()

    def on_download_progress(current, total, label):
        progress_label.config(text=label)
        if total > 0:
            progress_bar["maximum"] = total
            progress_bar["value"] = current
        progress_win.update()

    downloaded, not_found = download_missing_drawings(
        items_to_download, drawings_folder, progress_callback=on_download_progress
    )

    schematic_copy_path = os.path.join(drawings_folder, "__HYDRAULIC_SCHEMATIC__.pdf")
    with open(schematic_file, "rb") as src, open(schematic_copy_path, "wb") as dst:
        dst.write(src.read())

    progress_win.destroy()

    raw_entries = []

    for _, row in structure_df.iterrows():
        part = "" if pd.isna(row[part_col]) else str(row[part_col]).strip()
        desc = "" if pd.isna(row[desc_col]) else str(row[desc_col]).strip()
        code_text = "" if pd.isna(row[level_col]) else str(row[level_col]).strip()
        code_tuple = parse_level_code(code_text)

        if code_text == "1":
            filename = "__HYDRAULIC_SCHEMATIC__.pdf"
        elif part:
            filename = f"{part}.pdf"
        else:
            filename = None

        raw_entries.append(
            {
                "code_text": code_text,
                "code_tuple": code_tuple,
                "item_number": "" if code_text == "1" else part,
                "desc": desc,
                "part": part,
                "filename": filename,
            }
        )

    toc_entries = build_hierarchy(raw_entries)
    for i, entry in enumerate(toc_entries):
        entry["toc_index"] = i
    toc_display_entries = toc_entries + [
        {
            "desc": "Index",
            "item_number": "",
            "part": "",
            "indent_level": 0,
            "parent_index": None,
            "filename": None,
        }
    ]

    toc_packet, _, toc_pages = create_directory_pdf_bytes(toc_display_entries, "Table of Contents", None)
    index_entries = _build_index_entries(toc_entries)
    index_packet, _, index_pages = create_directory_pdf_bytes(index_entries, "Index", None, is_index=True)
    _, effective_page_map, index_start_page = build_page_maps(toc_entries, drawings_folder, toc_pages)
    toc_page_map = effective_page_map + [index_start_page]

    toc_packet, toc_placements, _ = create_directory_pdf_bytes(
        toc_display_entries,
        "Table of Contents",
        toc_page_map,
    )
    for entry in index_entries:
        pages = []
        for toc_index in entry["toc_indices"]:
            page = effective_page_map[toc_index]
            if page is not None and page not in pages:
                pages.append(page)
        entry["page_text"] = ", ".join(str(page + 1) for page in pages)
    index_packet, index_placements, _ = create_directory_pdf_bytes(index_entries, "Index", None, is_index=True)
    toc_reader = PdfReader(toc_packet)
    index_reader = PdfReader(index_packet)

    actual_toc_pages = len(toc_reader.pages)
    actual_index_pages = len(index_reader.pages)
    if actual_toc_pages != toc_pages or actual_index_pages != index_pages:
        toc_pages = actual_toc_pages
        index_pages = actual_index_pages
        _, effective_page_map, index_start_page = build_page_maps(toc_entries, drawings_folder, toc_pages)
        toc_page_map = effective_page_map + [index_start_page]

        toc_packet, toc_placements, _ = create_directory_pdf_bytes(
            toc_display_entries,
            "Table of Contents",
            toc_page_map,
        )
        for entry in index_entries:
            pages = []
            for toc_index in entry["toc_indices"]:
                page = effective_page_map[toc_index]
                if page is not None and page not in pages:
                    pages.append(page)
            entry["page_text"] = ", ".join(str(page + 1) for page in pages)
        index_packet, index_placements, _ = create_directory_pdf_bytes(index_entries, "Index", None, is_index=True)
        toc_reader = PdfReader(toc_packet)
        index_reader = PdfReader(index_packet)

    writer = PdfWriter()

    for page in toc_reader.pages:
        writer.add_page(page)

    for entry in toc_entries:
        filename = entry.get("filename")
        if not filename:
            continue

        fpath = os.path.join(drawings_folder, filename)
        if not os.path.exists(fpath):
            continue

        reader = PdfReader(fpath)
        for page in reader.pages:
            writer.add_page(page)

    for page in index_reader.pages:
        writer.add_page(page)

    bookmark_refs = {}

    for i, entry in enumerate(toc_entries):
        target_page = effective_page_map[i]
        if target_page is None:
            continue

        parent = None
        parent_index = entry["parent_index"]

        if parent_index is not None and parent_index in bookmark_refs:
            parent = bookmark_refs[parent_index]

        bookmark = writer.add_outline_item(
            entry["desc"],
            target_page,
            parent=parent,
        )
        bookmark_refs[i] = bookmark

    add_toc_hyperlinks(writer, toc_placements, toc_page_map, source_page_offset=0)

    total_pages = len(writer.pages)
    for i, page in enumerate(writer.pages):
        add_page_number_overlay(page, i + 1, total_pages)

    with open(output_pdf, "wb") as f:
        writer.write(f)

    # --- Missing drawings report ---
    missing_files = []
    for entry in toc_entries:
        filename = entry.get("filename")
        if filename and filename != "__HYDRAULIC_SCHEMATIC__.pdf":
            fpath = os.path.join(drawings_folder, filename)
            if not os.path.exists(fpath):
                missing_files.append(os.path.splitext(filename)[0])

    warning_items = sorted(set(not_found + missing_files))

    # Write missing list to a text file next to the output PDF
    if warning_items:
        missing_txt_path = os.path.splitext(output_pdf)[0] + "_missing_drawings.txt"
        try:
            with open(missing_txt_path, "w", encoding="utf-8") as f:
                f.write("Missing / not-found drawings\n")
                f.write("=" * 40 + "\n")
                for item in warning_items:
                    f.write(f"{item}\n")
        except Exception as exc:
            logger.warning("Could not write missing drawings file: %s", exc)
            missing_txt_path = None
    else:
        missing_txt_path = None

    message = (
        f"Automated Drawing Packet Builder v{APP_VERSION}\n\n"
        f"CAD exports used: {len(export_sections)}\n"
        f"Schematic used: {os.path.basename(schematic_file)}\n"
        f"Compiled PDF: {output_pdf}\n\n"
        f"Downloaded drawings: {len(downloaded)}"
    )

    if warning_items:
        message += "\n\nWarning - could not find/download these drawings:\n" + "\n".join(
            f"  - {item}" for item in warning_items
        )
        if missing_txt_path:
            message += f"\n\nFull list saved to:\n{missing_txt_path}"

    print(message)
    if warning_items:
        messagebox.showwarning("Done with Warnings", message, parent=root)
    else:
        messagebox.showinfo("Done", message, parent=root)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Error", str(e), parent=root)
        except Exception:
            pass
        raise
