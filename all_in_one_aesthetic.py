import os
import re
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
from reportlab.pdfgen import canvas

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

LOOKUP_URL = "http://prints.spudnik.local/api/prints/format-paths"
EXCLUDED_ITEMS = {"HA0814", "HA0815", "HA0816", "HA1129", "HA0817", "984398"}

# ─── Palette ────────────────────────────────────────────────────────────────
C = {
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

WORKFLOW_META = {
    "manual_packet":    {"color": C["accent"],  "muted": C["accent_muted"]},
    "automated_packet": {"color": C["green"],   "muted": C["green_muted"]},
    "cad_to_structure": {"color": C["amber"],   "muted": C["amber_muted"]},
    "reorder_structure":{"color": C["violet"],  "muted": C["violet_muted"]},
    "reference_download":{"color": C["rose"],   "muted": C["rose_muted"]},
}


@dataclass(frozen=True)
class Workflow:
    key: str
    title: str
    subtitle: str


WORKFLOWS = [
    Workflow("dashboard",          "Dashboard",             "Unified control center"),
    Workflow("manual_packet",      "Manual Packet",         "Build from local PDFs"),
    Workflow("automated_packet",   "Automated Packet",      "Download + build in one flow"),
    Workflow("cad_to_structure",   "CAD to Structure",      "Convert CAD exports"),
    Workflow("reorder_structure",  "Structure Reorder",     "Reorder and renumber levels"),
    Workflow("reference_download", "Drawing Downloader",    "Download drawing references"),
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
    page_size, placements, _ = _layout_directory_entries(entries, is_index=is_index, desc_font_name="Helvetica", desc_font_size=8)
    c = canvas.Canvas(packet, pagesize=page_size)
    width, _ = page_size

    def draw_header(title_y):
        c.setFont("Helvetica-Bold", 16)
        c.drawString(40, title_y, title)
        c.setLineWidth(0.5)
        c.line(40, title_y - 6, width - 40, title_y - 6)
        c.setFont("Helvetica", 10)

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
        c.setFont("Helvetica", 8)
        desc_line_height = 9
        for line_index, desc_line in enumerate(placement.get("desc_lines", [desc])):
            c.drawString(placement["desc_x"], placement["y"] - (line_index * desc_line_height), desc_line)
        if display_item:
            if is_index:
                item_line_height = 9
                for line_index, item_line in enumerate(placement.get("part_lines", [display_item])):
                    c.drawRightString(placement["item_x"], placement["y"] - (line_index * item_line_height), item_line)
            else:
                item_text = _trim_text_to_width(display_item, "Helvetica", 8, placement["item_x"] - placement["item_left_x"])
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
    c.setFont("Helvetica", 8)
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


def convert_cad_to_structure(input_path: str, output_path: str) -> dict:
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
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    out_df.to_excel(output_path, index=False)
    return {
        "input_path": input_path, "output_path": output_path,
        "source_rows": len(df), "rows_written": len(out_df),
        "mapping": {"object_col": object_col, "name_col": name_col, "item_number_col": item_col},
    }


def read_structure_references(structure_path: str) -> tuple[list[str], list[str]]:
    df = pd.read_excel(structure_path)
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
    failed = []
    total_downloads = len(found_paths) + len(direct_urls)
    completed_downloads = 0
    if progress_callback:
        progress_callback(completed_downloads, total_downloads, "Preparing downloads...")
    for part, url in found_paths.items():
        target = os.path.join(output_folder, f"{part}.pdf")
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
        try:
            download_url(session, url, target)
            downloaded.append(url)
        except Exception:
            failed.append(url)
        completed_downloads += 1
        if progress_callback:
            progress_callback(completed_downloads, total_downloads, f"Downloaded URL: {filename}")
    return {"downloaded": downloaded, "missing_parts": missing_parts, "failed": failed, "output_folder": output_folder}


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


def build_automated_packet(cad_export_path, schematic_pdf, temp_download_folder, output_pdf, progress_callback=None):
    structure_path = default_output_path(cad_export_path, "_structure", ".xlsx")
    convert_cad_to_structure(cad_export_path, structure_path)

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
        "included_parts": packet_result["included_parts"],
        "missing_parts": packet_result["missing_parts"],
        "failed_downloads": download_result["failed"],
        "not_found": download_result["missing_parts"],
    }


def load_structure_for_reorder(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)
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


def default_output_path(input_path, suffix, extension):
    if not input_path:
        return ""
    base = os.path.splitext(os.path.basename(input_path))[0]
    return os.path.join(os.path.dirname(input_path), f"{base}{suffix}{extension}")


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
            font=("Consolas", 9, "bold"), anchor="w",
        ).pack(fill="x", pady=(0, 12))

        self.status_var = tk.StringVar(value="Starting…")
        tk.Label(
            frame, textvariable=self.status_var,
            bg=C["card"], fg=C["text_dim"],
            font=("Segoe UI", 10), width=52, anchor="w", wraplength=380,
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
            font=("Consolas", 9),
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
        self.reorder_undo_stack: list[tuple] = []

        self._setup_styles()
        self._build_shell()
        self.show_workflow("dashboard", add_history=False)

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
            font=("Segoe UI", 10),
        )
        style.configure(
            "Reorder.Treeview.Heading",
            background=C["surface"],
            foreground=C["text_muted"],
            borderwidth=0,
            relief="flat",
            font=("Consolas", 9, "bold"),
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
            font=("Consolas", 10, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(0, 8))

        tk.Label(
            frame,
            text=message,
            justify="left",
            anchor="w",
            bg=C["card"],
            fg=C["text_dim"],
            font=("Segoe UI", 10),
            wraplength=560,
        ).pack(fill="x", pady=(0, 14))

        btn = tk.Button(
            frame,
            text=button_text,
            bg=muted,
            fg=accent,
            activebackground=accent,
            activeforeground="#FFFFFF",
            font=("Segoe UI", 10, "bold"),
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
                 font=("Consolas", 11, "bold")).pack(anchor="w")
        tk.Label(logo_text, text="Studio · v2", bg=C["surface"], fg=C["text_muted"],
                 font=("Consolas", 8)).pack(anchor="w")

        # Divider
        tk.Frame(f, bg=C["border"], height=1).pack(fill="x", padx=16, pady=(4, 12))

        # ── Nav label ──
        tk.Label(f, text="WORKFLOWS", bg=C["surface"], fg=C["text_muted"],
                 font=("Consolas", 8, "bold"), anchor="w",
                 padx=18).pack(fill="x", pady=(0, 6))

        # ── Nav buttons ──
        self.nav_buttons: dict[str, tk.Button] = {}
        icons = {
            "dashboard":           "⊞",
            "manual_packet":       "▤",
            "automated_packet":    "⚙",
            "cad_to_structure":    "⇄",
            "reorder_structure":   "≡",
            "reference_download":  "↓",
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
            font=("Segoe UI", 10), bd=0, cursor="hand2",
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
            font=("Segoe UI", 10),
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
                    font=("Segoe UI", 10, "bold"),
                )
            else:
                btn.configure(
                    bg=C["surface"],
                    fg=C["text_dim"],
                    font=("Segoe UI", 10),
                )

    # ── Navigation ────────────────────────────────────────────────────────────

    def _clear_main(self):
        for child in self.main_frame.winfo_children():
            child.destroy()
        self.reorder_tree = None

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
            "manual_packet":       self._manual_packet_page,
            "automated_packet":    self._automated_packet_page,
            "cad_to_structure":    self._cad_page,
            "reorder_structure":   self._reorder_page,
            "reference_download":  self._reference_page,
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
            font=("Segoe UI", 18, "bold"), anchor="w",
        ).pack(fill="x")
        tk.Label(
            text_block, text=subtitle,
            bg=C["bg"], fg=C["text_muted"],
            font=("Segoe UI", 10), anchor="w",
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
            font=("Consolas", 8, "bold"),
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
            font=("Segoe UI", 9), anchor="w",
        ).pack(side="left")
        if optional:
            tk.Label(
                lbl_row, text=" optional",
                bg=C["card"], fg=C["text_muted"],
                font=("Consolas", 8), anchor="w",
            ).pack(side="left")

        input_row = tk.Frame(row_frame, bg=C["card"])
        input_row.pack(fill="x")

        entry = tk.Entry(
            input_row, textvariable=var,
            bg=C["bg"], fg=C["text"],
            insertbackground=C["text"],
            selectbackground=C["accent_muted"],
            selectforeground=C["text"],
            font=("Consolas", 9),
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
            font=("Segoe UI", 9), bd=0, relief="flat", cursor="hand2",
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
            font=("Segoe UI", 10, "bold"), bd=0, relief="flat",
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
            "manual_packet":       "Build a merged PDF packet from local drawings and a structure workbook.",
            "automated_packet":    "Ingest a CAD export, download drawings, and build a full packet in one flow.",
            "cad_to_structure":    "Convert CAD export spreadsheets into the structure workbook format.",
            "reorder_structure":   "Full editor: add, reorder, edit, and remove structure rows, then save.",
            "reference_download":  "Download drawings by part number and URL from a structure workbook.",
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
                font=("Segoe UI", 12, "bold"), anchor="w",
            ).pack(side="left")

            # Description
            tk.Label(
                card_inner,
                text=descriptions.get(wf.key, wf.subtitle),
                bg=C["card"], fg=C["text_muted"],
                font=("Segoe UI", 9),
                anchor="w", wraplength=400, justify="left",
            ).pack(fill="x", pady=(0, 14))

            # Open button
            open_btn = tk.Button(
                card_inner, text="Open →",
                bg=muted, fg=color,
                activebackground=color, activeforeground="#fff",
                font=("Segoe UI", 9, "bold"), bd=0, relief="flat",
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
                                               [("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")]))
        self._field(card, "Drawings folder", drawings_var,
                    lambda: drawings_var.set(filedialog.askdirectory() or drawings_var.get()),
                    "Browse folder…")
        self._field(card, "Hydraulic schematic PDF", schematic_var,
                    lambda: schematic_var.set(
                        filedialog.askopenfilename(filetypes=[("PDF", "*.pdf"), ("All", "*.*")]) or schematic_var.get()),
                    optional=True)

        self._divider(card)
        self._section_label(card, "Output")
        self._field(card, "Output PDF path", output_var,
                    lambda: output_var.set(
                        filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")]) or output_var.get()),
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
        self._run_btn(card, "Build Manual Packet", run, color).pack(anchor="w")

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

        self._field(card, "CAD export (.xlsx / .csv)", cad_var,
                    lambda: self._browse_file(cad_var, output_var, "_automated_packet", ".pdf",
                                               [("Supported", "*.xlsx *.xlsm *.xls *.csv"), ("All", "*.*")]))
        self._field(card, "Hydraulic schematic PDF", schematic_var,
                    lambda: schematic_var.set(
                        filedialog.askopenfilename(filetypes=[("PDF", "*.pdf"), ("All", "*.*")]) or schematic_var.get()))
        self._field(card, "Download folder", download_var,
                    lambda: download_var.set(filedialog.askdirectory() or download_var.get()),
                    "Browse folder…")

        self._divider(card)
        self._section_label(card, "Output")
        self._field(card, "Output PDF path", output_var,
                    lambda: output_var.set(
                        filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")]) or output_var.get()),
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
                f"Structure file:     {result['structure_path']}\n"
                f"Parts included:     {result['included_parts']}\n"
                f"Missing in packet:  {len(result['missing_parts'])}\n"
                f"Download failures:  {len(result['failed_downloads'])}\n"
                f"Not found in lookup:{len(result['not_found'])}\n\n"
                f"Missing in packet:\n{summarize_list(result['missing_parts'])}\n\n"
                f"Lookup not found:\n{summarize_list(result['not_found'])}",
                tone="info",
            )

        self._divider(card)
        self._run_btn(card, "Run Automated Build", run, color).pack(anchor="w")

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

        self._field(card, "CAD export file (.xlsx / .xlsm / .xls / .csv)", input_var,
                    lambda: self._browse_file(input_var, output_var, "_structure", ".xlsx",
                                               [("Supported", "*.xlsx *.xlsm *.xls *.csv"), ("All", "*.*")]))

        self._divider(card)
        self._section_label(card, "Output")
        self._field(card, "Output structure workbook (.xlsx)", output_var,
                    lambda: output_var.set(
                        filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")]) or output_var.get()),
                    "Save as…")

        def run():
            if not input_var.get() or not output_var.get():
                messagebox.showwarning("Missing fields", "Choose input and output files.", parent=self)
                return
            try:
                result = convert_cad_to_structure(input_var.get(), output_var.get())
                m = result["mapping"]
                messagebox.showinfo("Conversion complete",
                    f"Input:   {result['input_path']}\n"
                    f"Output:  {result['output_path']}\n\n"
                    f"Rows read:    {result['source_rows']}\n"
                    f"Rows written: {result['rows_written']}\n\n"
                    f"Detected columns:\n"
                    f"  Object:      {m['object_col']}\n"
                    f"  Name:        {m['name_col']}\n"
                    f"  Item Number: {m['item_number_col']}", parent=self)
            except Exception as exc:
                messagebox.showerror("Conversion failed", str(exc), parent=self)

        self._divider(card)
        self._run_btn(card, "Generate Structure", run, color).pack(anchor="w")

    # ── Structure Reorder ─────────────────────────────────────────────────────

    def _reorder_page(self):
        color = WORKFLOW_META["reorder_structure"]["color"]

        # Header (non-scrollable area at top)
        header_frame = tk.Frame(self.main_frame, bg=C["bg"])
        header_frame.pack(fill="x")
        self._page_header(header_frame, "Structure Reorder",
                          "Add, reorder, edit, and remove rows — then save a renumbered structure workbook.", color)

        # Toolbar
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
                font=("Segoe UI", 9), bd=0, relief="flat", cursor="hand2",
                padx=10, pady=6,
            )
            btn.bind("<Enter>", lambda e: btn.configure(bg=hover, fg="#fff" if (accent or danger) else C["text"]))
            btn.bind("<Leave>", lambda e: btn.configure(bg=bg, fg=fg))
            btn.pack(side="left", padx=(0, 6))
            return btn

        tool_btn("Open structure…", self._reorder_open, accent=True)
        tool_btn("Save as…", self._reorder_save)

        tk.Frame(toolbar, bg=C["border"], width=1, height=24).pack(side="left", padx=10)

        tool_btn("Add top level", lambda: self._reorder_add_top())
        tool_btn("Add child", lambda: self._reorder_add_child())
        tool_btn("Add sibling", lambda: self._reorder_add_sibling())
        tool_btn("Edit", lambda: self._reorder_edit())

        tk.Frame(toolbar, bg=C["border"], width=1, height=24).pack(side="left", padx=10)

        tool_btn("↑ Up", lambda: self._reorder_move(-1))
        tool_btn("↓ Down", lambda: self._reorder_move(1))

        tk.Frame(toolbar, bg=C["border"], width=1, height=24).pack(side="left", padx=10)

        tool_btn("Remove", lambda: self._reorder_remove(), danger=True)
        tool_btn("Undo remove", lambda: self._reorder_undo())

        tk.Frame(toolbar, bg=C["border"], width=1, height=24).pack(side="left", padx=10)
        tool_btn("Expand all", lambda: self._reorder_expand(True))
        tool_btn("Collapse all", lambda: self._reorder_expand(False))

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

    def _reorder_open(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")])
        if not path:
            return
        try:
            df = load_structure_for_reorder(path)
            self.reorder_model = StructureModel.from_dataframe(df)
            self.reorder_source_path = path
            self.reorder_undo_stack.clear()
            if not self._reorder_tree_available():
                self.show_workflow("reorder_structure", add_history=False)
            self._reorder_refresh()
            self._reorder_apply_default_tree_layout()
        except Exception as exc:
            messagebox.showerror("Open failed", str(exc), parent=self)

    def _reorder_save(self):
        if not self.reorder_model:
            messagebox.showwarning("No data", "Open a structure file first.", parent=self)
            return
        initial = "reordered_structure.xlsx"
        if self.reorder_source_path:
            stem = os.path.splitext(os.path.basename(self.reorder_source_path))[0]
            initial = f"{stem}_reordered.xlsx"
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=initial, filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            self.reorder_model.to_dataframe().to_excel(path, index=False)
            messagebox.showinfo("Saved", f"Structure saved to:\n{path}", parent=self)
        except Exception as exc:
            messagebox.showerror("Save failed", str(exc), parent=self)

    def _reorder_refresh(self, select_node=None):
        if not self._reorder_tree_available():
            return
        for item in self.reorder_tree.get_children():
            self.reorder_tree.delete(item)
        self.reorder_item_lookup.clear()
        if not self.reorder_model:
            return

        def add_nodes(parent_id, nodes):
            for node in nodes:
                item_id = self.reorder_tree.insert(
                    parent_id, "end",
                    text=f"  {node.description}",
                    values=(node.part_number,), open=True,
                )
                self.reorder_item_lookup[item_id] = node
                add_nodes(item_id, node.children)

        add_nodes("", self.reorder_model.root.children)

        if select_node:
            for item_id, node in self.reorder_item_lookup.items():
                if node is select_node:
                    self.reorder_tree.selection_set(item_id)
                    self.reorder_tree.focus(item_id)
                    self.reorder_tree.see(item_id)
                    break

    def _reorder_apply_default_tree_layout(self):
        if not self._reorder_tree_available():
            return
        for top_item in self.reorder_tree.get_children():
            self.reorder_tree.item(top_item, open=True)
            for child in self.reorder_tree.get_children(top_item):
                self.reorder_tree.item(child, open=False)

    def _reorder_selected_node(self):
        if not self._reorder_tree_available():
            return None
        sel = self.reorder_tree.selection()
        return self.reorder_item_lookup.get(sel[0]) if sel else None

    def _reorder_tree_available(self):
        return bool(self.reorder_tree and self.reorder_tree.winfo_exists())

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
                 font=("Consolas", 9, "bold")).pack(anchor="w", pady=(0, 14))

        tk.Label(frame, text="Description", bg=C["card"], fg=C["text_dim"],
                 font=("Segoe UI", 9)).pack(anchor="w")
        desc_var = tk.StringVar(value=desc)
        tk.Entry(frame, textvariable=desc_var, bg=C["bg"], fg=C["text"],
                 insertbackground=C["text"], font=("Consolas", 10),
                 bd=0, highlightthickness=1, highlightbackground=C["border"],
                 highlightcolor=C["accent"], relief="flat", width=52
                 ).pack(fill="x", pady=(4, 12), ipady=6, ipadx=8)

        tk.Label(frame, text="Part Number", bg=C["card"], fg=C["text_dim"],
                 font=("Segoe UI", 9)).pack(anchor="w")
        part_var = tk.StringVar(value=part)
        tk.Entry(frame, textvariable=part_var, bg=C["bg"], fg=C["text"],
                 insertbackground=C["text"], font=("Consolas", 10),
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
                           font=("Segoe UI", 9), bd=0, relief="flat", cursor="hand2",
                           padx=12, pady=6)
        cancel.pack(side="right")

        ok_color = WORKFLOW_META["reorder_structure"]["color"]
        ok = self._run_btn(btns, "OK", on_ok, ok_color)
        ok.pack(side="right", padx=(0, 8))

        w, h = 440, 240
        dialog.geometry(f"{w}x{h}+{self.winfo_x() + (self.winfo_width()-w)//2}+{self.winfo_y() + (self.winfo_height()-h)//2}")
        self.wait_window(dialog)
        return result[0]

    def _reorder_add_top(self):
        if not self.reorder_model:
            return
        values = self._reorder_prompt("Add Top Level")
        if not values:
            return
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
        siblings[idx], siblings[new_idx] = siblings[new_idx], siblings[idx]
        self._reorder_refresh(node)

    def _reorder_remove(self):
        node = self._reorder_selected_node()
        if not node or not node.parent:
            return
        siblings = node.parent.children
        idx = siblings.index(node)
        siblings.pop(idx)
        self.reorder_undo_stack.append((node.parent, node, idx))
        self._reorder_refresh()

    def _reorder_undo(self):
        if not self.reorder_undo_stack:
            return
        parent, node, idx = self.reorder_undo_stack.pop()
        node.parent = parent
        parent.children.insert(idx, node)
        self._reorder_refresh(node)

    def _reorder_expand(self, expand):
        if not self._reorder_tree_available():
            return
        def toggle(item, force_open=False):
            self.reorder_tree.item(item, open=(True if force_open else expand))
            for child in self.reorder_tree.get_children(item):
                toggle(child)
        for item in self.reorder_tree.get_children():
            toggle(item, force_open=not expand)

    # ── Reference Download ────────────────────────────────────────────────────

    def _reference_page(self):
        main = self._scrollable_main()
        color = WORKFLOW_META["reference_download"]["color"]
        self._page_header(main, "Drawing Downloader",
                          "Download drawings referenced in a structure workbook by part number or direct URL.", color)

        card = self._card(main)
        self._section_label(card, "Inputs")

        structure_var = tk.StringVar()
        output_var = tk.StringVar()

        self._field(card, "Structure workbook (.xlsx)", structure_var,
                    lambda: structure_var.set(
                        filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")]) or structure_var.get()))

        self._divider(card)
        self._section_label(card, "Output")
        self._field(card, "Download folder", output_var,
                    lambda: output_var.set(filedialog.askdirectory() or output_var.get()),
                    "Browse folder…")

        def run():
            if not structure_var.get() or not output_var.get():
                messagebox.showwarning("Missing fields", "Choose structure file and output folder.", parent=self)
                return
            progress = ProgressDialog(self, "Downloading Drawings")
            result = None
            error = None
            try:
                result = download_references(structure_var.get(), output_var.get(), progress_callback=progress.update)
            except Exception as exc:
                error = str(exc)
            finally:
                progress.close()
            if error:
                self._show_themed_dialog("Download failed", error, tone="error")
                return
            self._show_themed_dialog(
                "Download complete",
                f"Downloaded:  {len(result['downloaded'])}\n"
                f"Not found:   {len(result['missing_parts'])}\n"
                f"Failed:      {len(result['failed'])}\n\n"
                f"Not found:\n{summarize_list(result['missing_parts'])}\n\n"
                f"Failed:\n{summarize_list(result['failed'])}",
                tone="info",
            )

        self._divider(card)
        self._run_btn(card, "Download Drawings", run, color).pack(anchor="w")

    # ── Shared file browse helper ─────────────────────────────────────────────

    def _browse_file(self, input_var, output_var, suffix, ext, filetypes):
        path = filedialog.askopenfilename(filetypes=filetypes)
        if not path:
            return
        input_var.set(path)
        if not output_var.get().strip():
            output_var.set(default_output_path(path, suffix, ext))


def main():
    app = DrawingCompilerStudio()
    app.mainloop()


if __name__ == "__main__":
    main()
