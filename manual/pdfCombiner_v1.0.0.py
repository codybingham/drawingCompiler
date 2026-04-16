import os
import re
from io import BytesIO
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

import pandas as pd
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


APP_VERSION = "1.0.0"


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


def is_hydraulic_schematic_entry(description):
    return str(description).strip().upper().startswith("HYDRAULIC SCHEMATIC")


def _build_index_entries(entries):
    grouped = {}
    for i, entry in enumerate(entries):
        key = entry["desc"].strip().casefold()
        if key not in grouped:
            grouped[key] = {
                "desc": entry["desc"].strip(),
                "part": "",
                "indent_level": 0,
                "toc_indices": [],
            }
        grouped[key]["toc_indices"].append(i)

    return sorted(grouped.values(), key=lambda entry: entry["desc"].casefold())


def _layout_directory_entries(entries, is_index=False):
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
    for idx, entry in enumerate(entries):
        per_page_index = idx % rows_per_page
        page_index = idx // rows_per_page
        col_index = per_page_index // rows_per_column
        row_index = per_page_index % rows_per_column

        column_left = column_lefts[col_index]
        column_right = column_left + column_width
        title_y = height - margin_top
        y = title_y - 0.45 * inch - row_index * row_height
        desc_x = column_left + (0 if is_index else entry["indent_level"] * indent_step)
        page_x = column_right - 4
        page_left_x = column_right - (0.8 * inch)
        part_x = page_left_x - 6

        placements.append(
            {
                "entry_index": idx,
                "page_index": page_index,
                "desc_x": desc_x,
                "part_x": part_x,
                "page_x": page_x,
                "page_left_x": page_left_x,
                "y": y,
                "title_y": title_y,
            }
        )

    total_pages = (len(entries) + rows_per_page - 1) // rows_per_page if entries else 1
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


def create_directory_pdf_bytes(entries, title, page_offset_map=None, is_index=False):
    packet = BytesIO()
    page_size, placements, total_pages = _layout_directory_entries(entries, is_index=is_index)
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
        part = entry["part"]
        display_part = part if part and not is_hydraulic_schematic_entry(desc) else ""

        entry_index = placement["entry_index"]
        page_num = ""
        if "page_text" in entry and entry["page_text"]:
            page_num = entry["page_text"]
        elif page_offset_map is not None and page_offset_map[entry_index] is not None:
            page_num = str(page_offset_map[entry_index] + 1)

        c.setFont("Helvetica", 8)
        desc_right_limit = placement["page_left_x"] - 8 if is_index or not display_part else placement["part_x"] - 8
        desc = _trim_text_to_width(desc, "Helvetica", 8, desc_right_limit - placement["desc_x"])
        c.drawString(placement["desc_x"], placement["y"], desc)
        if display_part:
            part_text = f"[{display_part}]"
            part_text = _trim_text_to_width(part_text, "Helvetica", 8, placement["page_left_x"] - placement["part_x"] - 4)
            c.drawRightString(placement["part_x"], placement["y"], part_text)
        if page_num:
            c.drawRightString(placement["page_x"], placement["y"], page_num)

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


def add_toc_hyperlinks(writer, toc_placements, effective_page_map, line_height=12):
    for placement in toc_placements:
        entry_index = placement["entry_index"]
        target_page = effective_page_map[entry_index]
        if target_page is None:
            continue

        _add_internal_link_annotation(
            writer,
            from_page_index=placement["page_index"],
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


def find_column(df, target_names):
    normalized = {str(col).strip().lower(): col for col in df.columns}
    for name in target_names:
        if name.lower() in normalized:
            return normalized[name.lower()]
    raise KeyError(f"Could not find any of these columns: {target_names}")


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


def main():
    root = tk.Tk()
    root.withdraw()

    base_folder = filedialog.askdirectory(
        title=f"Drawing Packet Builder v{APP_VERSION} - Select folder containing the Excel file and drawing PDFs"
    )

    if not base_folder:
        raise Exception("No folder selected.")

    drawing_folder = base_folder

    excel_files = sorted(
        [
            f for f in os.listdir(base_folder)
            if f.lower().endswith((".xlsx", ".xlsm", ".xls"))
            and not f.startswith("~$")
        ]
    )

    if len(excel_files) == 0:
        raise FileNotFoundError("No Excel file found in the selected folder.")

    if len(excel_files) == 1:
        structure_filename = excel_files[0]
    else:
        selection_text = "\n".join(
            [f"{i + 1}. {fname}" for i, fname in enumerate(excel_files)]
        )
        selection = simpledialog.askstring(
            "Select Excel File",
            "Multiple Excel files found.\n\n"
            f"{selection_text}\n\n"
            "Enter the number of the Excel file to use:",
            parent=root,
        )

        if selection is None:
            raise Exception("Excel file selection cancelled.")

        if not selection.isdigit():
            raise Exception("Invalid Excel file selection.")

        selection_index = int(selection) - 1
        if selection_index < 0 or selection_index >= len(excel_files):
            raise Exception("Excel file selection out of range.")

        structure_filename = excel_files[selection_index]

    structure_file = os.path.join(base_folder, structure_filename)

    output_name = simpledialog.askstring(
        "Output File Name",
        "Enter output PDF file name:",
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
    output_pdf = os.path.join(base_folder, output_name)

    df = pd.read_excel(structure_file)

    part_col = find_column(df, ["Part Number", "PartNumber", "Part #", "Part"])
    desc_col = find_column(df, ["Description", "Desc"])
    level_col = find_column(df, ["Level", "Level Code", "Structure Level"])

    df = df.dropna(subset=[part_col])

    raw_entries = []

    for _, row in df.iterrows():
        part = str(row[part_col]).strip()
        desc = str(row[desc_col]).strip()
        code_text = str(row[level_col]).strip()
        code_tuple = parse_level_code(row[level_col])

        raw_entries.append(
            {
                "code_text": code_text,
                "code_tuple": code_tuple,
                "desc": desc,
                "part": part,
                "filename": f"{part}.pdf",
            }
        )

    toc_entries = build_hierarchy(raw_entries)

    existing_entries = []
    missing_files = []

    for entry in toc_entries:
        fpath = os.path.join(drawing_folder, entry["filename"])
        if os.path.exists(fpath):
            existing_entries.append(entry)
        else:
            missing_files.append(entry["filename"])

    if not existing_entries:
        raise Exception("No drawing PDFs found in the selected folder.")

    index_entries = _build_index_entries(existing_entries)
    toc_entries = existing_entries + [
        {
            "desc": "Index",
            "part": "",
            "indent_level": 0,
        }
    ]
    toc_packet, _, toc_pages = create_directory_pdf_bytes(toc_entries, "Table of Contents", None)
    index_packet, _, index_pages = create_directory_pdf_bytes(index_entries, "Index", None, is_index=True)

    page_offset_map = []
    current_page = toc_pages

    for entry in existing_entries:
        fpath = os.path.join(drawing_folder, entry["filename"])
        reader = PdfReader(fpath)
        page_offset_map.append(current_page)
        current_page += len(reader.pages)
    index_start_page = current_page
    toc_page_map = page_offset_map + [index_start_page]

    toc_packet, toc_placements, _ = create_directory_pdf_bytes(
        toc_entries,
        "Table of Contents",
        toc_page_map,
    )
    for entry in index_entries:
        pages = []
        for toc_index in entry["toc_indices"]:
            page = page_offset_map[toc_index]
            if page is not None and page not in pages:
                pages.append(page)
        entry["page_text"] = ", ".join(str(page + 1) for page in pages)
    index_packet, _, _ = create_directory_pdf_bytes(index_entries, "Index", None, is_index=True)
    toc_reader = PdfReader(toc_packet)
    index_reader = PdfReader(index_packet)

    writer = PdfWriter()

    for page in toc_reader.pages:
        writer.add_page(page)

    entry_start_page = {}
    for entry in existing_entries:
        fpath = os.path.join(drawing_folder, entry["filename"])
        reader = PdfReader(fpath)

        start_page = len(writer.pages)
        entry_start_page[id(entry)] = start_page

        for page in reader.pages:
            writer.add_page(page)

    for page in index_reader.pages:
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

        bookmark = writer.add_outline_item(
            entry["desc"],
            start_page,
            parent=parent,
        )
        bookmark_refs[id(entry)] = bookmark

    add_toc_hyperlinks(writer, toc_placements, toc_page_map)

    total_pages = len(writer.pages)

    for i, page in enumerate(writer.pages):
        add_page_number_overlay(page, i + 1, total_pages)

    with open(output_pdf, "wb") as f:
        writer.write(f)

    message = (
        f"Drawing Packet Builder v{APP_VERSION}\n\n"
        f"Excel file used: {structure_filename}\n"
        f"Created: {output_pdf}"
    )

    if missing_files:
        message += "\n\nMissing files:\n" + "\n".join(f"  - {fname}" for fname in missing_files)

    print(message)
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
