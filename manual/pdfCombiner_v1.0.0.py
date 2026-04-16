import os
import re
from io import BytesIO
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

import pandas as pd
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas


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


def create_toc_pdf_bytes(toc_entries, page_offset_map=None):
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch

    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=letter)
    width, height = letter

    margin_left = 0.7 * inch
    margin_right = 0.7 * inch
    margin_top = 0.75 * inch
    margin_bottom = 0.6 * inch

    indent_step = 0.28 * inch
    row_height = 0.24 * inch

    title_y = height - margin_top
    page_x = width - margin_right
    y = title_y - 0.45 * inch

    def draw_header():
        nonlocal y
        c.setFont("Helvetica-Bold", 16)
        c.drawString(margin_left, title_y, "Table of Contents")

        c.setLineWidth(0.5)
        c.line(margin_left, title_y - 6, width - margin_right, title_y - 6)

        y = title_y - 0.45 * inch
        c.setFont("Helvetica", 10)

    draw_header()

    for i, entry in enumerate(toc_entries):
        if y < margin_bottom:
            c.showPage()
            draw_header()

        indent = margin_left + entry["indent_level"] * indent_step
        desc = entry["desc"]
        part = entry["part"]
        text = f"{desc} [{part}]"

        page_num = str(page_offset_map[i] + 1) if page_offset_map is not None else ""

        c.setFont("Helvetica", 10)
        c.drawString(indent, y, text)

        text_width = c.stringWidth(text, "Helvetica", 10)
        leader_start = indent + text_width + 6

        if page_num:
            c.setDash(1, 2)
            c.line(leader_start, y + 2, page_x - 18, y + 2)
            c.setDash()
            c.drawRightString(page_x, y, page_num)

        y -= row_height

    c.save()
    packet.seek(0)
    return packet


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

    toc_packet = create_toc_pdf_bytes(existing_entries, None)
    toc_reader = PdfReader(toc_packet)
    toc_pages = len(toc_reader.pages)

    page_offset_map = []
    current_page = toc_pages

    for entry in existing_entries:
        fpath = os.path.join(drawing_folder, entry["filename"])
        reader = PdfReader(fpath)
        page_offset_map.append(current_page)
        current_page += len(reader.pages)

    toc_packet = create_toc_pdf_bytes(existing_entries, page_offset_map)
    toc_reader = PdfReader(toc_packet)

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
