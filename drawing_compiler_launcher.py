import os
import re
import tkinter as tk
from dataclasses import dataclass
from tkinter import filedialog, messagebox, ttk

import pandas as pd
import requests
import urllib3
from pypdf import PdfReader, PdfWriter

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

LOOKUP_URL = "http://prints.spudnik.local/api/prints/format-paths"
EXCLUDED_ITEMS = {"HA0814", "HA0815", "HA0816", "HA1129", "HA0817", "984398"}


@dataclass(frozen=True)
class Workflow:
    key: str
    title: str
    subtitle: str


WORKFLOWS = [
    Workflow("dashboard", "Dashboard", "Unified control center"),
    Workflow("manual_packet", "Manual Packet Builder", "Build a packet from local PDFs + structure"),
    Workflow("automated_packet", "Automated Packet Builder", "Download + build packet in one flow"),
    Workflow("cad_to_structure", "CAD Export to Structure", "Convert CAD exports to structure format"),
    Workflow("reorder_structure", "Structure Reorder", "Edit row order and renumber levels"),
    Workflow("reference_download", "Reference Downloader", "Download drawing references from structure"),
]


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
        if token.isdigit():
            output.append(int(token))
        else:
            output.append(token)
    return tuple(output)


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
        rows.append(
            {
                "source_index": source_index,
                "indent": indent,
                "Description": description,
                "Part Number": item_number,
                "keep": False,
                "direct_match": is_valid_item_number(item_number),
                "skippable": is_skippable_nonpart_row(object_value, description),
            }
        )

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
        output_rows.append(
            {
                "Level": ".".join(parts),
                "Description": row["Description"],
                "Part Number": row["Part Number"],
            }
        )

    out_df = pd.DataFrame(output_rows, columns=["Level", "Description", "Part Number"])
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    out_df.to_excel(output_path, index=False)

    return {"output_path": output_path, "rows_written": len(out_df)}


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


def lookup_print_paths(session: requests.Session, part_numbers: list[str]) -> tuple[dict[str, str], list[str]]:
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


def download_url(session: requests.Session, url: str, out_path: str) -> None:
    response = session.get(url, timeout=90, verify=False)
    response.raise_for_status()
    with open(out_path, "wb") as f:
        f.write(response.content)


def download_references(structure_path: str, output_folder: str) -> dict:
    os.makedirs(output_folder, exist_ok=True)
    part_numbers, direct_urls = read_structure_references(structure_path)

    session = requests.Session()
    found_paths, missing_parts = lookup_print_paths(session, part_numbers)

    downloaded = []
    failed = []

    for part, url in found_paths.items():
        target = os.path.join(output_folder, f"{part}.pdf")
        try:
            download_url(session, url, target)
            downloaded.append(part)
        except Exception:
            failed.append(part)

    for url in direct_urls:
        filename = os.path.basename(url.split("?", 1)[0]) or "downloaded_file"
        target = os.path.join(output_folder, filename)
        try:
            download_url(session, url, target)
            downloaded.append(url)
        except Exception:
            failed.append(url)

    return {
        "downloaded": downloaded,
        "missing_parts": missing_parts,
        "failed": failed,
        "output_folder": output_folder,
    }


def _find_pdf_for_part(folder: str, part_number: str) -> str | None:
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


def build_manual_packet(structure_path: str, drawings_folder: str, output_pdf: str) -> dict:
    df = pd.read_excel(structure_path)
    level_col = find_column(df, ["Level"])
    desc_col = find_column(df, ["Description", "Name"])
    part_col = find_column(df, ["Part Number", "Item Number", "Part", "Item"])

    if not level_col or not desc_col or not part_col:
        raise ValueError("Structure file must include Level, Description, and Part Number columns")

    entries = []
    for _, row in df.iterrows():
        level = "" if pd.isna(row[level_col]) else str(row[level_col]).strip()
        desc = "" if pd.isna(row[desc_col]) else str(row[desc_col]).strip()
        part = "" if pd.isna(row[part_col]) else str(row[part_col]).strip()
        if not level:
            continue
        entries.append({"code": parse_level_code(level), "desc": desc, "part": part})

    entries.sort(key=lambda e: e["code"])

    writer = PdfWriter()
    missing = []
    included = 0

    for entry in entries:
        part = entry["part"]
        if not part:
            continue
        pdf_path = _find_pdf_for_part(drawings_folder, part)
        if not pdf_path:
            missing.append(part)
            continue

        reader = PdfReader(pdf_path)
        for page in reader.pages:
            writer.add_page(page)
        included += 1

    if not writer.pages:
        raise ValueError("No PDFs were added. Check your drawings folder and part numbers.")

    os.makedirs(os.path.dirname(output_pdf) or ".", exist_ok=True)
    with open(output_pdf, "wb") as f:
        writer.write(f)

    return {"output_pdf": output_pdf, "included_parts": included, "missing_parts": missing}


def build_automated_packet(structure_path: str, temp_download_folder: str, output_pdf: str) -> dict:
    download_result = download_references(structure_path, temp_download_folder)
    packet_result = build_manual_packet(structure_path, temp_download_folder, output_pdf)
    return {
        "output_pdf": packet_result["output_pdf"],
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
    out = pd.DataFrame(
        {
            "Level": df[level_col].fillna("").astype(str),
            "Description": df[desc_col].fillna("").astype(str),
            "Part Number": df[part_col].fillna("").astype(str),
        }
    )
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

        output.append(
            {
                "Level": new_level,
                "Description": row["Description"],
                "Part Number": row["Part Number"],
            }
        )
    return pd.DataFrame(output, columns=["Level", "Description", "Part Number"])


class DrawingCompilerStudio(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Drawing Compiler Studio - Single File Edition")
        self.geometry("1200x780")
        self.minsize(1080, 700)
        self.history: list[str] = []
        self.current = "dashboard"
        self.reorder_df: pd.DataFrame | None = None

        self._configure_theme()
        self._build_shell()
        self.show_workflow("dashboard", add_history=False)

    def _configure_theme(self) -> None:
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        bg, panel, card = "#0b1220", "#121a2b", "#1c2740"
        text, muted = "#e5e7eb", "#9ca3af"

        self.configure(bg=bg)
        style.configure("App.TFrame", background=bg)
        style.configure("Sidebar.TFrame", background=panel)
        style.configure("Main.TFrame", background=bg)
        style.configure("Card.TFrame", background=card)
        style.configure("H1.TLabel", background=bg, foreground=text, font=("Segoe UI", 24, "bold"))
        style.configure("H2.TLabel", background=card, foreground=text, font=("Segoe UI", 15, "bold"))
        style.configure("Body.TLabel", background=card, foreground=muted, font=("Segoe UI", 10))
        style.configure("SideTitle.TLabel", background=panel, foreground=text, font=("Segoe UI", 14, "bold"))
        style.configure("SideSub.TLabel", background=panel, foreground=muted, font=("Segoe UI", 9))
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=(12, 9))

    def _build_shell(self) -> None:
        shell = ttk.Frame(self, style="App.TFrame", padding=12)
        shell.pack(fill="both", expand=True)

        self.sidebar = ttk.Frame(shell, style="Sidebar.TFrame", padding=14)
        self.sidebar.pack(side="left", fill="y")

        self.main = ttk.Frame(shell, style="Main.TFrame", padding=(16, 10))
        self.main.pack(side="left", fill="both", expand=True)

        ttk.Label(self.sidebar, text="Drawing Compiler", style="SideTitle.TLabel").pack(anchor="w")
        ttk.Label(self.sidebar, text="Single-File Studio", style="SideSub.TLabel").pack(anchor="w", pady=(0, 12))

        for wf in WORKFLOWS:
            ttk.Button(
                self.sidebar,
                text=wf.title,
                command=lambda key=wf.key: self.show_workflow(key),
                width=30,
            ).pack(fill="x", pady=3)

        self.back_btn = ttk.Button(self.sidebar, text="◀ Back", command=self.go_back)
        self.back_btn.pack(side="bottom", fill="x", pady=(12, 0))
        self.back_btn.configure(state="disabled")

    def _clear_main(self) -> None:
        for child in self.main.winfo_children():
            child.destroy()

    def _push(self) -> None:
        self.history.append(self.current)
        self.back_btn.configure(state="normal")

    def go_back(self) -> None:
        if not self.history:
            return
        previous = self.history.pop()
        self.show_workflow(previous, add_history=False)
        self.back_btn.configure(state="normal" if self.history else "disabled")

    def show_workflow(self, key: str, add_history: bool = True) -> None:
        if add_history and key != self.current:
            self._push()
        self.current = key
        self._clear_main()

        if key == "dashboard":
            self._dashboard()
        elif key == "manual_packet":
            self._manual_packet_page()
        elif key == "automated_packet":
            self._automated_packet_page()
        elif key == "cad_to_structure":
            self._cad_page()
        elif key == "reorder_structure":
            self._reorder_page()
        elif key == "reference_download":
            self._reference_page()

    def _card(self, title: str, subtitle: str) -> ttk.Frame:
        card = ttk.Frame(self.main, style="Card.TFrame", padding=16)
        card.pack(fill="both", expand=True)
        ttk.Label(card, text=title, style="H2.TLabel").pack(anchor="w")
        ttk.Label(card, text=subtitle, style="Body.TLabel", wraplength=860, justify="left").pack(anchor="w", pady=(4, 12))
        return card

    def _dashboard(self) -> None:
        ttk.Label(self.main, text="Drawing Compiler Studio", style="H1.TLabel").pack(anchor="w", pady=(4, 10))
        card = self._card(
            "One-file unified application",
            "All workflows are implemented directly in this file. No external script loading is required.",
        )
        grid = ttk.Frame(card, style="Card.TFrame")
        grid.pack(fill="both", expand=True)
        for i, wf in enumerate(WORKFLOWS[1:]):
            item = ttk.Frame(grid, style="Card.TFrame", padding=10)
            item.grid(row=i // 2, column=i % 2, sticky="nsew", padx=8, pady=8)
            ttk.Label(item, text=wf.title, style="H2.TLabel").pack(anchor="w")
            ttk.Label(item, text=wf.subtitle, style="Body.TLabel", wraplength=360, justify="left").pack(anchor="w", pady=(4, 8))
            ttk.Button(item, text="Open", style="Primary.TButton", command=lambda key=wf.key: self.show_workflow(key)).pack(anchor="w")
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)

    def _path_row(self, parent: ttk.Frame, label: str, var: tk.StringVar, browse_cmd, row: int, browse_text: str = "Browse") -> None:
        ttk.Label(parent, text=label, style="Body.TLabel").grid(row=row, column=0, sticky="w")
        ttk.Entry(parent, textvariable=var, width=88).grid(row=row + 1, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(parent, text=browse_text, command=browse_cmd).grid(row=row + 1, column=1)

    def _manual_packet_page(self) -> None:
        card = self._card("Manual Packet Builder", "Build a merged PDF packet from structure rows and local drawing PDFs.")

        structure_var, drawings_var, output_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        form = ttk.Frame(card, style="Card.TFrame")
        form.pack(fill="x", pady=4)

        self._path_row(
            form,
            "Structure workbook",
            structure_var,
            lambda: structure_var.set(filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")]) or structure_var.get()),
            0,
        )
        self._path_row(
            form,
            "Drawings folder",
            drawings_var,
            lambda: drawings_var.set(filedialog.askdirectory() or drawings_var.get()),
            2,
        )
        self._path_row(
            form,
            "Output PDF",
            output_var,
            lambda: output_var.set(
                filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")]) or output_var.get()
            ),
            4,
        )
        form.columnconfigure(0, weight=1)

        def run_manual() -> None:
            if not structure_var.get() or not drawings_var.get() or not output_var.get():
                messagebox.showwarning("Missing data", "Choose structure file, drawings folder, and output PDF.", parent=self)
                return
            try:
                result = build_manual_packet(structure_var.get(), drawings_var.get(), output_var.get())
                messagebox.showinfo(
                    "Packet complete",
                    f"Output: {result['output_pdf']}\nIncluded parts: {result['included_parts']}\nMissing parts: {len(result['missing_parts'])}",
                    parent=self,
                )
            except Exception as exc:
                messagebox.showerror("Build failed", str(exc), parent=self)

        ttk.Button(card, text="Build Manual Packet", style="Primary.TButton", command=run_manual).pack(anchor="w", pady=(12, 0))

    def _automated_packet_page(self) -> None:
        card = self._card("Automated Packet Builder", "Download references then compile a final merged packet in one run.")

        structure_var, download_var, output_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        form = ttk.Frame(card, style="Card.TFrame")
        form.pack(fill="x", pady=4)

        self._path_row(
            form,
            "Structure workbook",
            structure_var,
            lambda: structure_var.set(filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")]) or structure_var.get()),
            0,
        )
        self._path_row(
            form,
            "Download folder",
            download_var,
            lambda: download_var.set(filedialog.askdirectory() or download_var.get()),
            2,
        )
        self._path_row(
            form,
            "Output PDF",
            output_var,
            lambda: output_var.set(
                filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")]) or output_var.get()
            ),
            4,
        )
        form.columnconfigure(0, weight=1)

        def run_auto() -> None:
            if not structure_var.get() or not download_var.get() or not output_var.get():
                messagebox.showwarning("Missing data", "Choose structure, download folder, and output PDF.", parent=self)
                return
            try:
                result = build_automated_packet(structure_var.get(), download_var.get(), output_var.get())
                messagebox.showinfo(
                    "Automated build complete",
                    f"Output: {result['output_pdf']}\nIncluded: {result['included_parts']}\n"
                    f"Missing in packet: {len(result['missing_parts'])}\nDownload failures: {len(result['failed_downloads'])}",
                    parent=self,
                )
            except Exception as exc:
                messagebox.showerror("Automated build failed", str(exc), parent=self)

        ttk.Button(card, text="Run Automated Build", style="Primary.TButton", command=run_auto).pack(anchor="w", pady=(12, 0))

    def _cad_page(self) -> None:
        card = self._card("CAD Export to Structure", "Convert CAD export spreadsheets into structure format.")
        input_var, output_var = tk.StringVar(), tk.StringVar()
        form = ttk.Frame(card, style="Card.TFrame")
        form.pack(fill="x", pady=4)

        self._path_row(
            form,
            "Input CAD export",
            input_var,
            lambda: input_var.set(
                filedialog.askopenfilename(filetypes=[("Supported", "*.xlsx *.xlsm *.xls *.csv"), ("All", "*.*")]) or input_var.get()
            ),
            0,
        )
        self._path_row(
            form,
            "Output structure workbook",
            output_var,
            lambda: output_var.set(
                filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")]) or output_var.get()
            ),
            2,
        )
        form.columnconfigure(0, weight=1)

        def run_conversion() -> None:
            if not input_var.get() or not output_var.get():
                messagebox.showwarning("Missing data", "Choose input and output files.", parent=self)
                return
            try:
                result = convert_cad_to_structure(input_var.get(), output_var.get())
                messagebox.showinfo("Conversion complete", f"Wrote: {result['output_path']}\nRows: {result['rows_written']}", parent=self)
            except Exception as exc:
                messagebox.showerror("Conversion failed", str(exc), parent=self)

        ttk.Button(card, text="Generate Structure", style="Primary.TButton", command=run_conversion).pack(anchor="w", pady=(12, 0))

    def _reorder_page(self) -> None:
        card = self._card("Structure Reorder", "Load a structure, move rows up/down, then save with renumbered levels.")

        controls = ttk.Frame(card, style="Card.TFrame")
        controls.pack(fill="x")

        tree_frame = ttk.Frame(card, style="Card.TFrame")
        tree_frame.pack(fill="both", expand=True, pady=(10, 0))

        self.reorder_tree = ttk.Treeview(tree_frame, columns=("level", "part"), show="headings", selectmode="browse")
        self.reorder_tree.heading("level", text="Level")
        self.reorder_tree.heading("part", text="Part Number")
        self.reorder_tree.column("level", width=160)
        self.reorder_tree.column("part", width=220)
        self.reorder_tree.pack(side="left", fill="both", expand=True)

        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.reorder_tree.yview)
        self.reorder_tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="left", fill="y")

        def refresh_tree() -> None:
            for item in self.reorder_tree.get_children(""):
                self.reorder_tree.delete(item)
            if self.reorder_df is None:
                return
            for idx, row in self.reorder_df.iterrows():
                text = row["Description"]
                self.reorder_tree.insert("", "end", iid=str(idx), values=(row["Level"], row["Part Number"]), text=text)

        def open_file() -> None:
            path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")])
            if not path:
                return
            try:
                self.reorder_df = load_structure_for_reorder(path)
                self.reorder_source_path = path
                refresh_tree()
            except Exception as exc:
                messagebox.showerror("Open failed", str(exc), parent=self)

        def move(delta: int) -> None:
            if self.reorder_df is None:
                return
            selected = self.reorder_tree.selection()
            if not selected:
                return
            idx = int(selected[0])
            new_idx = idx + delta
            if new_idx < 0 or new_idx >= len(self.reorder_df):
                return
            rows = self.reorder_df.to_dict("records")
            rows[idx], rows[new_idx] = rows[new_idx], rows[idx]
            self.reorder_df = pd.DataFrame(rows)
            refresh_tree()
            self.reorder_tree.selection_set(str(new_idx))

        def save_file() -> None:
            if self.reorder_df is None:
                messagebox.showwarning("No data", "Open a structure file first.", parent=self)
                return
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
            if not path:
                return
            try:
                out_df = renumber_structure(self.reorder_df)
                out_df.to_excel(path, index=False)
                messagebox.showinfo("Saved", f"Saved reordered structure:\n{path}", parent=self)
            except Exception as exc:
                messagebox.showerror("Save failed", str(exc), parent=self)

        ttk.Button(controls, text="Open Structure", command=open_file).pack(side="left")
        ttk.Button(controls, text="Move Up", command=lambda: move(-1)).pack(side="left", padx=(8, 0))
        ttk.Button(controls, text="Move Down", command=lambda: move(1)).pack(side="left", padx=(8, 0))
        ttk.Button(controls, text="Save Reordered", style="Primary.TButton", command=save_file).pack(side="left", padx=(14, 0))

    def _reference_page(self) -> None:
        card = self._card("Reference Downloader", "Download drawings by part/URL references from a structure workbook.")

        structure_var, output_var = tk.StringVar(), tk.StringVar()
        form = ttk.Frame(card, style="Card.TFrame")
        form.pack(fill="x", pady=4)

        self._path_row(
            form,
            "Structure workbook",
            structure_var,
            lambda: structure_var.set(filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")]) or structure_var.get()),
            0,
        )
        self._path_row(
            form,
            "Output folder",
            output_var,
            lambda: output_var.set(filedialog.askdirectory() or output_var.get()),
            2,
        )
        form.columnconfigure(0, weight=1)

        def run_download() -> None:
            if not structure_var.get() or not output_var.get():
                messagebox.showwarning("Missing data", "Choose structure file and output folder.", parent=self)
                return
            try:
                result = download_references(structure_var.get(), output_var.get())
                messagebox.showinfo(
                    "Download complete",
                    f"Downloaded: {len(result['downloaded'])}\nNot found: {len(result['missing_parts'])}\nFailed: {len(result['failed'])}",
                    parent=self,
                )
            except Exception as exc:
                messagebox.showerror("Download failed", str(exc), parent=self)

        ttk.Button(card, text="Download References", style="Primary.TButton", command=run_download).pack(anchor="w", pady=(12, 0))


def main() -> None:
    app = DrawingCompilerStudio()
    app.mainloop()


if __name__ == "__main__":
    main()
