import os
import re
import tkinter as tk
from dataclasses import dataclass
from tkinter import filedialog, messagebox, ttk
from typing import Callable

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

    return {
        "input_path": input_path,
        "output_path": output_path,
        "source_rows": len(df),
        "rows_written": len(out_df),
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


def download_references(
    structure_path: str,
    output_folder: str,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> dict:
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
            progress_callback(completed_downloads, total_downloads, f"Downloaded part: {part}")

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
            progress_callback(completed_downloads, total_downloads, f"Downloaded URL file: {filename}")

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


def build_manual_packet(
    structure_path: str,
    drawings_folder: str,
    output_pdf: str,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> dict:
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

    if progress_callback:
        progress_callback(0, len(entries), "Scanning structure entries...")

    for idx, entry in enumerate(entries, start=1):
        part = entry["part"]
        if not part:
            continue
        pdf_path = _find_pdf_for_part(drawings_folder, part)
        if not pdf_path:
            missing.append(part)
            if progress_callback:
                progress_callback(idx, len(entries), f"Missing drawing for {part}")
            continue

        reader = PdfReader(pdf_path)
        for page in reader.pages:
            writer.add_page(page)
        included += 1
        if progress_callback:
            progress_callback(idx, len(entries), f"Added {part}")

    if not writer.pages:
        raise ValueError("No PDFs were added. Check your drawings folder and part numbers.")

    os.makedirs(os.path.dirname(output_pdf) or ".", exist_ok=True)
    with open(output_pdf, "wb") as f:
        writer.write(f)

    return {"output_pdf": output_pdf, "included_parts": included, "missing_parts": missing}


def build_automated_packet(
    structure_path: str,
    temp_download_folder: str,
    output_pdf: str,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> dict:
    def phase_download(completed: int, total: int, message: str) -> None:
        if progress_callback:
            progress_callback(completed, total if total > 0 else 1, f"Download phase: {message}")

    download_result = download_references(structure_path, temp_download_folder, progress_callback=phase_download)

    def phase_build(completed: int, total: int, message: str) -> None:
        if progress_callback:
            progress_callback(completed, total if total > 0 else 1, f"Build phase: {message}")

    packet_result = build_manual_packet(
        structure_path,
        temp_download_folder,
        output_pdf,
        progress_callback=phase_build,
    )
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


@dataclass
class StructureNode:
    level: str
    description: str
    part_number: str
    children: list["StructureNode"]
    parent: "StructureNode | None" = None

    def __init__(self, level: str, description: str, part_number: str) -> None:
        self.level = level
        self.description = description
        self.part_number = part_number
        self.children = []
        self.parent = None

    def add_child(self, child: "StructureNode") -> None:
        child.parent = self
        self.children.append(child)


class StructureModel:
    def __init__(self) -> None:
        self.root = StructureNode(level="", description="ROOT", part_number="")

    @classmethod
    def from_dataframe(cls, df: pd.DataFrame) -> "StructureModel":
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
            rows.append(
                {
                    "level": level,
                    "code": parse_level_code(level),
                    "description": "" if pd.isna(row["Description"]) else str(row["Description"]).strip(),
                    "part": "" if pd.isna(row["Part Number"]) else str(row["Part Number"]).strip(),
                }
            )

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

    def to_dataframe(self) -> pd.DataFrame:
        rows = []

        def walk(nodes: list[StructureNode], prefix: list[int]) -> None:
            for idx, node in enumerate(nodes, start=1):
                level = ".".join(str(x) for x in (prefix + [idx]))
                rows.append({"Level": level, "Description": node.description, "Part Number": node.part_number})
                walk(node.children, prefix + [idx])

        walk(self.root.children, [])
        return pd.DataFrame(rows, columns=["Level", "Description", "Part Number"])


def default_output_path(input_path: str, suffix: str, extension: str) -> str:
    if not input_path:
        return ""
    base = os.path.splitext(os.path.basename(input_path))[0]
    return os.path.join(os.path.dirname(input_path), f"{base}{suffix}{extension}")


def summarize_list(values: list[str], limit: int = 10) -> str:
    if not values:
        return "None"
    shown = values[:limit]
    remaining = len(values) - len(shown)
    lines = [f"- {item}" for item in shown]
    if remaining > 0:
        lines.append(f"... and {remaining} more")
    return "\n".join(lines)


def validate_output_filename(filename: str) -> str:
    invalid_chars_pattern = r'[<>:"/\\|?*]'
    reserved_names = {
        "CON",
        "PRN",
        "AUX",
        "NUL",
        "COM1",
        "COM2",
        "COM3",
        "COM4",
        "COM5",
        "COM6",
        "COM7",
        "COM8",
        "COM9",
        "LPT1",
        "LPT2",
        "LPT3",
        "LPT4",
        "LPT5",
        "LPT6",
        "LPT7",
        "LPT8",
        "LPT9",
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


class ProgressDialog:
    def __init__(self, parent: tk.Misc, title: str) -> None:
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.resizable(False, False)
        self.window.transient(parent)
        self.window.grab_set()

        frame = ttk.Frame(self.window, padding=14)
        frame.pack(fill="both", expand=True)

        self.status_var = tk.StringVar(value="Starting...")
        ttk.Label(frame, textvariable=self.status_var, width=72).pack(anchor="w", pady=(0, 8))

        self.progress_var = tk.DoubleVar(value=0)
        progress = ttk.Progressbar(
            frame,
            orient="horizontal",
            mode="determinate",
            length=520,
            variable=self.progress_var,
            maximum=100,
        )
        progress.pack(fill="x")

        self.window.update_idletasks()

    def update(self, completed: int, total: int, message: str) -> None:
        self.status_var.set(message)
        self.progress_var.set((completed / total) * 100 if total else 0)
        self.parent.update_idletasks()

    def close(self) -> None:
        if self.window.winfo_exists():
            self.window.destroy()


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

    def _set_structure_and_default_output(
        self,
        structure_var: tk.StringVar,
        output_var: tk.StringVar,
        output_suffix: str,
        extension: str,
        filetypes: list[tuple[str, str]],
    ) -> None:
        path = filedialog.askopenfilename(filetypes=filetypes)
        if not path:
            return
        structure_var.set(path)
        if not output_var.get().strip():
            output_var.set(default_output_path(path, output_suffix, extension))

    def _set_input_and_default_output(
        self,
        input_var: tk.StringVar,
        output_var: tk.StringVar,
        output_suffix: str,
        extension: str,
        filetypes: list[tuple[str, str]],
    ) -> None:
        path = filedialog.askopenfilename(filetypes=filetypes)
        if not path:
            return
        input_var.set(path)
        if not output_var.get().strip():
            output_var.set(default_output_path(path, output_suffix, extension))

    def _manual_packet_page(self) -> None:
        card = self._card("Manual Packet Builder", "Build a merged PDF packet from structure rows and local drawing PDFs.")

        structure_var, drawings_var, output_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        form = ttk.Frame(card, style="Card.TFrame")
        form.pack(fill="x", pady=4)

        self._path_row(
            form,
            "Structure workbook",
            structure_var,
            lambda: self._set_structure_and_default_output(
                structure_var,
                output_var,
                "_packet",
                ".pdf",
                [("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")],
            ),
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
                validate_output_filename(output_var.get())
            except ValueError as exc:
                messagebox.showerror("Invalid Output Filename", str(exc), parent=self)
                return
            progress = ProgressDialog(self, "Building Manual Packet")
            try:
                result = build_manual_packet(
                    structure_var.get(),
                    drawings_var.get(),
                    output_var.get(),
                    progress_callback=progress.update,
                )
                messagebox.showinfo(
                    "Packet complete",
                    "\n".join(
                        [
                            f"Output: {result['output_pdf']}",
                            f"Included parts: {result['included_parts']}",
                            f"Missing parts: {len(result['missing_parts'])}",
                            "",
                            "Missing part list:",
                            summarize_list(result["missing_parts"]),
                        ]
                    ),
                    parent=self,
                )
            except Exception as exc:
                messagebox.showerror("Build failed", str(exc), parent=self)
            finally:
                progress.close()

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
            lambda: self._set_structure_and_default_output(
                structure_var,
                output_var,
                "_automated_packet",
                ".pdf",
                [("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")],
            ),
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
                validate_output_filename(output_var.get())
            except ValueError as exc:
                messagebox.showerror("Invalid Output Filename", str(exc), parent=self)
                return
            progress = ProgressDialog(self, "Running Automated Build")
            try:
                result = build_automated_packet(
                    structure_var.get(),
                    download_var.get(),
                    output_var.get(),
                    progress_callback=progress.update,
                )
                messagebox.showinfo(
                    "Automated build complete",
                    "\n".join(
                        [
                            f"Output: {result['output_pdf']}",
                            f"Included: {result['included_parts']}",
                            f"Missing in packet: {len(result['missing_parts'])}",
                            f"Download failures: {len(result['failed_downloads'])}",
                            f"Not found in lookup: {len(result['not_found'])}",
                            "",
                            "Missing in packet:",
                            summarize_list(result["missing_parts"]),
                            "",
                            "Lookup not found:",
                            summarize_list(result["not_found"]),
                        ]
                    ),
                    parent=self,
                )
            except Exception as exc:
                messagebox.showerror("Automated build failed", str(exc), parent=self)
            finally:
                progress.close()

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
            lambda: self._set_input_and_default_output(
                input_var,
                output_var,
                "_structure",
                ".xlsx",
                [("Supported", "*.xlsx *.xlsm *.xls *.csv"), ("All", "*.*")],
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
                mapping = result["mapping"]
                messagebox.showinfo(
                    "Conversion complete",
                    "\n".join(
                        [
                            f"Input: {result['input_path']}",
                            f"Output: {result['output_path']}",
                            f"Rows read: {result['source_rows']}",
                            f"Rows written: {result['rows_written']}",
                            "",
                            "Detected columns:",
                            f"- Object: {mapping['object_col']}",
                            f"- Name: {mapping['name_col']}",
                            f"- Item Number: {mapping['item_number_col']}",
                        ]
                    ),
                    parent=self,
                )
            except Exception as exc:
                messagebox.showerror("Conversion failed", str(exc), parent=self)

        ttk.Button(card, text="Generate Structure", style="Primary.TButton", command=run_conversion).pack(anchor="w", pady=(12, 0))

    def _reorder_page(self) -> None:
        card = self._card("Structure Reorder", "Full editor: reorder, add, edit, remove, undo remove, then save renumbered output.")
        self.reorder_model: StructureModel | None = None
        self.reorder_source_path: str | None = None
        self.reorder_item_lookup: dict[str, StructureNode] = {}
        self.reorder_undo_stack: list[tuple[StructureNode, StructureNode, int]] = []

        tools = ttk.Frame(card, style="Card.TFrame")
        tools.pack(fill="x", pady=(0, 8))

        tree_frame = ttk.Frame(card, style="Card.TFrame")
        tree_frame.pack(fill="both", expand=True)
        self.reorder_tree = ttk.Treeview(tree_frame, columns=("part",), show="tree headings", selectmode="browse")
        self.reorder_tree.heading("#0", text="Description")
        self.reorder_tree.heading("part", text="Part Number")
        self.reorder_tree.column("#0", width=700, anchor="w")
        self.reorder_tree.column("part", width=220, anchor="w")
        self.reorder_tree.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.reorder_tree.yview)
        self.reorder_tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="left", fill="y")

        actions = ttk.Frame(card, style="Card.TFrame")
        actions.pack(fill="x", pady=(8, 0))

        def selected_node() -> StructureNode | None:
            sel = self.reorder_tree.selection()
            return self.reorder_item_lookup.get(sel[0]) if sel else None

        def refresh_tree(select_node: StructureNode | None = None) -> None:
            for item in self.reorder_tree.get_children():
                self.reorder_tree.delete(item)
            self.reorder_item_lookup.clear()
            if not self.reorder_model:
                return

            def add_nodes(parent_id: str, nodes: list[StructureNode]) -> None:
                for node in nodes:
                    item_id = self.reorder_tree.insert(parent_id, "end", text=node.description, values=(node.part_number,), open=True)
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

        def prompt_item_values(title: str, initial_description: str = "", initial_part_number: str = "") -> tuple[str, str] | None:
            dialog = tk.Toplevel(self)
            dialog.title(title)
            dialog.transient(self)
            dialog.grab_set()
            frame = ttk.Frame(dialog, padding=12)
            frame.pack(fill="both", expand=True)
            ttk.Label(frame, text="Description").grid(row=0, column=0, sticky="w")
            desc_var = tk.StringVar(value=initial_description)
            ttk.Entry(frame, textvariable=desc_var, width=60).grid(row=1, column=0, sticky="ew", pady=(0, 8))
            ttk.Label(frame, text="Part Number").grid(row=2, column=0, sticky="w")
            part_var = tk.StringVar(value=initial_part_number)
            ttk.Entry(frame, textvariable=part_var, width=60).grid(row=3, column=0, sticky="ew")
            result: tuple[str, str] | None = None

            def on_ok() -> None:
                nonlocal result
                desc = desc_var.get().strip()
                if not desc:
                    messagebox.showwarning("Missing Description", "Description is required.", parent=dialog)
                    return
                result = (desc, part_var.get().strip())
                dialog.destroy()

            btns = ttk.Frame(frame)
            btns.grid(row=4, column=0, sticky="e", pady=(10, 0))
            ttk.Button(btns, text="Cancel", command=dialog.destroy).pack(side="right")
            ttk.Button(btns, text="OK", command=on_ok).pack(side="right", padx=(0, 8))
            self.wait_window(dialog)
            return result

        def open_file() -> None:
            path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")])
            if not path:
                return
            try:
                df = load_structure_for_reorder(path)
                self.reorder_model = StructureModel.from_dataframe(df)
                self.reorder_source_path = path
                self.reorder_undo_stack.clear()
                refresh_tree()
            except Exception as exc:
                messagebox.showerror("Open failed", str(exc), parent=self)

        def add_top() -> None:
            if not self.reorder_model:
                return
            values = prompt_item_values("Add Top Level")
            if not values:
                return
            node = StructureNode(level="", description=values[0], part_number=values[1])
            self.reorder_model.root.add_child(node)
            refresh_tree(node)

        def add_child() -> None:
            node = selected_node()
            if not node:
                return
            values = prompt_item_values("Add Child")
            if not values:
                return
            new_node = StructureNode(level="", description=values[0], part_number=values[1])
            node.add_child(new_node)
            refresh_tree(new_node)

        def add_sibling() -> None:
            node = selected_node()
            if not node or not node.parent:
                return
            values = prompt_item_values("Add Sibling")
            if not values:
                return
            siblings = node.parent.children
            idx = siblings.index(node) + 1
            new_node = StructureNode(level="", description=values[0], part_number=values[1])
            new_node.parent = node.parent
            siblings.insert(idx, new_node)
            refresh_tree(new_node)

        def edit_item() -> None:
            node = selected_node()
            if not node:
                return
            values = prompt_item_values("Edit Item", node.description, node.part_number)
            if not values:
                return
            node.description, node.part_number = values
            refresh_tree(node)

        def move(delta: int) -> None:
            node = selected_node()
            if not node or not node.parent:
                return
            siblings = node.parent.children
            idx = siblings.index(node)
            new_idx = idx + delta
            if new_idx < 0 or new_idx >= len(siblings):
                return
            siblings[idx], siblings[new_idx] = siblings[new_idx], siblings[idx]
            refresh_tree(node)

        def remove_item() -> None:
            node = selected_node()
            if not node or not node.parent:
                return
            siblings = node.parent.children
            idx = siblings.index(node)
            siblings.pop(idx)
            self.reorder_undo_stack.append((node.parent, node, idx))
            refresh_tree()

        def undo_remove() -> None:
            if not self.reorder_undo_stack:
                return
            parent, node, idx = self.reorder_undo_stack.pop()
            node.parent = parent
            parent.children.insert(idx, node)
            refresh_tree(node)

        def save_file() -> None:
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
                messagebox.showinfo("Saved", f"Saved reordered structure:\n{path}", parent=self)
            except Exception as exc:
                messagebox.showerror("Save failed", str(exc), parent=self)

        def expand_all() -> None:
            for item in self.reorder_tree.get_children():
                self.reorder_tree.item(item, open=True)
                for child in self.reorder_tree.get_children(item):
                    self.reorder_tree.item(child, open=True)

        def collapse_all() -> None:
            for item in self.reorder_tree.get_children():
                self.reorder_tree.item(item, open=False)

        ttk.Button(tools, text="Open Structure", command=open_file).pack(side="left")
        ttk.Button(tools, text="Save As", command=save_file).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Add Top Level", command=add_top).pack(side="left")
        ttk.Button(actions, text="Add Child", command=add_child).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Add Sibling", command=add_sibling).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Edit Item", command=edit_item).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Move Up", command=lambda: move(-1)).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Move Down", command=lambda: move(1)).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Remove", command=remove_item).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Undo Remove", command=undo_remove).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Expand All", command=expand_all).pack(side="left", padx=(20, 0))
        ttk.Button(actions, text="Collapse All", command=collapse_all).pack(side="left", padx=(8, 0))

    def _reference_page(self) -> None:
        card = self._card("Reference Downloader", "Download drawings by part/URL references from a structure workbook.")

        structure_var, output_var = tk.StringVar(), tk.StringVar()
        form = ttk.Frame(card, style="Card.TFrame")
        form.pack(fill="x", pady=4)

        self._path_row(
            form,
            "Structure workbook",
            structure_var,
            lambda: structure_var.set(
                filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")]) or structure_var.get()
            ),
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
            progress = ProgressDialog(self, "Downloading References")
            try:
                result = download_references(structure_var.get(), output_var.get(), progress_callback=progress.update)
                messagebox.showinfo(
                    "Download complete",
                    "\n".join(
                        [
                            f"Downloaded: {len(result['downloaded'])}",
                            f"Not found: {len(result['missing_parts'])}",
                            f"Failed: {len(result['failed'])}",
                            "",
                            "Part numbers not found:",
                            summarize_list(result["missing_parts"]),
                            "",
                            "Failed downloads:",
                            summarize_list(result["failed"]),
                        ]
                    ),
                    parent=self,
                )
            except Exception as exc:
                messagebox.showerror("Download failed", str(exc), parent=self)
            finally:
                progress.close()

        ttk.Button(card, text="Download References", style="Primary.TButton", command=run_download).pack(anchor="w", pady=(12, 0))


def main() -> None:
    app = DrawingCompilerStudio()
    app.mainloop()


if __name__ == "__main__":
    main()
