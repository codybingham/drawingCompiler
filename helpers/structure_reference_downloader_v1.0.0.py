import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

LOOKUP_URL = "http://prints.spudnik.local/api/prints/format-paths"


def normalize_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).strip().lower())


def find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    normalized_map = {normalize_header(col): col for col in df.columns}
    for candidate in candidates:
        key = normalize_header(candidate)
        if key in normalized_map:
            return normalized_map[key]
    return None


def read_structure_references(structure_path: str) -> tuple[list[str], list[str]]:
    df = pd.read_excel(structure_path)

    part_col = find_column(df, ["Part Number", "Item Number", "Part", "Item"])
    url_col = find_column(df, ["File URL", "Url", "PDF URL", "Link", "Path"])

    if not part_col and not url_col:
        raise ValueError(
            "No supported reference columns found. Include either Part Number/Item Number "
            "or URL/Link/Path columns."
        )

    part_numbers: set[str] = set()
    urls: set[str] = set()

    if part_col:
        for value in df[part_col].tolist():
            if pd.isna(value):
                continue
            text = str(value).strip()
            if not text:
                continue
            part_numbers.add(text)

    if url_col:
        for value in df[url_col].tolist():
            if pd.isna(value):
                continue
            text = str(value).strip()
            if not text:
                continue
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

    missing = [str(value).strip() for value in data.get("notFound", []) if str(value).strip()]
    return found, sorted(set(missing))


def download_url(session: requests.Session, url: str, out_path: str) -> None:
    response = session.get(url, timeout=90, verify=False)
    response.raise_for_status()
    with open(out_path, "wb") as f:
        f.write(response.content)


def run_download(structure_path: str, output_folder: str) -> dict:
    os.makedirs(output_folder, exist_ok=True)

    part_numbers, direct_urls = read_structure_references(structure_path)

    session = requests.Session()

    found_paths, missing_parts = lookup_print_paths(session, part_numbers)

    downloaded = []
    failed = []

    for part, url in found_paths.items():
        out_path = os.path.join(output_folder, f"{part}.pdf")
        try:
            download_url(session, url, out_path)
            downloaded.append(part)
        except Exception:
            failed.append(part)

    for url in direct_urls:
        filename = os.path.basename(url.split("?", 1)[0]) or "downloaded_file"
        out_path = os.path.join(output_folder, filename)
        try:
            download_url(session, url, out_path)
            downloaded.append(url)
        except Exception:
            failed.append(url)

    return {
        "part_numbers_total": len(part_numbers),
        "urls_total": len(direct_urls),
        "downloaded": downloaded,
        "missing_parts": missing_parts,
        "failed": failed,
        "output_folder": output_folder,
    }


def main() -> None:
    root = tk.Tk()
    root.withdraw()

    structure_path = filedialog.askopenfilename(
        title="Select structure Excel file",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
    )
    if not structure_path:
        return

    output_folder = filedialog.askdirectory(
        title="Select download output folder",
        mustexist=False,
        initialdir=os.path.dirname(structure_path),
    )
    if not output_folder:
        return

    try:
        result = run_download(structure_path, output_folder)
    except Exception as exc:
        messagebox.showerror("Download Error", str(exc))
        return

    summary = [
        f"Output folder: {result['output_folder']}",
        f"Part references found: {result['part_numbers_total']}",
        f"Direct URLs found: {result['urls_total']}",
        f"Successful downloads: {len(result['downloaded'])}",
    ]

    if result["missing_parts"]:
        summary.append("")
        summary.append("Part numbers not found:")
        summary.extend(f"- {part}" for part in result["missing_parts"][:20])

    if result["failed"]:
        summary.append("")
        summary.append("Failed downloads:")
        summary.extend(f"- {value}" for value in result["failed"][:20])

    messagebox.showinfo("Download Complete", "\n".join(summary))


if __name__ == "__main__":
    main()
