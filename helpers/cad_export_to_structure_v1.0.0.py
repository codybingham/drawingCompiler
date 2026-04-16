import argparse
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd


OUTPUT_COLUMNS = ["Level", "Description", "Part Number"]
EXCLUDED_ITEMS = {
    "HA0814",
    "HA0815",
    "HA0816",
    "HA1129",
    "HA0817",
    "984398",
}


def normalize_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).strip().lower())


def find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    normalized_map = {normalize_header(col): col for col in df.columns}
    for candidate in candidates:
        key = normalize_header(candidate)
        if key in normalized_map:
            return normalized_map[key]
    return None


def _clean_cell(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text:
        return ""

    number_match = re.fullmatch(r"(\d+)\.0+", text)
    if number_match:
        return number_match.group(1)
    return text


def get_indent_level(object_value) -> int:
    text = "" if object_value is None else str(object_value)
    leading_spaces = len(text) - len(text.lstrip(" "))
    return leading_spaces // 4


def is_valid_item_number(item_number) -> bool:
    text = "" if item_number is None else str(item_number).strip().upper()
    if not text.startswith(("13", "FB", "HA")):
        return False
    if text in EXCLUDED_ITEMS:
        return False
    return True


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
    raise ValueError("Unsupported input file type. Use .xlsx, .xlsm, .xls, or .csv")


def collect_preserved_rows(df: pd.DataFrame, object_col: str, name_col: str, item_col: str) -> list[dict]:
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

    return [row for row in rows if row["keep"]]


def assign_levels(filtered_rows: list[dict]) -> list[dict]:
    counters = {}
    output_rows = []

    for row in filtered_rows:
        indent = row["indent"]
        if indent < 0:
            indent = 0

        keys_to_remove = [k for k in counters if k > indent]
        for key in keys_to_remove:
            counters.pop(key, None)

        counters[indent] = counters.get(indent, 0) + 1

        if indent == 0:
            level_parts = [str(counters[0])]
        else:
            level_parts = [str(counters.get(i, 1)) for i in range(indent + 1)]

        output_rows.append(
            {
                "Level": ".".join(level_parts),
                "Description": row["Description"],
                "Part Number": row["Part Number"],
            }
        )

    return output_rows


def convert_to_structure(df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str]]:
    object_col = find_column(df, ["Object"])
    name_col = find_column(df, ["Name"])
    item_col = find_column(df, ["Item Number", "ItemNumber", "Item No", "Item"])

    if not object_col or not name_col or not item_col:
        raise ValueError(
            "Could not find required CAD columns. Required columns: Object, Name, Item Number."
        )

    filtered_rows = collect_preserved_rows(df, object_col, name_col, item_col)
    if not filtered_rows:
        raise ValueError(
            "No matching rows found. No item numbers started with 13, FB, or HA after exclusions."
        )

    leveled_rows = assign_levels(filtered_rows)
    out_df = pd.DataFrame(leveled_rows, columns=OUTPUT_COLUMNS)
    mapping = {
        "object_col": object_col,
        "name_col": name_col,
        "item_number_col": item_col,
    }
    return out_df, mapping


def run_conversion(input_path: str, output_path: str) -> dict:
    source_df = read_cad_export(input_path)
    structure_df, mapping = convert_to_structure(source_df)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    structure_df.to_excel(output_path, index=False)

    return {
        "input_path": input_path,
        "output_path": output_path,
        "source_rows": len(source_df),
        "output_rows": len(structure_df),
        "mapping": mapping,
    }


def run_gui() -> None:
    root = tk.Tk()
    root.withdraw()

    input_path = filedialog.askopenfilename(
        title="Select CAD export file",
        filetypes=[
            ("Supported files", "*.xlsx *.xlsm *.xls *.csv"),
            ("Excel files", "*.xlsx *.xlsm *.xls"),
            ("CSV files", "*.csv"),
            ("All files", "*.*"),
        ],
    )
    if not input_path:
        return

    base_name = os.path.splitext(os.path.basename(input_path))[0]
    default_output = os.path.join(os.path.dirname(input_path), f"{base_name}_structure.xlsx")

    output_path = filedialog.asksaveasfilename(
        title="Save generated structure file",
        defaultextension=".xlsx",
        initialfile=os.path.basename(default_output),
        initialdir=os.path.dirname(default_output),
        filetypes=[("Excel files", "*.xlsx")],
    )
    if not output_path:
        return

    try:
        result = run_conversion(input_path, output_path)
    except Exception as exc:
        messagebox.showerror("Conversion Error", str(exc))
        return

    mapping = result["mapping"]
    messagebox.showinfo(
        "Conversion Complete",
        "\n".join(
            [
                f"Input: {result['input_path']}",
                f"Output: {result['output_path']}",
                f"Rows read: {result['source_rows']}",
                f"Rows written: {result['output_rows']}",
                "",
                "Detected columns:",
                f"- Object: {mapping['object_col']}",
                f"- Name: {mapping['name_col']}",
                f"- Item Number: {mapping['item_number_col']}",
            ]
        ),
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Generate a structure Excel file from a CAD export using automated compiler hierarchy rules."
    )
    parser.add_argument("--input", dest="input_path", help="Path to CAD export file")
    parser.add_argument("--output", dest="output_path", help="Path for output structure .xlsx")
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if args.input_path and args.output_path:
        result = run_conversion(args.input_path, args.output_path)
        print(f"Wrote structure file: {result['output_path']}")
        print(f"Rows written: {result['output_rows']}")
        return

    run_gui()


if __name__ == "__main__":
    main()
