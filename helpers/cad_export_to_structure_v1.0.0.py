import argparse
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd


OUTPUT_COLUMNS = ["Level", "Description", "Part Number"]


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


def read_cad_export(path: str) -> pd.DataFrame:
    _, ext = os.path.splitext(path.lower())
    if ext in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(path)
    if ext == ".csv":
        return pd.read_csv(path)
    raise ValueError("Unsupported input file type. Use .xlsx, .xlsm, .xls, or .csv")


def convert_to_structure(df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str | None]]:
    level_col = find_column(
        df,
        [
            "Level",
            "Item Number",
            "Item No",
            "Item",
            "Find Number",
            "Position",
            "Pos",
            "Hierarchy",
        ],
    )
    desc_col = find_column(df, ["Description", "Part Description", "Name", "Title"])
    part_col = find_column(df, ["Part Number", "Part No", "Number", "Item Number", "Part"])

    if not desc_col and not part_col:
        raise ValueError(
            "Could not identify description or part number columns in the CAD export. "
            "Include at least one recognizable Description/Part Number column."
        )

    rows = []
    sequence = 1
    for _, row in df.iterrows():
        level_value = _clean_cell(row[level_col]) if level_col else str(sequence)
        desc_value = _clean_cell(row[desc_col]) if desc_col else ""
        part_value = _clean_cell(row[part_col]) if part_col else ""

        if not level_value:
            level_value = str(sequence)

        if not desc_value and not part_value:
            continue

        rows.append(
            {
                "Level": level_value,
                "Description": desc_value,
                "Part Number": part_value,
            }
        )
        sequence += 1

    if not rows:
        raise ValueError("No usable rows found in CAD export after conversion.")

    out_df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    mapping = {
        "level_col": level_col,
        "description_col": desc_col,
        "part_number_col": part_col,
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
                f"- Level: {mapping['level_col'] or 'Generated sequentially'}",
                f"- Description: {mapping['description_col'] or 'Not found'}",
                f"- Part Number: {mapping['part_number_col'] or 'Not found'}",
            ]
        ),
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert a CAD export (Excel/CSV) into a structure Excel file."
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
