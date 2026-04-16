# Helpers

This folder contains standalone helper tools that work with existing **structure Excel files** (not CAD exports).

## 1) Structure Reorder GUI
- Script: `helpers/structure_reorder_gui_v1.0.0.py`
- Purpose: Open a structure Excel file (`Level`, `Description`, `Part Number`) and reorder the hierarchy before saving a new structure file.

### Features
- Add top-level items, child items, and sibling items.
- Edit item description and part number.
- Move selected level up/down among siblings.
- Remove a level and its children.
- Undo remove (Ctrl+Z).
- Save reordered output to a new Excel file.

### Run
```bash
python helpers/structure_reorder_gui_v1.0.0.py
```

## 2) Structure Reference Downloader
- Script: `helpers/structure_reference_downloader_v1.0.0.py`
- Purpose: Read a structure Excel file and download all referenced files.

### Supported references
- **Part Number / Item Number** columns:
  - Uses internal lookup service (`prints.spudnik.local`) to find drawing PDFs.
  - Downloads as `<part_number>.pdf`.
- **URL/Link/Path** columns:
  - Downloads HTTP/HTTPS links directly.

### Run
```bash
python helpers/structure_reference_downloader_v1.0.0.py
```


## 3) CAD Export to Structure Converter
- Script: `helpers/cad_export_to_structure_v1.0.0.py`
- Purpose: Convert a CAD export (`.xlsx`/`.csv`) into a structure Excel file with columns:
  - `Level`
  - `Description`
  - `Part Number`

### Notes
- Uses the same row-preservation and hierarchy rules as the automated compiler's CAD parsing flow.
- Requires CAD columns equivalent to: `Object`, `Name`, and `Item Number`.
- Keeps matching part rows (`13*`, `FB*`, `HA*` with exclusions), preserves required parents, and writes levels like `1`, `1.1`, `1.1.1`.
- Supports both CLI and file-picker workflow.

### Run
```bash
python helpers/cad_export_to_structure_v1.0.0.py
```

### CLI Example
```bash
python helpers/cad_export_to_structure_v1.0.0.py --input /path/to/cad_export.xlsx --output /path/to/output_structure.xlsx
```

## Requirements
These helpers rely on the same Python dependencies used in this repo:
- `pandas`
- `requests`
- `openpyxl` (for `.xlsx` I/O through pandas)
- `tkinter` (standard with most Python installs)
