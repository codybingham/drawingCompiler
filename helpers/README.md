# Helpers

This folder contains standalone helper tools that work with existing **structure Excel files** (not CAD exports).

## 1) Structure Reorder GUI
- Script: `helpers/structure_reorder_gui_v1.0.0.py`
- Purpose: Open a structure Excel file (`Level`, `Description`, `Part Number`) and reorder the hierarchy before saving a new structure file.

### Features
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

## Requirements
These helpers rely on the same Python dependencies used in this repo:
- `pandas`
- `requests`
- `openpyxl` (for `.xlsx` I/O through pandas)
- `tkinter` (standard with most Python installs)
