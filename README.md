# Drawing Compiler Repository

This repository contains tools for working with structured drawing data and PDFs.

## Projects

### 1. Manual Drawing Book Builder
- Location: `manual/`
- Script: `manual/pdfCombiner_v1.0.0.py`
- Documentation: `manual/README.md`

Use this when you already have local drawing PDFs and a structured Excel file that defines hierarchy/order.

### 2. Automated Drawing Downloader
- Location: `automated/`
- Script: `automated/AutomatedpdfCombiner_v1.0.0.py`
- Documentation: `automated/README.md`

Use this when you want to ingest a CAD export + schematic, download drawings, and compile a final packet in the proper structure order with TOC and index.

### 3. Structure Helpers
- Location: `helpers/`
- Scripts:
  - `helpers/structure_reorder_gui_v1.0.0.py`
  - `helpers/structure_reference_downloader_v1.0.0.py`
  - `helpers/cad_export_to_structure_v1.0.0.py`
- Documentation: `helpers/README.md`

Use these when you want to edit existing structure files and/or bulk-download all referenced files.

## Repository Structure

```text
drawingCompiler/
├── README.md
├── automated/
│   ├── AutomatedpdfCombiner_v1.0.0.py
│   ├── README.md
│   └── README.txt
├── helpers/
│   ├── structure_reorder_gui_v1.0.0.py
│   ├── structure_reference_downloader_v1.0.0.py
│   ├── cad_export_to_structure_v1.0.0.py
│   └── README.md
└── manual/
    ├── pdfCombiner_v1.0.0.py
    ├── README.md
    ├── README.txt
    └── ExampleStructure.xlsx
```

## Notes
- `README.md` files are now provided for both tools.
- Original `.txt` readmes are retained for compatibility/reference.
- Internal network dependencies for the automated tool are documented in `automated/README.md`.

## Quick Start
### Drawing Compiler Studio (Recommended)
Run the root program to access all workflows from one cohesive GUI:

```bash
python drawing_compiler_launcher.py
```

The studio is implemented as a single-file program with unified workflow pages, so the project behaves like one cohesive application.
The **Automated Packet Builder** flow now accepts:
- CAD export (`.xlsx/.xls/.csv`)
- schematic PDF
- download folder
- output PDF

It then generates a structure workbook, downloads referenced drawings, and assembles TOC + packet + index in one run.

### Run Individual Tools
1. Pick the tool (`manual` or `automated`) based on your workflow.
2. Review the corresponding Markdown readme for requirements and steps.
3. Run the tool script or packaged executable in your environment.
