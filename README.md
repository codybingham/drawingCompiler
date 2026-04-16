# Drawing Compiler Repository

This repository contains two related Python tools for building compiled drawing PDFs from structured inputs.

## Projects

### 1. Manual Drawing Book Builder
- Location: `manual/`
- Script: `manual/pdfCombiner_v1.0.0.py`
- Documentation: `manual/README.md`

Use this when you already have local drawing PDFs and a structured Excel file that defines hierarchy/order.

### 2. Automated Drawing Packet Builder
- Location: `automated/`
- Script: `automated/AutomatedpdfCombiner_v1.0.0.py`
- Documentation: `automated/README.md`

Use this when you want to assemble one packet from CAD exports and automatically download drawing PDFs from internal services.

## Repository Structure

```text
drawingCompiler/
├── README.md
├── automated/
│   ├── AutomatedpdfCombiner_v1.0.0.py
│   ├── README.md
│   └── README.txt
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
1. Pick the tool (`manual` or `automated`) based on your workflow.
2. Review the corresponding Markdown readme for requirements and steps.
3. Run the tool script or packaged executable in your environment.
