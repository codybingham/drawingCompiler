# Drawing Book Builder

## Overview
Drawing Book Builder is a tool for assembling a single, organized PDF drawing book from individual drawing PDFs. The program uses a structured Excel file to define drawing order and hierarchy, and then automatically generates a table of contents (TOC), bookmarks, and page numbers.

This tool streamlines drawing book creation for hydraulic systems and other assemblies so packages are easier to share, review, and archive.

## Features
- Reads a structured Excel file to determine drawing order and hierarchy.
- Merges individual PDF drawings into one PDF.
- Generates a table of contents (TOC) with page numbers.
- Adds PDF bookmarks for quick navigation.
- Numbers all pages in the final PDF.
- Provides a simple interface for non-technical users.

## How to Use

### 1) Prepare files
- Ensure you have a structured Excel file (see format below).
- Gather all individual drawing PDFs referenced in the Excel file.

### 2) Run the program
- Launch the executable (`DrawingBookBuilder.exe`).
- Select the folder containing your PDFs and Excel structure file.

### 3) Generate the drawing book
- The program merges PDFs according to the Excel structure.
- A combined PDF with TOC and bookmarks is created in the output folder.

### 4) Review output
- Open the final PDF and verify order, TOC, bookmarks, and page numbers.

## Excel File Format
The Excel file should include at least these columns:

| Level | Description | Part Number |
|---|---|---|
| 1 | STANDARD HYDRAULICS,1252 | FB9893 |
| 1.1 | MIDDLE AXLE JACKS + STEER, 1252 | HA2526 |
| 1.1.1 | VLV ASY,JACKS,P.O. CHECK,1252 | 137691 |
| 1.1.2 | VLV ASY, JACK FLOW DIV, 1252 | 137688 |

- **Level**: Hierarchical number (for example `1`, `1.1`, `1.1.1`).
- **Description**: Drawing or assembly description.
- **Part Number**: Used to match the PDF file (for example `137691.pdf`).

## Understanding `Level` Numbering
The **Level** column defines the hierarchy, similar to an outline:
- Whole numbers (`1`, `2`, `3`) are top-level assemblies.
- Decimal levels (`1.1`, `1.2`) are sub-assemblies.
- Further decimals (`1.1.1`, `1.1.2`) are components/drawings under sub-assemblies.

Example:
- `1` is the main assembly.
- `1.1` is a sub-assembly within `1`.
- `1.1.1` and `1.1.2` are components within `1.1`.

This hierarchy allows the program to:
- Organize the merged PDF in a logical order.
- Generate TOC entries/bookmarks that reflect assembly structure.
- Make navigation intuitive for reviewers.

## Notes & Troubleshooting
- PDF filenames must match Excel part numbers (for example `137691.pdf`).
- If a drawing is missing, the program reports it in logs.
- For best results, keep PDFs in one folder and ensure they are not password-protected.

## Example

### Input
- `1252DrawingStructure.xlsx` (structure file)
- PDFs such as `FB9893.pdf`, `HA2526.pdf`, `137691.pdf`

### Output
- `1252_DrawingBook.pdf` (combined, organized PDF with TOC and bookmarks)
