# Drawing Packet Builder

**Version:** 1.0.0

## Overview
This tool creates a compiled drawing packet PDF from one or more CAD-exported Excel files.

In one workflow, it:
- Reads CAD-exported Excel files.
- Preserves assembly structure.
- Allows reorder/removal of levels before generation.
- Allows adding multiple CAD exports into one packet.
- Downloads required drawing PDFs from the internal print lookup system.
- Lets you choose the hydraulic schematic PDF.
- Builds a final compiled PDF with:
  - Table of contents
  - Bookmarks
  - Page numbers

## Prerequisites
Before running:
1. The executable for this tool.
2. At least one CAD-exported Excel file.
3. The hydraulic schematic PDF to use.
4. Network access to:
   - `prints.spudnik.local`
   - `api.spudnik.local`

If the program cannot reach internal print systems, it cannot download drawings.

## Workflow
1. Select the initial CAD export Excel file.
2. Select the hydraulic schematic PDF.
3. Enter the output PDF file name.
4. Use the reorder window to:
   - Move assemblies up/down
   - Remove levels
   - Add more CAD export files
5. Continue.
6. The tool downloads required drawing PDFs.
7. The tool builds the final compiled PDF.
8. The tool exports the generated packet structure to an Excel file.
9. The tool warns if any drawings are missing.

## Step-by-Step

### Step 1 — Start the program
- Launch the EXE.
- A small startup delay is normal.

### Step 2 — Select initial CAD export
The selected Excel file should include columns:
- `Object`
- `Name`
- `Item Number`

The program locates columns by name.

### Step 3 — Select hydraulic schematic PDF
The schematic becomes:
- Level `1` in the structure
- The first drawing in the packet

Level 1 description format:
- `HYDRAULIC SCHEMATIC [FILE NAME]`

Example:
- File: `1660 Main Hyd Schematic.pdf`
- Label: `HYDRAULIC SCHEMATIC [1660 MAIN HYD SCHEMATIC]`

### Step 4 — Enter output PDF name
- Example: `CombinedDrawings.pdf`
- If `.pdf` is omitted, it is appended automatically.
- Output is saved in the same folder as the first CAD export.

### Step 5 — Use the reorder window
Available actions:
- **Move Up/Move Down**: reorder selected level among siblings.
- **Remove Level**: remove selected level and its full subtree.
- **Add CAD Export**: add another CAD-exported Excel file.
- **Expand All / Collapse All**: tree view controls.
- **Continue**: accept structure and begin creation.
- **Cancel**: stop the process.

Reordering behavior:
- Moving a parent moves all children.
- Removing a parent removes all children.
- Children remain attached to parents.

## Adding Additional CAD Exports
To combine multiple CAD exports:
1. Click **Add CAD Export**.
2. Select another CAD-exported file.
3. Repeat as needed.

Important behavior:
- Excel file names are **not** used in final TOC/bookmarks.
- Top-level assemblies become top-level entries.
- Multiple top-level assemblies become separate top-level numbers.

Example:
1. HYDRAULIC SCHEMATIC [1660 MAIN HYD SCHEMATIC]
2. POWER UNIT
3. TELESCOPE
4. AXLE JACKS

## Included/Excluded Rows
Included item numbers start with:
- `13`
- `FB`
- `HA`

The program also preserves required parent assembly rows for structure integrity.

Excluded item numbers:
- `HA0814`
- `HA0815`
- `HA0816`
- `HA1129`
- `HA0817`
- `984398`

## Downloading Drawings
After **Continue**, the tool downloads required PDFs into:
- `_downloaded_drawings`

Location:
- Next to the first CAD export selected.

Saved filename format:
- `<part_number>.pdf` (example: `137792.pdf`)

## Final Output
The compiled PDF includes:
- Table of contents
- Hierarchical bookmarks
- Page numbers in `# / #` format

The output filename is the one entered by the user.

The tool also exports a structure Excel file next to the compiled PDF:
- `<output_pdf_name>_Structure.xlsx`
- Uses manual-builder style columns:
  - `Level`
  - `Description`
  - `Part Number`

## Missing Drawings
If some drawings cannot be found/downloaded, the tool still builds a packet from available files and shows a warning list at the end.

## Duplicate Part Numbers
If the same part number appears multiple times:
- Structure may show it more than once.
- Download step fetches each unique PDF once.
- Final packet may include repeated drawings if repeated in structure.

This is intentional and follows structure order.

## Troubleshooting

### Cannot find drawings
Possible causes:
- Part number does not exist in print system.
- Drawing not in current location.
- Network/print server unavailable.

What to do:
- Verify part number.
- Verify company network access.
- Check print lookup website.

### Program does not start
Possible causes:
- EXE blocked by Windows.
- Network path issue.
- Antivirus/IT policy interference.

What to do:
- Run from local folder.
- Check file **Properties** for **Unblock**.
- Contact IT.

### Reorder window is empty
Possible causes:
- Export has no valid matching items.
- No rows start with `13`, `FB`, or `HA` after exclusions.

What to do:
- Verify expected rows exist.
- Verify column names.

### Top-level assembly missing
Parent assembly rows are preserved when needed, including blank item numbers. If still missing, export structure/indentation may be unusual.

### Tool starts slowly
Minor startup delay is normal in packaged EXE.

## Best Practices
- Use clean CAD exports.
- Verify schematic PDF before starting.
- Review reorder window before continuing.
- Use meaningful output file names.
- Check end-of-run warning list.

## Recommended Workflow
1. Export CAD structure to Excel.
2. Launch Drawing Packet Builder.
3. Select CAD export.
4. Select schematic PDF.
5. Add additional exports as needed.
6. Reorder structure.
7. Continue.
8. Review warning list (if any).
9. Use compiled PDF packet.
