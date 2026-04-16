# Drawing Downloader

**Version:** 1.0.0

## Overview
This tool is now a **standalone drawing downloader**.

It:
- Reads an existing structure Excel file.
- Downloads required drawing PDFs from the internal lookup service.

It **does not** compile or merge PDFs.

## Prerequisites
1. Python environment or packaged executable for this tool.
2. A structure Excel file with a `Part Number` column.
3. Network access to:
   - `prints.spudnik.local`
   - `api.spudnik.local`

## Workflow
1. Launch the tool.
2. Select a structure Excel file (`Part Number` column).
3. Select the folder that should contain downloaded PDFs.
4. Continue to start downloading.
5. Downloaded PDFs are saved in the selected folder.
6. If any drawings are missing, `missing_drawings.txt` is written in the selected folder.

## Notes
- Included item numbers start with `13`, `FB`, or `HA` (with predefined exclusions in code).
- Existing PDFs in the download folder are reused.
- Duplicate part numbers are downloaded once.
