DRAWING DOWNLOADER
Version 1.0.0

OVERVIEW
This tool is a standalone drawing downloader.
It reads a structure Excel file and downloads required drawing PDFs
from internal lookup services.

This tool does NOT compile or merge PDFs.

REQUIREMENTS
- Tool executable (or Python runtime)
- A structure Excel file with Part Number column
- Network access to:
  - prints.spudnik.local
  - api.spudnik.local

WORKFLOW
1) Start the program
2) Select structure Excel file with Part Number column
3) Select folder to contain downloaded PDFs
4) Click Continue
5) Program downloads PDFs to selected folder
6) If drawings are missing, missing_drawings.txt is written in selected folder

NOTES
- Item numbers are included when they start with 13, FB, or HA (with exclusions).
- Existing downloaded PDFs are reused.
- Duplicate part numbers are downloaded once.
