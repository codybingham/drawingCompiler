Drawing Book Builder
Overview
Drawing Book Builder is a tool for assembling a single, organized PDF drawing book from individual drawing PDFs. The program uses a structured Excel file to define the order and hierarchy of the drawings, automatically generating a Table of Contents (TOC) and bookmarks for easy navigation. Page numbers are added to each page.

This tool streamlines the creation of drawing books for hydraulic systems and other assemblies, making it easy to share, review, and archive drawing packages.

Features
Reads a structured Excel file to determine drawing order and hierarchy
Merges individual PDF drawings into a single PDF
Automatically generates a Table of Contents (TOC) with page numbers
Adds PDF bookmarks for quick navigation
Numbers all pages in the final PDF
Simple interface for non-technical users
How to Use
Prepare Your Files

Ensure you have a structured Excel file (see format below).
Gather all individual drawing PDFs referenced in the Excel file.
Run the Program

Double-click the executable (DrawingBookBuilder.exe).
When prompted, select the folder containing your PDFs and Excel structure file.
Generate the Drawing Book

The program will merge the PDFs according to the Excel structure.
A combined PDF with TOC and bookmarks will be created in the output folder.
Review the Output

Open the final PDF to verify the order, TOC, bookmarks, and page numbers.
Excel File Format
The Excel file should contain at least the following columns:

Level	Description	Part Number
1	STANDARD HYDRAULICS,1252	FB9893
1.1	MIDDLE AXLE JACKS + STEER, 1252	HA2526
1.1.1	VLV ASY,JACKS,P.O. CHECK,1252	137691
1.1.2	VLV ASY, JACK FLOW DIV, 1252	137688
...	...	...
Level: Hierarchical number (e.g., 1, 1.1, 1.1.1)
Description: Drawing or assembly description
Part Number: Used to match the PDF file (e.g., 137691.pdf)
Understanding "Level" Numbering
The Level column defines the hierarchical structure of the drawing book, similar to an outline or table of contents. Each number represents the position and relationship of a drawing or assembly within the overall system.

Whole numbers (e.g., 1, 2, 3):

Represent top-level assemblies or main sections.

Decimal levels (e.g., 1.1, 1.2):

Indicate sub-assemblies or major components within the parent assembly.

Further decimals (e.g., 1.1.1, 1.1.2):

Represent components or drawings within those sub-assemblies.

Example:

Level	Description	Part Number
1	STANDARD HYDRAULICS,1252	FB9893
1.1	MIDDLE AXLE JACKS + STEER, 1252	HA2526
1.1.1	VLV ASY,JACKS,P.O. CHECK,1252	137691
1.1.2	VLV ASY, JACK FLOW DIV, 1252	137688
1 is the main assembly.
1.1 is a sub-assembly within 1.
1.1.1 and 1.1.2 are components within 1.1.
This structure allows the program to:

Organize the merged PDF in a logical, hierarchical order.
Generate a Table of Contents and bookmarks that reflect the assembly breakdown.
Make navigation intuitive for anyone reviewing the drawing book.

Notes & Troubleshooting
All PDF filenames must match the Part Numbers in the Excel file (e.g., 137691.pdf).
If a drawing is missing, the program will report it in the log.
For best results, ensure all PDFs are in the same folder and are not password-protected.
Example
Input:

1252DrawingStructure.xlsx (structure file)
PDFs: FB9893.pdf, HA2526.pdf, 137691.pdf, etc.
Output:

1252_DrawingBook.pdf (combined, organized PDF with TOC and bookmarks)