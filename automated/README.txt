DRAWING PACKET BUILDER
Version 1.0.0

OVERVIEW
This tool creates a compiled drawing packet PDF from one or more CAD-exported Excel files.

The program does all of the following in one workflow:
- Reads a CAD-exported Excel file
- Preserves the assembly structure
- Lets you reorder or remove levels before generating the packet
- Lets you add additional CAD export files into the same packet
- Downloads required drawing PDFs from the internal print lookup system
- Lets you choose the hydraulic schematic PDF to use
- Builds a final compiled PDF with:
  - table of contents
  - bookmarks
  - page numbers

The final output is a single compiled PDF drawing packet.


WHAT YOU NEED BEFORE STARTING
Before you run the program, make sure you have:

1. The executable for this tool
2. At least one CAD-exported Excel file
3. The hydraulic schematic PDF you want to use
4. Network access to the internal print server:
   - prints.spudnik.local
   - api.spudnik.local

If the program cannot reach the internal print system, it will not be able to download drawings.


HOW THE PROGRAM WORKS
The program is designed to be used in this order:

1. Select the initial CAD export Excel file
2. Select the hydraulic schematic PDF
3. Enter the output PDF file name
4. Use the reorder window to:
   - move assemblies up or down
   - remove levels if needed
   - add more CAD export Excel files if needed
5. Continue
6. The tool downloads the needed drawing PDFs
7. The tool builds the final compiled PDF
8. The tool warns you if any drawings could not be found or downloaded


STEP-BY-STEP INSTRUCTIONS

STEP 1 - START THE PROGRAM
Double-click the EXE to launch it.

If startup feels a little slow, that is normal for the packaged EXE.


STEP 2 - SELECT THE INITIAL CAD EXPORT
You will first be prompted to select the initial CAD export Excel file.

This should be the Excel file exported from CAD that contains at least these columns:
- Object
- Name
- Item Number

The program intelligently finds these columns by name.

The CAD export is used to build the structure for the drawing packet.


STEP 3 - SELECT THE HYDRAULIC SCHEMATIC PDF
Next, you will be prompted to choose the hydraulic schematic PDF.

This schematic becomes:
- level 1 in the structure
- the first drawing in the compiled packet

The level 1 description will be written as:
HYDRAULIC SCHEMATIC [FILE NAME]

Example:
If the schematic file is:
1660 Main Hyd Schematic.pdf

then level 1 will be:
HYDRAULIC SCHEMATIC [1660 MAIN HYD SCHEMATIC]


STEP 4 - ENTER THE OUTPUT PDF NAME
You will then be asked for the output file name.

Example:
CombinedDrawings.pdf

If you do not include ".pdf", the program will add it automatically.

The final compiled packet will be saved in the same folder as the first CAD export you selected.


STEP 5 - USE THE REORDER WINDOW
After the files are loaded, the reorder window opens.

This is where you can organize the structure before the packet is built.

AVAILABLE BUTTONS

Move Up
Moves the selected level up among its siblings.

Move Down
Moves the selected level down among its siblings.

Remove Level
Removes the selected level and everything beneath it.

Important:
Removing a level removes its entire subtree.

Add CAD Export
Lets you add another CAD-exported Excel file into the same packet.

Expand All
Expands the full tree.

Collapse All
Collapses the tree.

Continue
Accepts the structure and moves on to downloading and packet creation.

Cancel
Stops the process.


HOW REORDERING WORKS
The tool preserves child structure automatically.

That means:
- moving a parent moves all of its children with it
- removing a parent removes all of its children with it
- children stay attached to their parent

This prevents the structure from becoming invalid.


ADDING ADDITIONAL CAD EXPORTS
If you want to build one packet from more than one CAD export:

1. Click "Add CAD Export"
2. Select another CAD-exported Excel file
3. Repeat as needed

The tool will add the top-level assemblies from those exports into the combined structure.

Important behavior:
- The program does NOT use the Excel file name in the final TOC or bookmarks
- The highest-level assembly from the CAD export becomes the top-level parent entry
- If an export contains more than one top-level assembly, each one gets its own new top-level number

Example:
1  HYDRAULIC SCHEMATIC [1660 MAIN HYD SCHEMATIC]
2  POWER UNIT
3  TELESCOPE
4  AXLE JACKS

This keeps the final packet based on real assemblies instead of file names.


WHAT ROWS ARE INCLUDED FROM THE CAD EXPORT
The program includes parts whose item numbers start with:
- 13
- FB
- HA

It also preserves parent assembly rows when needed so the structure remains correct.

This means:
- valid child parts are included
- their parent assembly headers are also preserved when needed
- junk rows such as Sections or Constraints are ignored


EXCLUDED ITEM NUMBERS
The following items are intentionally excluded:
- HA0814
- HA0815
- HA0816
- HA1129
- HA0817
- 984398


DOWNLOADING DRAWINGS
After you click Continue, the program looks up the required drawings using the internal print lookup system.

It downloads PDFs for the part numbers in the structure and stores them in a working folder named:

_downloaded_drawings

This folder is created next to the first CAD export file you selected.

Downloaded PDFs are saved using their part number as the file name.

Example:
137792.pdf


FINAL OUTPUT
The program creates a compiled PDF that includes:
- table of contents
- hierarchical bookmarks
- page numbers in the form:
  # / #

The final compiled PDF is saved using the output name you entered earlier.


WARNINGS FOR MISSING DRAWINGS
If the program cannot find or download one or more drawings, it will still build the packet using everything it was able to get.

At the end, it will show a warning listing the missing drawings.

This allows you to:
- still get a usable packet
- see exactly which drawings were missing


IMPORTANT NOTES ABOUT DUPLICATES
If the same part number appears in multiple places in the structure:
- the structure may show it more than once
- the download step only downloads each unique PDF once
- the compiled packet may include the same PDF more than once if it appears in more than one place in the structure

This is intentional because the packet follows the structure.


TROUBLESHOOTING

Problem: The program says it cannot find drawings
Possible causes:
- the part number does not exist in the internal print system
- the drawing is not in the current location
- the network or print server is unavailable

What to do:
- verify the part number
- verify you are on the company network
- check whether the drawing exists in the print lookup website

Problem: The program does not start
Possible causes:
- EXE blocked by Windows
- network path issue
- antivirus or IT policy interference

What to do:
- try running it from a local folder
- right-click the EXE, open Properties, and check for an Unblock option
- ask IT if executables from your folder are being blocked

Problem: The reorder window is empty
Possible causes:
- the selected CAD export did not contain valid matching items
- none of the rows started with 13, FB, or HA after exclusions

What to do:
- verify the export contains the expected rows
- verify the columns exist and are named correctly

Problem: The top-level assembly is missing
The program preserves parent assembly rows when needed, even if those rows have blank item numbers.
If this still happens, the CAD export may have an unusual structure or different indentation behavior.

Problem: The tool is slow to start
A small startup delay is normal, especially in the packaged EXE version.


BEST PRACTICES
- Use clean CAD exports
- Verify the schematic PDF before starting
- Review the reorder window before clicking Continue
- Keep the final packet name meaningful
- Check warnings at the end for missing drawings


RECOMMENDED WORKFLOW
1. Export the CAD structure to Excel
2. Launch Drawing Packet Builder
3. Select the CAD export
4. Select the schematic PDF
5. Add any additional CAD exports if needed
6. Reorder the structure as desired
7. Continue
8. Review warning list if any
9. Use the compiled PDF packet


VERSION
Drawing Packet Builder v1.0.0
