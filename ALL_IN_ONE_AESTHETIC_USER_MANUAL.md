# Drawing Compiler Studio User Manual

User manual for `all_in_one_aesthetic.py`

## Table of Contents

1. [What the Program Does](#what-the-program-does)
2. [Who Should Use This Manual](#who-should-use-this-manual)
3. [Important Terms](#important-terms)
4. [Before You Start](#before-you-start)
5. [Starting the Program](#starting-the-program)
6. [Getting Around the Application](#getting-around-the-application)
7. [Input File Requirements](#input-file-requirements)
8. [Workflow 1: Automated Packet Builder](#workflow-1-automated-packet-builder)
9. [Workflow 2: Structure Editor](#workflow-2-structure-editor)
10. [Workflow 3: Drawing Downloader](#workflow-3-drawing-downloader)
11. [Workflow 4: Settings](#workflow-4-settings)
12. [Understanding Output Files](#understanding-output-files)
13. [Recommended Workflows](#recommended-workflows)
14. [Troubleshooting](#troubleshooting)
15. [Best Practices](#best-practices)
16. [Quick Reference](#quick-reference)

---

## What the Program Does

**Drawing Compiler Studio** is a desktop application for preparing drawing packets from structured drawing data.

It can:

- Read a **CAD export** spreadsheet.
- Convert CAD export rows into a structured drawing list.
- Read an existing **structure workbook**.
- Download drawing PDFs from the internal print lookup service.
- Use drawing URLs already present in a workbook.
- Build a final PDF packet with:
  - A table of contents.
  - Drawing PDFs in structure order.
  - Optional hydraulic schematic at the front.
  - Page numbers.
  - PDF bookmarks/outlines.
  - A final index.
- Let you manually edit, combine, reorder, and renumber structure files.
- Save preferences such as theme, default browse folder, and output naming patterns.

The program is meant to combine several drawing preparation tasks into one application so users do not need to run separate scripts for downloading, restructuring, and compiling packets.

---

## Who Should Use This Manual

This manual assumes you have **not had one-on-one training** in the program.

You can use it if you need to:

- Build a complete drawing packet from a CAD export.
- Build a packet from an already prepared structure workbook.
- Download drawings without building a packet yet.
- Fix the order or hierarchy of a structure workbook.
- Combine multiple CAD exports or structure files into one organized structure.
- Understand what each button and field does.

---

## Important Terms

### CAD Export

A spreadsheet exported from CAD or another source system. The program expects columns similar to:

- `Object`
- `Name`
- `Item Number`

The exact capitalization and spacing do not need to be perfect. For example, `ItemNumber`, `Item No`, and `Item` are also recognized for item numbers.

### Structure Workbook

An Excel workbook that describes the hierarchy and order of the drawing packet.

A structure workbook must contain these columns:

| Column | Meaning |
| --- | --- |
| `Level` | The hierarchy/order number, such as `1`, `1.1`, or `2.3.1`. |
| `Description` | The display name for the row. |
| `Part Number` | The drawing or item number used to find a PDF. |

The program also recognizes several alternate column names in some workflows, such as `Name`, `Item Number`, `Part`, or `Item`.

### Level

A dotted number that controls hierarchy and order.

Examples:

| Level | Meaning |
| --- | --- |
| `1` | First top-level item. |
| `1.1` | First child under item `1`. |
| `1.2` | Second child under item `1`. |
| `2` | Second top-level item. |

### Drawing PDF

A PDF file containing a drawing. In local folders, the program looks first for an exact filename such as:

```text
PARTNUMBER.pdf
```

If an exact file is not found, it searches the folder for any PDF filename that contains the part number.

### Hydraulic Schematic

An optional PDF that can be placed at the beginning of the final drawing packet. The program labels this entry as `HYDRAULIC SCHEMATIC` in the table of contents.

### Internal Print Lookup Service

The program can contact the internal service at:

```text
http://prints.spudnik.local/api/prints/format-paths
```

This service is used to find drawing PDF paths for part numbers. You must be on the appropriate company/internal network for this lookup to work.

---

## Before You Start

### Required Python Packages

The program is a Python desktop application. It uses:

- `pandas`
- `requests`
- `urllib3`
- `pypdf`
- `reportlab`
- `openpyxl` for Excel reading/writing through pandas
- `tkinter` for the desktop interface

`tkinter` is included with many Python installations, but some Linux distributions require it to be installed separately.

### Recommended Preparation

Before running a workflow, gather the files you will need:

- CAD export spreadsheet, if starting from CAD data.
- Structure workbook, if one already exists.
- Hydraulic schematic PDF, if the final packet should include one.
- A folder where downloaded drawings should be stored.
- A location and filename for the final output PDF or saved structure workbook.

### Network Requirements

The automated downloader and Drawing Downloader use the internal print lookup service. If you are not connected to the correct network, part-number lookup may fail.

Direct HTTP/HTTPS links in a workbook can still be attempted if your computer can reach those URLs.

---

## Starting the Program

From the repository root, run:

```bash
python all_in_one_aesthetic.py
```

A desktop window titled **Drawing Compiler Studio** opens.

The main window contains:

- A left sidebar with workflow buttons.
- A dashboard in the main area.
- A **Back** button in the lower-left sidebar.
- Workflow pages that appear in the main area.

If the program does not open, see [Troubleshooting](#troubleshooting).

---

## Getting Around the Application

### Dashboard

The dashboard is the first screen. It shows cards for the main workflows.

Use **Open →** on a card to enter that workflow.

### Sidebar

The sidebar lists the same workflows:

- **Dashboard**
- **Automated Packet**
- **Structure Editor**
- **Drawing Downloader**
- **Settings**

Click a sidebar item to switch workflows.

### Back Button

The **Back** button returns to the previous workflow page you visited. It is disabled when there is no previous page.

### Browse Buttons

Most workflows have buttons such as:

- **Browse…**
- **Browse folder…**
- **Save as…**

Use these to select files and folders instead of typing full paths manually.

### Required Fields

Run buttons stay disabled until required fields have values. If a run button is still disabled, check that all required fields on the page are filled in.

### Progress Windows

Long operations show a progress window. The progress window displays:

- Current phase or action.
- Progress percentage.
- Status text such as the part number currently being downloaded or queued.

Do not close the main program while a progress operation is running.

---

## Input File Requirements

### Structure Workbook Requirements

A structure workbook should be an Excel file such as `.xlsx`, `.xlsm`, or `.xls`.

It must contain:

```text
Level
Description
Part Number
```

Example:

| Level | Description | Part Number |
| --- | --- | --- |
| 1 | Main Assembly | 130001 |
| 1.1 | Sub Assembly | FB1234 |
| 1.1.1 | Detail Drawing | HA5678 |
| 2 | Second Assembly | 130002 |

Rows with blank `Level` values are ignored.

### CAD Export Requirements

A CAD export can be `.xlsx`, `.xlsm`, `.xls`, or `.csv`.

The program expects columns equivalent to:

```text
Object
Name
Item Number
```

The CAD conversion keeps rows with valid item numbers and required parent rows.

Valid item numbers must begin with one of these prefixes:

- `13`
- `FB`
- `HA`

Some known excluded item numbers are skipped by the program:

- `HA0814`
- `HA0815`
- `HA0816`
- `HA1129`
- `HA0817`
- `984398`

Rows named `SECTIONS` or `CONSTRAINTS` are treated as non-part rows and are not preserved unless needed by valid child rows.

### URL Columns for Downloading

The Drawing Downloader can also use URL columns. Supported column names include:

- `File URL`
- `Url`
- `PDF URL`
- `Link`
- `Path`

Only URLs beginning with `http://` or `https://` are downloaded directly.

### Multiple Input Files

Some workflows allow multiple files. When multiple files are selected, their paths are shown in one field separated by semicolons.

Do not choose the same file twice. The program blocks duplicate inputs in workflows that combine files.

---

## Workflow 1: Automated Packet Builder

Use **Automated Packet Builder** when you want the program to do the full process:

1. Read a CAD export or structure workbook.
2. Convert CAD data to a structure workbook if needed.
3. Download referenced drawing PDFs.
4. Build the final packet PDF.

### When to Use This Workflow

Use this workflow when:

- You have a CAD export and want a complete drawing packet.
- You already have a structure workbook and want drawings downloaded and compiled.
- You want one start-to-finish process instead of downloading and compiling separately.

### Inputs

The page has four required fields.

#### CAD export(s) or structure workbook

Choose one of the following:

- One existing structure workbook.
- One or more CAD export files.

Important: Do **not** mix a structure workbook with CAD exports in the same automated run. The program accepts either:

- One structure workbook, or
- One or more CAD exports.

If you provide CAD exports, the program creates a structure workbook automatically.

#### Hydraulic schematic PDF

Choose the schematic PDF that should be placed at the beginning of the output packet.

This field is required in the Automated Packet Builder.

#### Download folder

Choose the folder where downloaded drawing PDFs should be saved.

If a required drawing already exists in the folder, the downloader skips that file instead of downloading it again.

#### Output PDF path

Choose where the finished drawing packet PDF should be saved.

The output filename cannot:

- Be blank.
- Contain invalid Windows filename characters: `< > : " / \ | ? *`
- End with a space or period.
- Use a reserved Windows device name such as `CON`, `PRN`, or `LPT1`.

### Step-by-Step Instructions

1. Open **Automated Packet** from the dashboard or sidebar.
2. In **CAD export(s) or structure workbook**, click **Browse…**.
3. Select either one structure workbook or one or more CAD exports.
4. In **Hydraulic schematic PDF**, click **Browse…** and select the schematic PDF.
5. In **Download folder**, click **Browse folder…** and select or create the folder for drawing PDFs.
6. In **Output PDF path**, click **Save as…** and choose the final PDF name.
7. Click **Run Automated Build**.
8. Wait for the progress window to finish.
9. Review the completion summary.

### What Happens During the Run

If the input is a CAD export:

1. The program reads the CAD export.
2. It identifies supported columns.
3. It filters for valid item numbers.
4. It preserves necessary parent rows.
5. It creates a structure workbook next to the first input file.

Then, for both CAD-derived and existing structure inputs:

1. The program reads part numbers.
2. It looks up drawing paths using the internal print lookup service.
3. It downloads found PDFs into the download folder.
4. It skips PDFs already present in the download folder.
5. It builds the final packet from the downloaded/local PDFs.
6. It adds a table of contents, bookmarks, page numbers, and index.

### Completion Summary

After a successful run, the program shows a summary with:

- Output PDF path.
- Source type, such as `cad_export` or `structure`.
- Structure file path.
- Number of parts included.
- Number of skipped downloads.
- Number of drawings missing from the packet.
- Number of download failures.
- Number of part numbers not found by lookup.
- A shortened list of missing items.

### Common Automated Packet Problems

#### “File is neither a supported structure workbook nor a CAD export.”

The input file does not have the required columns for either supported file type.

Check that the file has either:

- `Level`, `Description`, and `Part Number`, or
- `Object`, `Name`, and `Item Number`.

#### “Use one pre-created structure workbook, or one or more CAD export files; do not mix them.”

You selected a structure workbook and at least one CAD export together.

Run them separately, or convert/combine CAD exports first and then use the resulting structure workbook.

#### Downloads fail or many parts are “not found”

Possible causes:

- You are not connected to the internal network.
- The internal print lookup service is unavailable.
- The part numbers are not recognized by the service.
- The CAD export contains unexpected item numbers.

---

## Workflow 2: Structure Editor

Use **Structure Editor** to edit the hierarchy and order of structure data before saving it as a clean workbook.

### When to Use This Workflow

Use this workflow when:

- The order of a structure workbook is wrong.
- You need to combine multiple structure files.
- You need to combine multiple CAD exports into one edited structure.
- You need to add, remove, or rename rows.
- You need to move items under different parent assemblies.
- You want the program to renumber levels automatically when saving.

### Screen Layout

The Structure Editor has:

- A toolbar at the top.
- A tree table below the toolbar.
- A `DESCRIPTION` column.
- A `PART NUMBER` column.

The tree indentation shows parent/child relationships.

### Toolbar Buttons

Depending on window width, some buttons may move into the **More ▾** menu.

| Button | What it does |
| --- | --- |
| **Add file** | Add one or more structure workbooks or CAD exports to the editor. |
| **Save structure** | Save the current edited structure as an Excel workbook. |
| **↑** | Move the selected row up among its siblings. |
| **↓** | Move the selected row down among its siblings. |
| **Add item** | Add a new row manually. |
| **Delete** | Delete the selected row and all of its children. |
| **Make child** | Move the selected row under another parent. |
| **Promote** | Move the selected row up one level. |
| **Expand all** | Open all branches in the tree. |
| **Collapse all** | Close all branches in the tree. |
| **Undo** | Undo the last editor change. |
| **Redo** | Redo the last undone editor change. |
| **Clear editor** | Remove all rows from the editor. |

### Keyboard Shortcuts

| Shortcut | Action |
| --- | --- |
| `Delete` | Delete the selected row and its children. |
| `Enter` | Edit the selected row. |
| `Ctrl+Z` | Undo. |
| `Ctrl+Y` | Redo. |

### Adding Files

1. Open **Structure Editor**.
2. Click **Add file**.
3. Select one or more supported files.
4. The program identifies whether each file is a structure workbook or CAD export.
5. CAD exports are converted into structure data before being added.
6. The data appears in the tree.

If a selected file has already been loaded, the program warns you and skips that duplicate file.

### Adding a Manual Item

1. Click **Add item**.
2. Enter a description.
3. Enter a part number, if applicable.
4. Choose the parent location when prompted.
5. Click **OK**.

Description is required. Part number may be blank for parent/grouping rows.

### Editing an Item

1. Select a row.
2. Press `Enter`.
3. Change the description or part number.
4. Click **OK**.

### Moving Items Up or Down

1. Select a row.
2. Click **↑** or **↓**.

The item moves only within its current sibling group. It does not become a child or parent through the arrow buttons.

### Making an Item a Child of Another Item

1. Select the row you want to move.
2. Click **Make child**.
3. In the parent selection dialog, choose the new parent.
4. Click **OK**.

The program hides invalid parent choices that would create circular hierarchy, such as making an item a child of itself or of one of its own descendants.

### Promoting an Item

1. Select a row that is currently a child.
2. Click **Promote**.

The selected row moves up one level and is placed after its old parent.

### Drag-and-Drop Reordering

You can also move rows by dragging them in the tree.

General behavior:

- Drop directly on another row to make the dragged row a child of that row.
- Drop near the top edge of a row to place it before that row.
- Drop near the bottom edge of a row to place it after that row.
- Drop in an invalid location to cancel the move.

The status message explains what will happen before you release the mouse.

### Deleting Items

1. Select a row.
2. Click **Delete** or press `Delete`.

Warning: deleting a row also deletes all of its child rows. Use **Undo** if you delete something by mistake.

### Undo and Redo

The editor tracks changes such as adding, deleting, moving, editing, and clearing rows.

Use:

- **Undo** or `Ctrl+Z` to reverse the previous change.
- **Redo** or `Ctrl+Y` to reapply an undone change.

### Saving a Structure

1. Click **Save structure**.
2. Choose an output filename.
3. Save the file.

When saving, the program writes a new workbook with these columns:

```text
Level
Description
Part Number
```

Levels are regenerated in the current tree order. This means the saved workbook is automatically renumbered.

### Using Structure Editor Data in the Drawing Downloader

After editing a structure, you can go to **Drawing Downloader** and check **Use Structure Editor structure**. The downloader then uses the current in-memory editor data instead of requiring you to save a workbook first.

---

## Workflow 3: Drawing Downloader

Use **Drawing Downloader** when you want to download drawing PDFs but do not necessarily want to build a packet yet.

### When to Use This Workflow

Use this workflow when:

- You want to pre-download all drawings for a structure.
- You want to check which drawings can be found before building a packet.
- You want to download drawings from the current Structure Editor data.
- You have a CAD export and want drawings downloaded into a folder.

### Inputs

#### Use Structure Editor structure

Check this box if you want the downloader to use the structure currently loaded in Structure Editor.

When checked:

- The file field is disabled.
- You do not need to browse for a workbook.
- The program temporarily exports the current editor structure in the background.

If the editor is empty, the download will fail with a message that no Structure Editor data is loaded.

#### Structure workbook or CAD export

Use this field if you are not using the Structure Editor data.

You can select:

- A structure workbook.
- A CAD export.

If you select a CAD export, the program converts it to temporary structure data before downloading.

#### Download folder

Choose the folder where downloaded PDFs should be saved.

### Step-by-Step Instructions

1. Open **Drawing Downloader**.
2. Choose your source:
   - Check **Use Structure Editor structure**, or
   - Browse for a structure workbook or CAD export.
3. Choose the **Download folder**.
4. Click **Download Drawings**.
5. Wait for the progress window to finish.
6. Review the completion summary.

### What the Downloader Uses

The downloader reads:

- Part numbers from supported part/item columns.
- Direct URLs from supported URL/link columns.

For part numbers, it contacts the internal print lookup service. For direct URLs, it downloads the URL directly.

### Completion Summary

After the run, the program shows:

- Source used.
- Download folder.
- Number downloaded.
- Number skipped because the file already existed.
- Number of part numbers not found by lookup.
- Number failed.
- A shortened list of not-found and failed items.

---

## Workflow 4: Settings

Use **Settings** to control appearance and default naming behavior.

### Theme

The program supports:

- Dark theme.
- Light theme.

Choose the theme you prefer and save settings.

### Default Browse Folder

The default browse folder controls where file picker dialogs start.

If you often work in the same project folder, set it here.

### Output Naming Templates

Naming templates control suggested output names when you browse for inputs.

Supported placeholders:

| Placeholder | Meaning |
| --- | --- |
| `{base}` | The selected input filename without extension. |
| `{ext}` | The output extension without a dot. |
| `{extension}` | The output extension with a dot. |

Default templates include:

| Template Key | Default Pattern | Meaning |
| --- | --- | --- |
| Automated packet | `{base}_automated_packet` | Suggested final automated packet filename. |
| Structure export | `{base}_structure` | Suggested structure workbook generated from CAD data. |
| Structure Editor save workbook | `{base}_reordered` | Suggested filename for saved edited structures. |
| Drawing download folder | `{base}_drawings` | Suggested folder name for downloaded drawings. |

The settings page also includes a manual packet template in the saved preferences, although the current main navigation focuses on the automated packet, editor, downloader, and settings workflows.

### Saving Settings

1. Open **Settings**.
2. Select a theme.
3. Choose a default browse folder if desired.
4. Edit naming templates if desired.
5. Click **Save Settings**.

Settings are saved to this file in your user home folder:

```text
.drawing_compiler_studio.json
```

If saving this file fails, the application continues running, but settings may not persist.

---

## Understanding Output Files

### Final Packet PDF

The final packet PDF contains:

1. Table of contents.
2. Optional hydraulic schematic.
3. Drawing PDFs in structure order.
4. Index.

The program also adds:

- Page numbering at the bottom of each page.
- Clickable table-of-contents links.
- PDF outline/bookmark entries.

### Structure Workbook

A saved structure workbook contains:

```text
Level
Description
Part Number
```

Use this file as input for:

- Automated Packet Builder.
- Drawing Downloader.
- Future Structure Editor sessions.

### Download Folder

The download folder contains drawing PDFs downloaded from lookup results or direct URLs.

Part-number downloads are saved as:

```text
PARTNUMBER.pdf
```

Direct URL downloads use the filename from the URL when possible.

---

## Recommended Workflows

### Scenario A: Build a Packet from a CAD Export

Use this when you are starting from CAD data.

1. Open **Automated Packet**.
2. Select the CAD export file or files.
3. Select the hydraulic schematic PDF.
4. Select a download folder.
5. Select an output PDF path.
6. Run the automated build.
7. Review the missing/not-found lists.

### Scenario B: Edit CAD Export Structure Before Building

Use this when the CAD export order needs correction.

1. Open **Structure Editor**.
2. Click **Add file**.
3. Select the CAD export file or files.
4. Reorder, add, edit, or delete rows as needed.
5. Click **Save structure**.
6. Open **Automated Packet**.
7. Select the saved structure workbook.
8. Complete and run the automated packet build.

### Scenario C: Download Drawings First, Build Later

Use this when you want to verify drawing availability before building.

1. Open **Drawing Downloader**.
2. Select a structure workbook, CAD export, or current Structure Editor data.
3. Choose a download folder.
4. Click **Download Drawings**.
5. Review skipped, not-found, and failed items.
6. Use the downloaded folder in a packet workflow later.

### Scenario D: Combine Multiple CAD Exports

1. Open **Structure Editor**.
2. Click **Add file**.
3. Select multiple CAD export files.
4. Edit the combined tree.
5. Save the structure workbook.
6. Use the saved workbook for downloading or packet building.

---

## Troubleshooting

### The program window does not open

Possible causes:

- Python is not installed or not on your `PATH`.
- Required packages are missing.
- `tkinter` is not installed.
- You are running the command from the wrong folder.

Try running from the repository root:

```bash
python all_in_one_aesthetic.py
```

If dependencies are missing, install the needed packages in your Python environment.

### A run button is disabled

A required field is blank.

Check all fields on the page. Automated Packet Builder requires all four fields.

### “Missing fields” warning appears

The workflow needs one or more fields filled in before it can run.

Complete all required fields and try again.

### “Invalid filename” appears

Your output filename is not allowed.

Avoid:

- Blank filenames.
- Characters such as `< > : " / \ | ? *`.
- Filenames ending in a space or period.
- Reserved Windows names such as `CON`, `PRN`, `AUX`, `NUL`, `COM1`, or `LPT1`.

### “Required columns not found” appears

The selected file does not contain the expected column names.

For a structure workbook, verify:

```text
Level
Description
Part Number
```

For a CAD export, verify:

```text
Object
Name
Item Number
```

### “No matching rows found after filters” appears

The CAD export was read, but no rows matched the item-number rules.

Check whether item numbers begin with supported prefixes:

- `13`
- `FB`
- `HA`

Also check whether the item numbers are in the correct column.

### Drawings are missing from the final packet

Possible causes:

- The drawing was not downloaded.
- The part number does not match the PDF filename.
- The internal lookup service did not find the part.
- The PDF exists in a different folder.
- The structure workbook has a blank or incorrect part number.

Check the completion summary and the download folder.

### Downloads are skipped

This usually means the PDF already exists in the download folder. Skipped downloads are not necessarily a problem.

### A direct URL does not download

Possible causes:

- The URL is not reachable.
- The URL requires credentials.
- The URL does not start with `http://` or `https://`.
- The server is down.

### Structure Editor buttons are disabled

Some buttons require a selected row. Select a row in the tree first.

Other buttons require loaded data. Use **Add file** or **Add item** first.

### I cannot see all toolbar buttons

If the window is narrow, extra Structure Editor actions move into **More ▾**. Click **More ▾** to access hidden actions.

### My saved structure has different Level numbers

This is expected. The Structure Editor renumbers levels based on the current tree order when saving.

---

## Best Practices

### Use a Dedicated Folder Per Job

For each drawing packet job, create a folder with subfolders such as:

```text
JobName/
├── input/
├── downloads/
├── output/
└── structure/
```

This makes it easier to find files and troubleshoot missing drawings.

### Save an Edited Structure Before Building

If you spend time reordering a structure, save it before running packet builds or downloads.

### Review Missing and Not-Found Lists

After automated runs, always review:

- Missing in packet.
- Lookup not found.
- Download failures.

These lists tell you what may need manual correction.

### Keep Original Inputs

Do not overwrite original CAD exports or original structure files. Save edited structures with a new name, such as:

```text
OriginalName_reordered.xlsx
```

### Check Part Numbers Carefully

Most missing drawing problems come from part-number mismatches. Confirm that part numbers in the structure match drawing filenames or lookup records.

### Avoid Duplicate Inputs

Do not load the same file multiple times into a workflow. The program blocks duplicates in several places, but keeping inputs clean helps avoid confusion.

---

## Quick Reference

### Main Workflows

| Workflow | Use it for |
| --- | --- |
| Automated Packet | Full download-and-build process. |
| Structure Editor | Reorder, edit, combine, and save structure data. |
| Drawing Downloader | Download PDFs without building a packet. |
| Settings | Change theme, default folder, and naming templates. |

### Required Structure Columns

```text
Level
Description
Part Number
```

### Required CAD Export Columns

```text
Object
Name
Item Number
```

### Valid CAD Item Prefixes

```text
13
FB
HA
```

### Common Output Files

| Output | Produced by |
| --- | --- |
| Final packet PDF | Automated Packet Builder. |
| Structure workbook | Automated CAD conversion or Structure Editor save. |
| Downloaded drawing PDFs | Automated Packet Builder or Drawing Downloader. |

### Most Useful Shortcuts

| Shortcut | Action |
| --- | --- |
| `Enter` | Edit selected Structure Editor row. |
| `Delete` | Delete selected Structure Editor row. |
| `Ctrl+Z` | Undo Structure Editor change. |
| `Ctrl+Y` | Redo Structure Editor change. |

---

## Final Notes for New Users

If you are unsure which workflow to start with:

- Choose **Automated Packet** if you want a complete final PDF packet.
- Choose **Structure Editor** if the hierarchy/order needs to be fixed first.
- Choose **Drawing Downloader** if you only want to download drawings.
- Choose **Settings** if you want to change defaults before working.

The safest first-time approach is:

1. Use **Structure Editor** to inspect or prepare the structure.
2. Save the structure workbook.
3. Use **Drawing Downloader** to verify drawing availability.
4. Use **Automated Packet** to build the final PDF.
