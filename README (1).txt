ALLOY BENCH - Report Builder (Desktop GUI App)
=================================================

SETUP (one time):
1. Python 3.8+ installed (python.org - tick "Add python.exe to PATH"
   during install on Windows).
2. Install dependencies:
   pip install pandas openpyxl reportlab

RUN:
   python report_builder.py

This creates an "output_templates" folder next to the script the first
time you save a template - keep report_builder.py in the same folder
every time you run it so your saved templates are still there.

WHAT'S NEW - WORKSPACES (multitasking):
- The window now opens with a top navy/cherry bar and one "Workspace 1"
  tab. Click "+ New Workspace (Ctrl+T)" top-right to open another -
  each workspace is completely independent: its own loaded file,
  columns, filters, sort, undo history and bulk-process state. Work on
  two different extraction jobs side by side without them interfering.
- "Close This Workspace" (top-left of each workspace) or Ctrl+W closes
  the current one. At least one workspace always stays open.
- Ctrl+Z / Ctrl+Y (undo/redo) apply to whichever workspace tab is
  currently active.
- Output Templates you save are shared across all workspaces (they
  live on disk), so you build the library once and use it from any tab.

BASIC WORKFLOW (single file, inside a workspace):
1. "Upload Excel File" -> pick your file.
2. A window shows the top ~20 rows and auto-guesses which row has your
   real column headers (skips logos/addresses/filter-summary junk above
   the table). Confirm or click a different row.
3. "Columns & Data Types" tab: uncheck columns you don't want, set a
   type per column (Text/Integer/Decimal/Date/Yes-No).
4. "Filters (by condition)" tab: column + operator + value + "Add
   Filter" to drop rows automatically. Add as many as needed (AND logic).
5. "Sort" tab: add one or more sort keys, or just click a column
   heading in the preview table to sort instantly (click again to
   reverse) - same as Excel.
6. In the preview: click a row (Ctrl/Shift+click for more) then
   "Delete Selected Rows" to remove exact rows you don't want, no
   condition needed.
7. Click "Apply & Preview" any time you change columns/filters/sort.
8. "Export to Excel" / "Export to PDF".

OUTPUT TEMPLATES (reuse a fixed output structure):
1. Load a source file as above (steps 1-2).
2. Go to "Output Templates" tab -> "Upload New Template Structure..."
   -> pick an Excel file whose column headers are the structure you
   want your output to have -> confirm its header row -> give it a
   name -> it's saved to disk for future sessions and every workspace.
3. Select that template in the list -> "Auto-Fill Selected Template
   From Loaded Data". It matches template columns to your loaded
   file's columns automatically (exact name match first, close-spelling
   match second) and shows a review screen - fix any wrong guesses,
   leave ones blank if there's no source data for them, then Apply.
4. The working table below becomes the template's structure filled
   with your data - now use Filters/Sort/row deletion/Apply & Preview/
   Export exactly as in the basic workflow.

BULK PROCESS (many source files at once, one saved template):
1. Save at least one Output Template first.
2. "Bulk Process" tab -> "Select Multiple Excel Files..." -> pick all
   the source files for this batch (same column layout as each other).
3. Pick the Output Template to map them into from the dropdown.
4. "1) Auto-Map Columns (using first file)" - guesses the mapping and
   shows the same review screen; this mapping is reused for every file.
5. Choose ONE of:
   - "2a) Combine All Into Working Data" - stacks every file's mapped
     rows into one table (optionally tagged with a 'Source File'
     column), loaded into the workspace for further filter/sort/
     delete/export.
   - "2b) Export Each File Separately" - applies your current Filters/
     Sort settings to each file individually, writing one Excel (and
     optionally PDF) per source file into a folder you choose.
6. The log box at the bottom shows per-file success/failure.

UNDO / REDO:
- Ctrl+Z / Ctrl+Y, or the buttons in each workspace's toolbar.
- Covered: uploading/re-loading a file (including a different header
  row), Apply & Preview, deleting rows, applying an Output Template,
  and Bulk Process > Combine All Into Working Data.
- Applies per-workspace and clears when the app closes (not saved
  to disk).

Everything runs locally on your machine - no data leaves your computer.

MULTI-COLUMN SORT FROM THE PREVIEW TABLE:
- Click a column heading to sort by it (click again to flip direction).
- Shift+click another heading to add it as a secondary sort key (Shift+
  click a third for a tertiary key, and so on). Each sorted heading
  shows an arrow and a small priority number when more than one is
  active, e.g. Territory (up arrow)1, Route (up arrow)2.
- "Clear Sort" button next to "Delete Selected Rows" resets it.
- This is in addition to the Sort tab's "Add Sort Key" list, which
  does the same multi-column sort but from a dropdown instead of
  clicking headings directly - use whichever is faster for you.

SCROLLING FIX:
- The "Columns & Data Types" list and the "Review Column Mapping"
  popup now respond to mouse wheel / two-finger trackpad scroll while
  your pointer is over them (previously only the scrollbar worked).

DESKTOP SHORTCUTS (no need to open a terminal each time):
- Windows: double-click Run_AlloyBench.bat (shows a console window
  briefly; pauses on error so you can read it) - or
  Run_AlloyBench_Silent.vbs for no console window at all, closer to a
  real installed app. To put either on your Desktop: right-click it >
  Send to > Desktop (create shortcut).
- macOS/Linux: Run_AlloyBench.command - first run, either right-click >
  Open, or in Terminal: chmod +x Run_AlloyBench.command
- Keep whichever launcher you use in the SAME FOLDER as
  report_builder.py (and the output_templates folder that appears next
  to it) - the launcher just runs "python report_builder.py" from
  wherever it sits.

WINDOW SIZE / EXPORT BUTTONS:
- The app now auto-sizes its window to fit YOUR screen on launch
  (previously it opened at a fixed 1350x950, which on smaller/laptop
  screens pushed the bottom of the window - including the Export
  buttons - off screen with no way to reach them). It also starts
  maximized where your OS supports it.
- The window is fully resizable and reflows correctly when you
  minimize/maximize/restore or drag-resize it - the data preview area
  grows or shrinks to fill available space; the toolbars stay a fixed
  height.
- Minimum window size is 850x550 so it stays usable even on small
  displays; a horizontal scrollbar appears under the preview table if
  a sheet has more columns than fit on screen.
- "Export to Excel" and "Export to PDF" now also appear in the TOP
  toolbar of each workspace (top-right, next to Upload/Undo/Redo) as
  well as at the bottom, so they're always visible no matter the
  window size.

SORTING THE TEMPLATE LIST (Output Templates tab):
- Click "Template Name" or "# Columns" at the top of the saved-
  templates list to sort it that way; click again to reverse. Useful
  once you've saved several output structures and want to find one
  quickly.

SORTING FILLED DATA WITHOUT LEAVING THE TEMPLATES TAB:
- The "Output Templates" tab now has its own "Sort the filled data"
  section below the Auto-Fill button: pick a column + Ascending/
  Descending, click "Add / Update Sort Column". Add more columns the
  same way for a multi-level sort (first one added is the primary
  key). "Clear Sort" resets it.
- This is the exact same sort as clicking column headings in the
  preview table (and Shift+click for multi-column) - just reachable
  without switching to the Sort tab. All three (Templates tab
  controls, preview-header clicks, Sort tab) share the same result.
