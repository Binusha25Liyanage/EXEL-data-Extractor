REPORT BUILDER - Desktop GUI App
=================================

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

BASIC WORKFLOW (single file):
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
   want your output to have (a blank one, or a past finished report) ->
   confirm its header row -> give it a name -> it's saved to disk for
   future sessions.
3. Select that template in the list -> "Auto-Fill Selected Template
   From Loaded Data". It matches template columns to your loaded
   file's columns automatically (exact name match first, close-spelling
   match second) and shows a review screen - fix any wrong guesses,
   leave ones blank if there's no source data for them, then Apply.
4. The working table below becomes the template's structure filled
   with your data - now use Filters/Sort/row deletion/Apply & Preview/
   Export exactly as in the basic workflow.
- Delete or preview a template's columns any time from that tab.

BULK PROCESS (many source files at once, one saved template):
1. Save at least one Output Template first (see above) - you can save
   as many different structures as you use regularly and pick a
   different one each run.
2. Go to "Bulk Process" tab -> "Select Multiple Excel Files..." -> pick
   all the source files for this batch (they should share the same
   column layout as each other, e.g. daily exports from the same
   system).
3. Pick the Output Template to map them into from the dropdown.
4. "1) Auto-Map Columns (using first file)" - guesses the mapping from
   the first file and shows the same review screen; this one mapping
   is then reused for every file in the batch.
5. Choose ONE of:
   - "2a) Combine All Into Working Data" - stacks every file's mapped
     rows into a single table (optionally tagging each row with a
     'Source File' column) and loads it into the same working table
     used above, so you can then filter/sort/delete rows/export it as
     one combined report.
   - "2b) Export Each File Separately" - keeps each source file as its
     own output, applying whatever Filters/Sort you've currently set
     up on the Filters/Sort tabs to each one, and writes one Excel (and
     optionally PDF) per source file into a folder you choose.
6. The log box at the bottom of the tab shows per-file success/failure.

Notes:
- Manual row deletions and preview-header sorting happen on top of the
  last "Apply & Preview" (or template/bulk-combine) result. Re-running
  Apply & Preview, or re-doing a template/bulk fill, recomputes from
  scratch and clears manual deletions - do those last.
- Everything runs locally on your machine - no data leaves your computer.
