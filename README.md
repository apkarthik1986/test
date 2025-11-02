<<<<<<< HEAD
# Replace Tags AutoLISP Tool

This workspace contains an AutoLISP utility to find tag text in an AutoCAD drawing
and replace the text immediately to the right of each tag with a replacement value
from a spreadsheet.

Files
- `replace-tags.lsp` — AutoLISP script. Load in AutoCAD and run the `REPLTAG` command.

Purpose
- Read pairs (tag, replacement) from Excel (.xlsx/.xls) or CSV files and for each matching
  tag found in modelspace, locate the nearest text/MTEXT/attribute to the right and
  replace that text with the replacement value.

Prerequisites
- Full AutoCAD (Visual LISP support). AutoCAD LT does not support VLisp/COM features.
- Excel file (.xlsx/.xls) with two columns: column 1 = tag; column 2 = replacement.
- OR CSV file with same format for systems without Excel COM access.Input file formats

**Excel format (.xlsx/.xls):**
- Create a spreadsheet with two columns (no header required)
- Column A: Tag text to find (e.g., TAG001, TAG002)
- Column B: Replacement text (e.g., Replacement A, Replacement B)
- Save as .xlsx or .xls format

**CSV format (fallback):**
TAG001,Replacement A
TAG002,Replacement B

Notes:
- Empty rows are ignored in both formats.
- Leading/trailing spaces are trimmed.
- Excel COM requires appropriate system permissions.

How to use
1. Open your drawing in AutoCAD.
2. In the command line type `APLOAD` and load `replace-tags.lsp` (or drag it into the drawing).
3. Run the command `REPLTAG`.
4. When prompted, choose your Excel (.xlsx/.xls) or CSV file.
5. The script will auto-detect the file format and process accordingly.
6. Monitor the command line for progress and summary reports.

Behavior and assumptions
- The script searches only ModelSpace entities that expose a `TextString` property
	(TEXT, MTEXT, and attribute references are supported).
- For each tag match (exact string equality), the script looks for the nearest
	text object whose X coordinate is greater than the tag's X (i.e., to the right)
	and whose Y coordinate differs by at most 1.0 units (vertical tolerance). Adjust
	the tolerance in the LISP (`find-right-neighbor`) if needed.
- If multiple candidates exist to the right, the closest (smallest X distance)
	will be chosen.
- If no right neighbor is found for a tag occurrence, the tag is reported and left unchanged.

Limitations and edge cases
- If your tags and values are part of block attributes, the script will operate on
  attribute references in modelspace but will not modify attribute definitions inside blocks.
- Very dense drawings or unusual coordinate systems may require adjusting the vertical
  tolerance or pre-filtering the set of candidate text objects.
- Excel COM functionality requires full AutoCAD with Visual LISP and appropriate system
  permissions. If Excel COM fails, the script will report the error and suggest using CSV format.

Excel COM behavior
- The script automatically detects .xlsx/.xls files and uses Excel COM to read them directly.
- Excel application is opened invisibly, data is read from the active worksheet, then Excel is closed.
- If Excel COM is not available or fails, the script will display an error message and abort.
- For systems without Excel or COM access issues, export to CSV format as an alternative.
- The script reads all rows with data in the worksheet's used range (columns A and B).Testing and verification
- Before running on a large drawing, test on a copy of the DWG and a small CSV.
- Review replaced text and save a backup of your DWG before bulk changes.

Next steps / Improvements
- Add an option to preview replacements and confirm per-item.
- Add support for searching in PaperSpace/layouts or within nested block references.
- Add worksheet selection for Excel files with multiple sheets.
- Add progress indicator for large files with many rows.

Troubleshooting Excel COM issues
- If you get "Excel COM not available" errors, ensure:
  - Microsoft Excel is installed on the system
  - AutoCAD has appropriate permissions to launch COM applications
  - No Excel security policies are blocking COM automation
- Alternative: Export your Excel data to CSV format which always works.

---
Generated: updated to include the Replace Tags tool and usage instructions.
=======
# Replace Tags AutoLISP Tool

This workspace contains an AutoLISP utility to find tag text in an AutoCAD drawing
and replace the text immediately to the right of each tag with a replacement value
from a spreadsheet.

Files
- `replace-tags.lsp` — AutoLISP script. Load in AutoCAD and run the `REPLTAG` command.

Purpose
- Read pairs (tag, replacement) from Excel (.xlsx/.xls) or CSV files and for each matching
  tag found in modelspace, locate the nearest text/MTEXT/attribute to the right and
  replace that text with the replacement value.

Prerequisites
- Full AutoCAD (Visual LISP support). AutoCAD LT does not support VLisp/COM features.
- Excel file (.xlsx/.xls) with two columns: column 1 = tag; column 2 = replacement.
- OR CSV file with same format for systems without Excel COM access.Input file formats

**Excel format (.xlsx/.xls):**
- Create a spreadsheet with two columns (no header required)
- Column A: Tag text to find (e.g., TAG001, TAG002)
- Column B: Replacement text (e.g., Replacement A, Replacement B)
- Save as .xlsx or .xls format

**CSV format (fallback):**
TAG001,Replacement A
TAG002,Replacement B

Notes:
- Empty rows are ignored in both formats.
- Leading/trailing spaces are trimmed.
- Excel COM requires appropriate system permissions.

How to use
1. Open your drawing in AutoCAD.
2. In the command line type `APLOAD` and load `replace-tags.lsp` (or drag it into the drawing).
3. Run the command `REPLTAG`.
4. When prompted, choose your Excel (.xlsx/.xls) or CSV file.
5. The script will auto-detect the file format and process accordingly.
6. Monitor the command line for progress and summary reports.

Behavior and assumptions
- The script searches only ModelSpace entities that expose a `TextString` property
	(TEXT, MTEXT, and attribute references are supported).
- For each tag match (exact string equality), the script looks for the nearest
	text object whose X coordinate is greater than the tag's X (i.e., to the right)
	and whose Y coordinate differs by at most 1.0 units (vertical tolerance). Adjust
	the tolerance in the LISP (`find-right-neighbor`) if needed.
- If multiple candidates exist to the right, the closest (smallest X distance)
	will be chosen.
- If no right neighbor is found for a tag occurrence, the tag is reported and left unchanged.

Limitations and edge cases
- If your tags and values are part of block attributes, the script will operate on
  attribute references in modelspace but will not modify attribute definitions inside blocks.
- Very dense drawings or unusual coordinate systems may require adjusting the vertical
  tolerance or pre-filtering the set of candidate text objects.
- Excel COM functionality requires full AutoCAD with Visual LISP and appropriate system
  permissions. If Excel COM fails, the script will report the error and suggest using CSV format.

Excel COM behavior
- The script automatically detects .xlsx/.xls files and uses Excel COM to read them directly.
- Excel application is opened invisibly, data is read from the active worksheet, then Excel is closed.
- If Excel COM is not available or fails, the script will display an error message and abort.
- For systems without Excel or COM access issues, export to CSV format as an alternative.
- The script reads all rows with data in the worksheet's used range (columns A and B).Testing and verification
- Before running on a large drawing, test on a copy of the DWG and a small CSV.
- Review replaced text and save a backup of your DWG before bulk changes.

Next steps / Improvements
- Add an option to preview replacements and confirm per-item.
- Add support for searching in PaperSpace/layouts or within nested block references.
- Add worksheet selection for Excel files with multiple sheets.
- Add progress indicator for large files with many rows.

Troubleshooting Excel COM issues
- If you get "Excel COM not available" errors, ensure:
  - Microsoft Excel is installed on the system
  - AutoCAD has appropriate permissions to launch COM applications
  - No Excel security policies are blocking COM automation
- Alternative: Export your Excel data to CSV format which always works.

---
Generated: updated to include the Replace Tags tool and usage instructions.

>>>>>>> 7923db6d0ab75b2d306cf7d3a9ae803750203dfd