# 📊 Simple-Grid (`PySheet`)

> **A fast, crisp, Excel-like desktop spreadsheet built with Python & PySide6—engineered to handle 20–30+ lakh rows smoothly without freezing.**

---

## ✨ Highlights

* **⚡ Built for Massive Data (20–30+ Lakh Rows):** Virtual grid rendering, sparse column storage, and O(1) header selection keep scrolling, selecting, and editing responsive even on multi-million-row sheets.
* **📁 Excel & CSV Read/Write:** Open and save `.xlsx` Excel workbooks as well as `.csv` and `.tsv` files, preserving formulas in Excel exports.
* **📑 Multi-Sheet Workbooks (`Sheet1`–`Sheet10`):** Add up to 10 sheets using the bottom `+` tab bar, switch tabs while preserving scroll position and filters, rename/delete sheets, and use cross-sheet formulas like `=Sheet2!A1`.
* **🧮 Excel-Style Formulas & Autocomplete:** Live function suggestions, argument hints, and point-and-click cell/range selection for `SUM`, `AVERAGE`, `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `SUMIF`, `COUNTIF`, `IF`, `IFERROR`, text utilities, and more.
* **🔍 AutoFilter & Text to Columns:** Select data with `Ctrl+A`, toggle column filter dropdowns with `Ctrl+Shift+L`, and split delimited text into columns using the **Text to Columns** wizard.
* **🖱️ Drag-Fill, Resizing & Row/Column Shifting:** Drag the bottom-right `+` fill handle to auto-fill series or formulas, drag header borders to resize columns/rows, and right-click row/column headers to shift positions with `Ctrl+Z` / `Ctrl+Y` Undo/Redo.
* **🔄 Built-In Auto-Updates:** Automatically checks your GitHub manifest for newer versions and updates `.exe` or `.py` builds with one click.

---

## 🚀 Quick Start

### Option 1: Run from Python

```bash
pip install PySide6 openpyxl
pip install polars   # Optional: extra-fast CSV loading

python SimpleGrid_v9.py [file.csv|file.xlsx]
