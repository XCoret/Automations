# Signature Comparer

Small Tkinter app that compares two paths (a file or a folder) and writes an Excel report of the differences.

## What it does

Give it a "previous" path and a "current" path. It scans both (recursively if they're folders), hashes every file, and matches them up by relative path. For each file it works out one of:

- **EQUAL** – same path, same hash
- **NOT EQUAL** – same path, different hash
- **MOVED** – different path, same hash
- **NEW** – only in the current path
- **DELETED** – only in the previous path

## Usage

```
pip install -r requirements.txt
python signature_comparer.py
```

Fill in **Previous path** and **Current path** (use Browse, or type a path directly — either can be a single file or a whole folder), pick which signatures to use (SHA1 is always on, CRC16 and MD5 are optional), and click **Compare**. Choose where to save, and you get an `.xlsx` file.

## About the Excel file

Everything lives on one sheet:

- Rows 1–7: a summary with the count and percentage of EQUAL / NOT EQUAL / NEW / DELETED / MOVED.
- Row 9 onward: the full detail table, one row per file, with the signature columns and paths.

The Status column and the summary counts are actual Excel formulas, not fixed text — so if you edit a cell by hand, the status and its color (green/red/yellow) recalculate on their own.

## Requirements

- Python 3
- `openpyxl` (see `requirements.txt`)
- Tkinter (ships with standard Python installs)
