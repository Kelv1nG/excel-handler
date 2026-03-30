# Architecture

## Overview

This module is a standalone Excel processing library with two responsibilities:

1. **Data extraction** — read structured data out of arbitrary Excel files
2. **Template processing** — fill Excel templates with variable data

These are kept deliberately separate. Extraction reads files you don't control (client reports, exports, etc.). Template processing writes files from templates you author.

---

## Module Layout

```
excel/
├── table_reader.py       # ExcelTableReader  — extracts tables as DataFrames
├── cell_reader.py        # ExcelCellReader   — reads individual cell values
├── template_reader.py    # ExcelTemplateReader — parses {{variable}} templates
├── template_writer.py    # ExcelTemplateWriter — fills templates with data
├── utils.py              # Shared workbook loading + sheet helpers
└── exceptions.py         # All exception classes

abstract.py               # TemplateReader / TemplateWriter ABCs
```

---

## Components

### ExcelTableReader (`excel/table_reader.py`)

Extracts rectangular tables from Excel files into Polars DataFrames. Three search strategies:

| Method | When to use |
|---|---|
| `extract_table_by_column_names` | Don't know which sheet or row the table is on |
| `extract_table_by_column_names_from_sheet` | Know the sheet, not the exact position |
| `extract_table_by_range` | Know the exact cell range |
| `extract_table_from_cell` | Know the top-left corner, want auto-boundary detection |

The search-by-column methods scan for a header row containing all required column names, then auto-detect the data boundaries from that position downward.

### ExcelCellReader (`excel/cell_reader.py`)

Lightweight scalar cell reader. No DataFrame overhead. Use this when you need specific values from known positions (config values, report headers, metadata fields).

Supports both `"Sheet1!B5"` and `"B5"` (active sheet) reference formats.

### ExcelTemplateReader (`excel/template_reader.py`)

Parses an Excel template file and returns the structure of `{{variable}}` marked cells, grouped by worksheet. Built on the `TemplateReader` ABC.

### ExcelTemplateWriter (`excel/template_writer.py`)

Accepts a template file and a variables dict, then writes a filled output file. Built on the `TemplateWriter` ABC.

### utils.py

Shared low-level helpers used by all components:

- `load_excel_workbook(filepath, read_only, data_only)` — wraps openpyxl with typed error handling
- `get_sheet_names(filepath)` — list sheets without loading the full workbook
- `sheet_exists(filepath, sheet_name)` — boolean check
- `validate_excel_file(filepath)` — returns True/False without raising

---

## Abstract Base Classes (`abstract.py`)

```python
class TemplateReader[T](ABC):
    def read(self, file: str | PathLike[str]) -> T: ...

class TemplateWriter(ABC):
    def write(self, vars: dict[str, Any], file: str | PathLike[str]) -> None: ...
```

These exist so the template system can be swapped to other formats (Word, PDF) without changing call sites. Any new template reader/writer must implement these.

---

## Error Hierarchy

```
ExcelError  (base for all file-level errors)
├── ExcelFileNotFoundError
├── ExcelPermissionError
└── ExcelCorruptedError

ExcelTableReaderError  (base for all table/cell reader errors)
├── TableNotFoundError
└── MultipleTablesFoundError
    └── .found_in: list[str]  — locations where table was found

TemplateReadError  (template parsing failures)
```

`ExcelTableReaderError` inherits from `ExcelError`, so you can catch the base if you want to handle all Excel failures in one place.

`MultipleTablesFoundError` carries a `.found_in` attribute listing every matched location — useful for falling back to a more specific extraction method.

---

## Design Patterns

### Context Manager (readers only)

All reader classes require a `with` block. The workbook is opened in `__enter__` and closed in `__exit__`, ensuring no file handles are leaked even on exception.

```python
with ExcelTableReader('file.xlsx') as reader:
    df = reader.extract_table_by_column_names(['Col1', 'Col2'])
```

Calling any extraction method outside a `with` block raises `ExcelTableReaderError` immediately.

### data_only=True

All readers open workbooks with `data_only=True`. This means formula results (last calculated values) are read instead of formula strings. Files that have never been opened in Excel will return `None` for formula cells.

### Boundary Detection

When no explicit range is given, `_detect_boundaries_in_sheet` expands right from the start column until an empty header cell, then expands downward until `max_empty_rows` (default 2) consecutive empty rows are found.

### Merged Cell Handling

When `unmerge_cells=True` (default), the reader unmerges all merged ranges in the sheet before extraction and fills the top-left value into every cell in the range. This is a destructive in-memory operation — the original file is never modified.

---

## Dependencies

| Package | Purpose |
|---|---|
| `openpyxl` | Read/write `.xlsx` files |
| `polars` | DataFrame representation of extracted tables |
| `pydantic` | Validation (used in template processing) |

Python 3.12+ required (uses `type` alias syntax and generic class syntax).
