# API Reference

## ExcelTableReader

**Module:** `excel.table_reader`

Extracts rectangular tables from Excel files as Polars DataFrames. Must be used as a context manager.

```python
from excel.table_reader import ExcelTableReader
```

---

### `__init__(filepath: str)`

| Parameter | Type | Description |
|---|---|---|
| `filepath` | `str` | Path to the `.xlsx` file |

---

### `extract_table_by_column_names`

Search every sheet for a header row containing all required columns. Expects exactly one match across the entire workbook.

```python
def extract_table_by_column_names(
    required_columns: list[str],
    unmerge_cells: bool = True,
    fill_forward: bool = True,
) -> pl.DataFrame
```

| Parameter | Default | Description |
|---|---|---|
| `required_columns` | — | Column names that must all appear in the header row |
| `unmerge_cells` | `True` | Unmerge merged cells and fill the value into every cell in the range |
| `fill_forward` | `True` | Forward-fill `None` values column-by-column (useful for hierarchical data) |

**Raises:**
- `TableNotFoundError` — no sheet contains a row with all required columns
- `MultipleTablesFoundError` — columns found in more than one location; `.found_in` lists them

---

### `extract_table_by_column_names_from_sheet`

Same as above but restricted to one sheet.

```python
def extract_table_by_column_names_from_sheet(
    required_columns: list[str],
    sheet_name: str,
    unmerge_cells: bool = True,
    fill_forward: bool = True,
) -> pl.DataFrame
```

| Parameter | Default | Description |
|---|---|---|
| `required_columns` | — | Column names to find |
| `sheet_name` | — | Sheet to search |
| `unmerge_cells` | `True` | See above |
| `fill_forward` | `True` | See above |

**Raises:**
- `TableNotFoundError` — columns not found in this sheet
- `MultipleTablesFoundError` — columns found in more than one row within the sheet
- `ExcelTableReaderError` — sheet does not exist, or a required column appears more than once in the header row

---

### `extract_table_by_range`

Extract from an explicit Excel range string.

```python
def extract_table_by_range(
    range_str: str,
    sheet: str,
    has_headers: bool = True,
    column_names: list[str] | None = None,
    unmerge_cells: bool = True,
    fill_forward: bool = True,
) -> pl.DataFrame
```

| Parameter | Default | Description |
|---|---|---|
| `range_str` | — | Excel range, e.g. `"A5:C20"` |
| `sheet` | — | Sheet name (required) |
| `has_headers` | `True` | Whether the first row of the range is a header row |
| `column_names` | `None` | Manual column names when `has_headers=False`; auto-generates `col_0`, `col_1`... if omitted |
| `unmerge_cells` | `True` | See above |
| `fill_forward` | `True` | See above |

**Raises:**
- `ExcelTableReaderError` — sheet does not exist

---

### `extract_table_from_cell`

Extract from a known top-left corner cell, auto-detecting the right and bottom boundaries.

```python
def extract_table_from_cell(
    start_cell: str,
    sheet: str,
    has_headers: bool = True,
    column_names: list[str] | None = None,
    unmerge_cells: bool = True,
    fill_forward: bool = True,
    max_empty_rows: int = 2,
) -> pl.DataFrame
```

| Parameter | Default | Description |
|---|---|---|
| `start_cell` | — | Top-left cell of the table, e.g. `"A5"` |
| `sheet` | — | Sheet name (required) |
| `has_headers` | `True` | Whether the start row is a header row |
| `column_names` | `None` | Manual column names when `has_headers=False` |
| `unmerge_cells` | `True` | See above |
| `fill_forward` | `True` | See above |
| `max_empty_rows` | `2` | Stop expanding downward after this many consecutive empty rows |

**Raises:**
- `ExcelTableReaderError` — sheet does not exist

---

## ExcelCellReader

**Module:** `excel.cell_reader`

Reads individual cell values from an Excel file. Must be used as a context manager.

```python
from excel.cell_reader import ExcelCellReader
```

---

### `__init__(filepath: str)`

| Parameter | Type | Description |
|---|---|---|
| `filepath` | `str` | Path to the `.xlsx` file |

---

### `get`

Read a single cell value.

```python
def get(cell_ref: str) -> Any
```

| Parameter | Description |
|---|---|
| `cell_ref` | Either `"Sheet1!B5"` (explicit sheet) or `"B5"` (active sheet) |

**Returns:** The cell value — `str`, `int`, `float`, `datetime`, or `None` for an empty cell.

**Raises:**
- `ExcelTableReaderError` — sheet name in the reference does not exist, or workbook has no active sheet

---

### `get_many`

Read multiple cells at once.

```python
def get_many(cell_refs: list[str]) -> dict[str, Any]
```

| Parameter | Description |
|---|---|
| `cell_refs` | List of cell references in the same format as `get()` |

**Returns:** `dict` mapping each reference string to its value.

**Raises:** Same as `get()` for any invalid reference in the list.

---

## ExcelTemplateReader

**Module:** `excel.template_reader`

Parses an Excel template file and returns the locations of `{{variable}}` marked cells.

```python
from excel.template_reader import ExcelTemplateReader
```

### `read`

```python
def read(file: str | PathLike[str]) -> WorksheetMarkedCells
```

**Returns:** `dict[str, list[MarkedCell]]` — sheet name → list of marked cells on that sheet.

**Raises:**
- `TemplateReadError` — file not found or unreadable

---

### `MarkedCell`

| Attribute | Description |
|---|---|
| `name` | Variable name extracted from `{{name}}` |
| `cell_addr` | Cell address, e.g. `"B5"` |
| `metadata` | Raw string after the variable name (for loop/filter directives) |

```python
def parse_metadata(self) -> dict[str, Any]
```

Parses the metadata string into a structured dict.

---

## ExcelTemplateWriter

**Module:** `excel.template_writer`

Fills an Excel template with a variables dict and writes the result to a new file.

```python
from excel.template_writer import ExcelTemplateWriter
```

### Template Syntax

| Syntax | Behaviour |
|---|---|
| `{{variable}}` | Replaced with the scalar value from the variables dict |
| `{{table \| loop, }}` | Expands rows for each item in a list |

### `write`

```python
def write(vars: dict[str, Any], file: str | PathLike[str]) -> None
```

| Parameter | Description |
|---|---|
| `vars` | Variables dict — scalars, lists of dicts for loops |
| `file` | Output file path |

---

## Exceptions

**Module:** `excel.exceptions`

```python
from excel.exceptions import (
    ExcelError,
    ExcelFileNotFoundError,
    ExcelPermissionError,
    ExcelCorruptedError,
    ExcelTableReaderError,
    TableNotFoundError,
    MultipleTablesFoundError,
    TemplateReadError,
)
```

| Exception | Base | When raised |
|---|---|---|
| `ExcelError` | `Exception` | Base for all file-level errors |
| `ExcelFileNotFoundError` | `ExcelError` | File path does not exist |
| `ExcelPermissionError` | `ExcelError` | No read permission on the file |
| `ExcelCorruptedError` | `ExcelError` | File exists but is not a valid Excel file |
| `ExcelTableReaderError` | `ExcelError` | Base for all reader errors (sheet missing, not in context manager, etc.) |
| `TableNotFoundError` | `ExcelTableReaderError` | No table found matching the required columns |
| `MultipleTablesFoundError` | `ExcelTableReaderError` | Ambiguous — required columns matched more than one location |
| `TemplateReadError` | `Exception` | Template file could not be read or parsed |

### `MultipleTablesFoundError`

```python
MultipleTablesFoundError(message: str, found_in: list[str])
```

| Attribute | Description |
|---|---|
| `.found_in` | List of location strings where the columns were found, e.g. `["Sheet1 row 3", "Sheet2 row 7"]` |

---

## Shared Utilities

**Module:** `excel.utils`

```python
from excel.utils import load_excel_workbook, get_sheet_names, sheet_exists, validate_excel_file
```

### `load_excel_workbook`

```python
def load_excel_workbook(
    filepath: str | os.PathLike,
    read_only: bool = False,
    data_only: bool = False,
) -> Workbook
```

Wraps `openpyxl.load_workbook` with typed error handling. Raises project exceptions instead of raw openpyxl/OS errors.

### `get_sheet_names`

```python
def get_sheet_names(filepath: str | os.PathLike) -> list[str]
```

Returns sheet names without loading cell data. Opens and closes the workbook internally.

### `sheet_exists`

```python
def sheet_exists(filepath: str | os.PathLike, sheet_name: str) -> bool
```

### `validate_excel_file`

```python
def validate_excel_file(filepath: str | os.PathLike) -> bool
```

Returns `True` if the file can be opened as a workbook, `False` otherwise. Does not raise.
