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

Search every sheet for a header row containing all listed columns. By default returns the full detected range (including any extra columns). Set `exact_columns=True` to require the table to have no extra columns beyond those listed.

```python
def extract_table_by_column_names(
    column_names: list[str],
    exact_columns: bool = False,
    unmerge_cells: bool = True,
    fill_forward: bool = True,
) -> pl.DataFrame
```

| Parameter | Default | Description |
|---|---|---|
| `column_names` | — | Column names that must all appear in the header row (mandatory subset used to locate the table) |
| `exact_columns` | `False` | If `True`, raises `TableNotFoundError` when the detected table has extra columns beyond those in `column_names` |
| `unmerge_cells` | `True` | Unmerge merged cells and fill the value into every cell in the range |
| `fill_forward` | `True` | Forward-fill `None` values column-by-column (useful for hierarchical data) |

**Raises:**
- `TableNotFoundError` — no sheet contains a row with all listed columns, or `exact_columns=True` and no sheet has an exact column set match
- `MultipleTablesFoundError` — columns found in more than one location; `.found_in` lists them

---

### `extract_table_by_column_names_from_sheet`

Same as above but restricted to one sheet.

```python
def extract_table_by_column_names_from_sheet(
    column_names: list[str],
    sheet_name: str,
    exact_columns: bool = False,
    unmerge_cells: bool = True,
    fill_forward: bool = True,
) -> pl.DataFrame
```

| Parameter | Default | Description |
|---|---|---|
| `column_names` | — | Column names to find |
| `sheet_name` | — | Sheet to search |
| `exact_columns` | `False` | If `True`, raises `TableNotFoundError` when the table has extra columns |
| `unmerge_cells` | `True` | See above |
| `fill_forward` | `True` | See above |

**Raises:**
- `TableNotFoundError` — columns not found in this sheet, or `exact_columns=True` and the table has extra columns
- `MultipleTablesFoundError` — columns found in more than one row within the sheet
- `ExcelTableReaderError` — sheet does not exist, or a required column appears more than once in the header row

---

### `extract_table_by_range`

Extract from an Excel range string.  Two modes:

* **Exact** (`dynamic=False`, default) — fixed rectangle.  Both column and row boundaries come from `range_str`.
* **Dynamic** (`dynamic=True`) — column span fixed by `range_str`; row boundary auto-detected downward.  Pass a single-row range (e.g. `"A1:D1"`) to pin the header row and let data rows grow freely.

```python
def extract_table_by_range(
    range_str: str,
    sheet: str,
    has_headers: bool = True,
    column_names: list[str] | None = None,
    dynamic: bool = False,
    unmerge_cells: bool = True,
    fill_forward: bool = True,
    max_empty_rows: int = 2,
    stop_at: str | None = None,
    stop_before: str | None = None,
) -> pl.DataFrame
```

| Parameter | Default | Description |
|---|---|---|
| `range_str` | — | Excel range, e.g. `"A5:C20"` or `"A1:D1"` |
| `sheet` | — | Sheet name (required) |
| `has_headers` | `True` | Whether the first row of the range is a header row |
| `column_names` | `None` | Manual column names when `has_headers=False`; auto-generates `col_0`, `col_1`… if omitted |
| `dynamic` | `False` | If `True`, fix the column span from `range_str` but auto-detect the bottom row |
| `unmerge_cells` | `True` | See above |
| `fill_forward` | `True` | See above |
| `max_empty_rows` | `2` | *(dynamic only)* Stop after this many consecutive empty rows |
| `stop_at` | `None` | *(dynamic only)* Stop when a row contains this value, including that row |
| `stop_before` | `None` | *(dynamic only)* Stop when a row contains this value, excluding that row |

**Raises:**
- `ValueError` — both `stop_at` and `stop_before` are provided
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

Scans an Excel workbook for `{{ variable }}` tags and returns their locations grouped by worksheet. Tags are Jinja-style and support an optional metadata block after a `|` separator.

```python
from excel.template_reader import ExcelTemplateReader, MarkedCell, WorksheetMarkedCells
```

### Tag syntax

| Form | Example |
|---|---|
| Simple scalar | `{{ revenue }}` |
| Function-call form | `{{ sector_table \| table(join=outer, on=Sector) }}` |
| Flat key=value form | `{{ count \| skip=2, flag=True }}` |

The **function-call form** is the standard syntax for typed tags (tables, loops). The type name goes before the parentheses; options go inside as `key=value` pairs. The **flat key=value form** is for simple scalar metadata with no type.

Values are auto-coerced in both forms: `True`/`False` → `bool`, digit strings → `int`, decimal strings → `float`, everything else → `str`.

### `read`

```python
def read(file: str | PathLike[str]) -> WorksheetMarkedCells
```

**Returns:** `WorksheetMarkedCells` — `dict[sheet_name, list[MarkedCell]]`. Only sheets containing at least one tag are included.

**Raises:**
- `TemplateReadError` — file not found, unreadable, or a tag has an empty variable name (e.g. `{{ }}`)

---

### `WorksheetMarkedCells`

```python
type WorksheetMarkedCells = dict[str, list[MarkedCell]]
```

Maps each sheet name to the list of `MarkedCell` objects found on that sheet.

---

### `MarkedCell`

```python
@dataclass
class MarkedCell:
    name: str        # variable name, e.g. "revenue"
    sheet: str       # worksheet name
    cell_addr: str   # A1-notation address, e.g. "B5"
    raw: str         # full original tag, e.g. "{{ revenue | loop }}"
    metadata: str    # raw string after the | (empty if none)
```

| Attribute | Example value |
|---|---|
| `name` | `"revenue"` |
| `sheet` | `"Summary"` |
| `cell_addr` | `"B5"` |
| `raw` | `"{{ revenue \| orientation=horizontal }}"` |
| `metadata` | `"orientation=horizontal"` |

#### `parse_metadata`

```python
def parse_metadata(self) -> dict[str, Any]
```

Parses `metadata` into a typed dict. Supports two forms:

**Function-call form** — type name becomes `result["type"]`, inner pairs are merged in:
```python
# {{ sector_table | table(join=outer, on=Sector, orientation=vertical) }}
MarkedCell(..., metadata="table(join=outer, on=Sector, orientation=vertical)").parse_metadata()
# {"type": "table", "join": "outer", "on": "Sector", "orientation": "vertical"}

# {{ items | loop() }}
MarkedCell(..., metadata="loop()").parse_metadata()
# {"type": "loop"}
```

**Flat key=value form** — for scalar metadata with no type:
```python
# {{ count | skip=2, flag=True }}
MarkedCell(..., metadata="skip=2, flag=True").parse_metadata()
# {"skip": 2, "flag": True}

MarkedCell(..., metadata="").parse_metadata()
# {}
```

**Raises:**
- `TemplateReadError` — a fragment is not in `key=value` form

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
| `{{ variable }}` | Replaced with the scalar value from the variables dict |
| `{{ variable \| loop() }}` | Expands the template row into N rows, one per list element |
| `{{ variable \| table() }}` | Fills a multi-row table using join semantics |

Tags use the **function-call form**: the type goes before `()`, options go inside.

#### Loop tags

All `loop()` tagged cells in the same row must reference list variables of equal length. The template row is duplicated N times.

```
{{ month | loop() }}    {{ amount | loop() }}
```

```python
vars = {
    "month":  TypedValue(["Jan", "Feb", "Mar"], "list"),
    "amount": TypedValue([100, 200, 300], "list"),
}
```

#### Table tags

One tag per table, placed in the first data row of the first fill column (one column right of the join key column). The row above must contain column headers matching the DataFrame.

```
# Template layout:
# Row N-1 (headers): | Sector   | Stocks                              | Benchmark |
# Row N   (data):    | Energy   | {{ sector_table | table(join=outer, on=Sector) }} |  |
# Row N+1 (data):    | Tech     |                                     |           |
```

**Join modes** (`join=` option, default `left`):

| Mode | Behaviour |
|---|---|
| `left` | Fill matched rows; leave unmatched template rows as-is |
| `inner` | Fill matched rows; clear data columns on unmatched template rows |
| `outer` | Fill matched rows; append unmatched DataFrame rows at the bottom |
| `right` | Overwrite all template rows top-down in DataFrame order; insert rows if DataFrame is longer; clear extras if shorter |

**`on=` option** — name of the DataFrame column to join against. Defaults to the template header above the join key column.

```
{{ sector_table | table(join=outer, on=Sector) }}
{{ prices       | table(join=right, on=ticker) }}
{{ data         | table() }}    ← left join, on= inferred from header
```

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
    ColumnNamesMismatchError,
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
| `ColumnNamesMismatchError` | `ExcelTableReaderError` | `column_names` length does not match the number of columns in the range |
| `TemplateReadError` | `Exception` | Template file could not be read or parsed |

### `MultipleTablesFoundError`

```python
MultipleTablesFoundError(message: str, found_in: list[str])
```

| Attribute | Description |
|---|---|
| `.found_in` | List of location strings where the columns were found, e.g. `["Sheet1 row 3", "Sheet2 row 7"]` |

### `ColumnNamesMismatchError`

```python
ColumnNamesMismatchError(message: str, expected: int, got: int)
```

| Attribute | Description |
|---|---|
| `.expected` | Number of columns in the range |
| `.got` | Number of names provided in `column_names` |

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
