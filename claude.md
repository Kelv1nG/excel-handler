# Excel Data & Template Processing Module

A standalone Python module for extracting data from Excel files and processing Excel templates with variable placeholders.

## Documentation

| Doc | Contents |
|---|---|
| [docs/architecture.md](docs/architecture.md) | Module layout, component responsibilities, error hierarchy, design patterns |
| [docs/api-reference.md](docs/api-reference.md) | Every class, method, parameter, return type, and exception |
| [docs/development.md](docs/development.md) | Naming conventions, how to add new classes/exceptions, type hint and error handling rules |
| [docs/testing.md](docs/testing.md) | Fixture structure, full test case coverage, what to assert |

## Overview

This module provides four components:

1. **ExcelTableReader** — extract tables from Excel files as Polars DataFrames
2. **ExcelCellReader** — read individual cell values from known positions
3. **ExcelTemplateReader** — parse `{{variable}}` marked cells in a template
4. **ExcelTemplateWriter** — fill a template with data and write the output

## Core Components

> For full details on each component see [docs/api-reference.md](docs/api-reference.md).

### 1. ExcelTableReader

**Purpose:** Extract structured data from Excel files as Polars DataFrames

**Key Methods:**
- `extract_table_by_column_names(column_names, exact_columns=False)` - Search all sheets; `column_names` is a mandatory subset to locate the table, result includes all detected columns unless `exact_columns=True`
- `extract_table_by_column_names_from_sheet(column_names, sheet_name, exact_columns=False)` - Extract from a specific sheet
- `extract_table_by_range(range_str, sheet)` - Extract explicit range like `"A5:C20"`
- `extract_table_from_cell(start_cell, sheet)` - Auto-detect boundaries from start position

**Basic Usage:**
```python
from excel.table_reader import ExcelTableReader

with ExcelTableReader('sales.xlsx') as reader:
    df = reader.extract_table_by_column_names(['Company', 'Amount'])

with ExcelTableReader('data.xlsx') as reader:
    df = reader.extract_table_by_range('A5:C20', sheet='Sheet1')

with ExcelTableReader('report.xlsx') as reader:
    df = reader.extract_table_from_cell('A5', sheet='Data')
```

---

### 2. ExcelCellReader

**Purpose:** Read individual scalar cell values from known positions

**Key Methods:**
- `get(cell_ref)` - Single cell, supports `"Sheet1!B5"` or `"B5"`
- `get_many(cell_refs)` - Multiple cells, returns `dict[str, Any]`

**Basic Usage:**
```python
from excel.cell_reader import ExcelCellReader

with ExcelCellReader('config.xlsx') as reader:
    rate = reader.get('Config!B5')
    values = reader.get_many(['Sheet1!B5', 'Sheet2!C10'])
```

---

### 3. ExcelTemplateReader

**Purpose:** Scan an Excel template for `{{ variable }}` tags and return their locations grouped by worksheet.

**Tag syntax:**
- `{{ variable }}` — simple scalar tag
- `{{ variable | key=value, key2=value2 }}` — tag with metadata (Jinja-style)

Metadata values are auto-coerced: `True`/`False` → `bool`, digit strings → `int`, decimal strings → `float`, everything else → `str`.

**Returns:** `WorksheetMarkedCells` = `dict[sheet_name, list[MarkedCell]]`. Sheets with no tags are excluded.

**Key types:**
- `MarkedCell.name` — variable name
- `MarkedCell.cell_addr` — A1 address, e.g. `"B5"`
- `MarkedCell.sheet` — worksheet name
- `MarkedCell.raw` — full original tag string
- `MarkedCell.metadata` — raw string after `|`
- `MarkedCell.parse_metadata()` → `dict[str, Any]`

**Basic Usage:**
```python
from excel.template_reader import ExcelTemplateReader

reader = ExcelTemplateReader()
structure = reader.read('template.xlsx')
# {"Sheet1": [MarkedCell(name="revenue", cell_addr="B2", ...), ...]}

# Access metadata
for sheet, cells in structure.items():
    for cell in cells:
        print(cell.name, cell.cell_addr, cell.parse_metadata())
```

---

### 4. ExcelTemplateWriter

**Purpose:** Fill a template with data and write the output file

**Template Syntax:**
- `{{variable}}` - Simple scalar replacement
- `{{variable | loop()}}` - Expand rows for a list
- `{{variable | table(join=outer, on=ColName)}}` - Fill a multi-row table

**Basic Usage:**
```python
from excel.template_writer import ExcelTemplateWriter

writer = ExcelTemplateWriter('template.xlsx', variables)
writer.fill('output.xlsx')
```

**Tag examples in the Excel cell:**
```
{{ revenue }}                                      ← scalar
{{ month | loop() }}                               ← loop row
{{ sector_table | table(join=outer, on=Sector) }}  ← table fill
{{ prices | table(join=right) }}                   ← right join, on= inferred from header
```

---

## Key Patterns

See [docs/architecture.md](docs/architecture.md) for full design notes.

### Context manager (required for all readers)
```python
with ExcelTableReader('file.xlsx') as reader:
    df = reader.extract_table_by_column_names(['Col1', 'Col2'])
```

### Handle ambiguous tables
```python
from excel.exceptions import TableNotFoundError, MultipleTablesFoundError

try:
    df = reader.extract_table_by_column_names(['Product', 'Sales'])
except MultipleTablesFoundError as e:
    # e.found_in lists every matched location
    df = reader.extract_table_from_cell('A5', sheet='Sheet1')
```

### Merged / hierarchical data
```python
df = reader.extract_table_by_column_names_from_sheet(
    ['Region', 'Country', 'Sales'],
    sheet_name='Data',
    unmerge_cells=True,
    fill_forward=True,
)
```

---

## Error Hierarchy

```
ExcelError
├── ExcelFileNotFoundError
├── ExcelPermissionError
├── ExcelCorruptedError
└── ExcelTableReaderError
    ├── TableNotFoundError
    └── MultipleTablesFoundError  (.found_in: list[str])
    └── ColumnNamesMismatchError  (.expected: int, .got: int)

TemplateReadError
```

---

## Future Enhancements

- [ ] Summary row exclusion (`exclude_summary` option)
- [ ] Lookup methods (VLOOKUP-style)
- [ ] Named range support (`extract_by_named_range`)
- [ ] Excel Table object support (`extract_excel_table`)
- [ ] Advanced column filtering
- [ ] Template grouped output (one file per group)