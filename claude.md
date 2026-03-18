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
- `extract_table_by_column_names(required_columns)` - Search all sheets, expect exactly one match
- `extract_table_by_column_names_from_sheet(required_columns, sheet_name)` - Extract from a specific sheet
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

**Purpose:** Parse `{{variable}}` marked cells in an Excel template

**Basic Usage:**
```python
from excel.template_reader import ExcelTemplateReader

reader = ExcelTemplateReader()
structure = reader.read('template.xlsx')
```

---

### 4. ExcelTemplateWriter

**Purpose:** Fill a template with data and write the output file

**Template Syntax:**
- `{{variable}}` - Simple scalar replacement
- `{{table | loop, }}` - Expand rows for a list

**Basic Usage:**
```python
from excel.template_writer import ExcelTemplateWriter

writer = ExcelTemplateWriter('template.xlsx', variables)
writer.fill('output.xlsx')
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