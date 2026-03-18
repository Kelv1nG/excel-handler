# Testing

## Structure

```
tests/
├── test_table_reader.py
├── test_cell_reader.py
├── test_template_reader.py
├── test_template_writer.py
└── fixtures/
    ├── simple_table.xlsx        # Single table, clean headers, no merged cells
    ├── no_headers.xlsx          # Table without a header row
    ├── merged_cells.xlsx        # Hierarchical data with merged cells
    ├── multiple_sheets.xlsx     # Same columns appearing in multiple sheets
    ├── multiple_tables.xlsx     # Two tables on the same sheet
    ├── with_summary.xlsx        # Table with a total/summary row at the bottom
    ├── empty_table.xlsx         # Table with headers but no data rows
    └── template.xlsx            # Template with {{variable}} marked cells
```

---

## Fixture Guidelines

Build fixtures with openpyxl in a `conftest.py` or a `create_fixtures.py` script — don't commit hand-crafted `.xlsx` files that are hard to understand or maintain.

Each fixture should test exactly one scenario. Prefer many small files over one large file that tests everything.

---

## ExcelTableReader Test Cases

### Happy path

```python
def test_extract_by_column_names(simple_table_path):
    with ExcelTableReader(simple_table_path) as reader:
        df = reader.extract_table_by_column_names(['Name', 'Amount'])
    assert list(df.columns) == ['Name', 'Amount']
    assert len(df) > 0
```

```python
def test_extract_by_range(simple_table_path):
    with ExcelTableReader(simple_table_path) as reader:
        df = reader.extract_table_by_range('A1:C5', sheet='Sheet1')
    assert df.shape == (4, 3)  # 4 data rows, 3 columns
```

```python
def test_extract_from_cell(simple_table_path):
    with ExcelTableReader(simple_table_path) as reader:
        df = reader.extract_table_from_cell('A1', sheet='Sheet1')
    assert len(df.columns) > 0
```

```python
def test_no_headers(no_headers_path):
    with ExcelTableReader(no_headers_path) as reader:
        df = reader.extract_table_by_range(
            'A1:C5', sheet='Sheet1',
            has_headers=False,
            column_names=['ID', 'Name', 'Value']
        )
    assert df.columns == ['ID', 'Name', 'Value']
```

```python
def test_merged_cells_filled(merged_cells_path):
    with ExcelTableReader(merged_cells_path) as reader:
        df = reader.extract_table_by_column_names(
            ['Region', 'Country', 'Sales'],
            unmerge_cells=True,
            fill_forward=True,
        )
    # Every row should have a Region value — no Nones from the merge
    assert df['Region'].null_count() == 0
```

### Error cases

```python
def test_table_not_found(simple_table_path):
    with ExcelTableReader(simple_table_path) as reader:
        with pytest.raises(TableNotFoundError):
            reader.extract_table_by_column_names(['NonExistent', 'Column'])
```

```python
def test_multiple_tables_raises(multiple_tables_path):
    with ExcelTableReader(multiple_tables_path) as reader:
        with pytest.raises(MultipleTablesFoundError) as exc_info:
            reader.extract_table_by_column_names(['Name', 'Amount'])
    assert len(exc_info.value.found_in) > 1
```

```python
def test_multiple_sheets_raises(multiple_sheets_path):
    with ExcelTableReader(multiple_sheets_path) as reader:
        with pytest.raises(MultipleTablesFoundError):
            reader.extract_table_by_column_names(['Name', 'Amount'])
```

```python
def test_sheet_not_found(simple_table_path):
    with ExcelTableReader(simple_table_path) as reader:
        with pytest.raises(ExcelTableReaderError):
            reader.extract_table_by_range('A1:C5', sheet='DoesNotExist')
```

```python
def test_context_manager_required():
    reader = ExcelTableReader('any.xlsx')
    with pytest.raises(ExcelTableReaderError):
        reader.extract_table_by_column_names(['Col'])
```

### Edge cases

```python
def test_empty_table_returns_empty_df(empty_table_path):
    with ExcelTableReader(empty_table_path) as reader:
        df = reader.extract_table_by_column_names(['Name', 'Amount'])
    assert len(df) == 0
    assert df.columns == ['Name', 'Amount']
```

```python
def test_fill_forward_false_leaves_nulls(merged_cells_path):
    with ExcelTableReader(merged_cells_path) as reader:
        df = reader.extract_table_by_column_names(
            ['Region', 'Country', 'Sales'],
            fill_forward=False,
        )
    # With fill_forward=False, merged cells produce Nones after the first row
    assert df['Region'].null_count() > 0
```

---

## ExcelCellReader Test Cases

```python
def test_get_single_cell(simple_table_path):
    with ExcelCellReader(simple_table_path) as reader:
        value = reader.get('Sheet1!A1')
    assert value is not None
```

```python
def test_get_many(simple_table_path):
    with ExcelCellReader(simple_table_path) as reader:
        values = reader.get_many(['Sheet1!A1', 'Sheet1!B1'])
    assert set(values.keys()) == {'Sheet1!A1', 'Sheet1!B1'}
```

```python
def test_get_without_sheet_uses_active(simple_table_path):
    with ExcelCellReader(simple_table_path) as reader:
        value = reader.get('A1')
    assert value is not None
```

```python
def test_get_empty_cell_returns_none(simple_table_path):
    with ExcelCellReader(simple_table_path) as reader:
        value = reader.get('Sheet1!Z99')  # Guaranteed empty
    assert value is None
```

```python
def test_get_invalid_sheet_raises(simple_table_path):
    with ExcelCellReader(simple_table_path) as reader:
        with pytest.raises(ExcelTableReaderError):
            reader.get('NoSuchSheet!A1')
```

---

## File-level Error Cases (apply to all readers)

```python
def test_file_not_found():
    with pytest.raises(ExcelFileNotFoundError):
        with ExcelTableReader('does_not_exist.xlsx') as reader:
            pass
```

```python
def test_corrupted_file(tmp_path):
    bad = tmp_path / 'bad.xlsx'
    bad.write_text('not an excel file')
    with pytest.raises(ExcelCorruptedError):
        with ExcelTableReader(str(bad)) as reader:
            pass
```

---

## What to Assert

- **Column names** — check `df.columns` exactly when the test depends on specific headers
- **Row count** — use `len(df)` or `df.shape[0]`; know your fixture row count
- **Null counts** — `df['col'].null_count()` for fill-forward and merge tests
- **Exception type** — always use `pytest.raises(SpecificError)`, not the base class
- **`.found_in` content** — check length or specific values when testing `MultipleTablesFoundError`

---

## What Not to Test

- Internal methods (`_*`) — test them through the public interface, not directly
- openpyxl behaviour — trust the library; test that our wrappers raise the right project exceptions
- Exact cell values in fixtures you control — use relative assertions (`> 0`, `is not None`) unless the value is the point of the test
