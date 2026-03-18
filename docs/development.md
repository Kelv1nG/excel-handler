# Development

## Conventions

### Naming

| Thing | Convention | Example |
|---|---|---|
| Classes | `PascalCase` | `ExcelTableReader` |
| Methods | `snake_case` | `extract_table_by_range` |
| Internal methods | `_snake_case` (single underscore) | `_detect_boundaries_in_sheet` |
| Exceptions | `PascalCase` + `Error` suffix | `TableNotFoundError` |
| Modules | `snake_case` | `table_reader.py` |
| Type aliases | `PascalCase` | `WorksheetMarkedCells` |

Public extraction methods are prefixed with the noun they return: `extract_table_*`, not `get_*` or `read_*`. Cell reading methods use `get` / `get_many` because they return scalars.

### File Placement

- New reader/extractor classes → `excel/`
- Shared file-level utilities (workbook loading, sheet checks) → `excel/utils.py`
- All exceptions → `excel/exceptions.py`
- Abstract base classes → `abstract.py` (root)

Do not put business logic in `utils.py`. It contains only generic Excel I/O helpers that any class might need.

---

## Adding a New Reader Class

1. Create `excel/<name>.py`
2. Accept `filepath: str` in `__init__`, store as `self.filepath`
3. Store the workbook as `self._wb: Workbook | None = None`
4. Implement `__enter__` / `__exit__` using `load_excel_workbook` from `excel.utils`
5. Add a `wb` property that raises `ExcelTableReaderError` if `_wb` is None
6. Add your public methods
7. Register any new exceptions in `excel/exceptions.py`

Skeleton:

```python
from types import TracebackType
from openpyxl.workbook.workbook import Workbook
from excel.exceptions import ExcelTableReaderError
from excel.utils import load_excel_workbook


class ExcelFooReader:
    def __init__(self, filepath: str):
        self.filepath = filepath
        self._wb: Workbook | None = None

    @property
    def wb(self) -> Workbook:
        if self._wb is None:
            raise ExcelTableReaderError(
                "Workbook not loaded. Use ExcelFooReader as context manager."
            )
        return self._wb

    def __enter__(self) -> "ExcelFooReader":
        self._wb = load_excel_workbook(self.filepath, data_only=True)
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: TracebackType | None,
    ) -> None:
        if self._wb:
            self._wb.close()
```

---

## Adding a New Exception

Add to `excel/exceptions.py`. Place it under the correct base:

- File-level failures (missing, permission, corrupted) → extend `ExcelError`
- Table/cell reader failures → extend `ExcelTableReaderError`
- Template failures → extend `TemplateReadError`

If the error carries data (like a list of locations), add `__init__` with typed attributes:

```python
class MyNewError(ExcelTableReaderError):
    def __init__(self, message: str, context: list[str]):
        super().__init__(message)
        self.context = context
```

---

## Type Hints

Use them on all methods and properties. Minimum expected coverage:

- Return type on every public method
- Parameter types on every public method
- `__init__` parameter types

type hints / strict type checking should be followed unless there are no stubs/proper types in the external library that is being used

Use built-in generics (`list[str]`, `dict[str, Any]`) — not `List`, `Dict` from `typing`. Python 3.12+ supports this natively.

---

## Error Handling

**Only validate at boundaries.** Don't add defensive checks inside `_*` internal methods — trust that the public methods have already validated inputs.

**Wrap openpyxl errors** in project exceptions. Never let raw `FileNotFoundError` or openpyxl exceptions escape the public surface. `load_excel_workbook` in `utils.py` handles this for file loading; do the same for any other openpyxl calls that can fail.

**Don't swallow exceptions.** Use `raise X from e` to preserve the original traceback.

---

## What Not to Do

- Don't add `exclude_summary` logic inline — it's a planned feature, keep the interface clean until it's fully designed
- Don't add transformation logic (groupby, pivots, type casting) to any reader — return raw values and let the caller use Polars
- Don't open the same workbook twice in one `with` block — pass the `Worksheet` object between internal methods instead
- Don't modify the source file — all mutations (unmerge, fill) happen on the in-memory workbook object only
