from dataclasses import dataclass
from os import PathLike
from types import TracebackType
from typing import Any, Protocol, Literal, Self

import polars as pl

kind = Literal["single", "list", "table", "record"]


@dataclass
class TypedValue:
    """A variable value paired with its kind hint for template rendering.

    Attributes:
        value: The variable's value.
        kind: Rendering hint.  ``"single"`` – scalar replacement;
            ``"list"`` – loop-row expansion; ``"table"`` – DataFrame join.
    """

    value: Any
    kind: kind

    def __getitem__(self, col: str) -> Any:
        """Return the scalar value at *col* from a single-row record DataFrame.

        Raises:
            ValueError: If the DataFrame does not have exactly one row.
        """
        if self.value.height != 1:
            raise ValueError(
                f"Record variable has {self.value.height} rows; expected exactly 1"
            )
        return self.value[col][0]


class TemplateReader[T](Protocol):
    """Protocol for reading a template file and returning its tagged structure."""

    def read(self, file: str | PathLike[str] | bytes) -> T:
        """Read *file* and return the parsed template structure."""
        ...


class TemplateWriter(Protocol):
    """Protocol for filling a template with variable data and saving the result."""

    def write(self, vars: dict[str, TypedValue], file: str | PathLike[str]) -> None:
        """Fill the template with *vars* and save to *file*."""
        ...


class CellReader(Protocol):
    """Protocol for reading individual cell values from a workbook.

    Must be used as a context manager to ensure the workbook is properly
    opened and closed.
    """

    def __enter__(self) -> Self: ...

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: TracebackType | None,
    ) -> None: ...

    def get(self, cell_ref: str) -> Any:
        """Return the value of a single cell.

        Args:
            cell_ref: Cell address in ``"Sheet!A1"`` or ``"A1"`` form.
        """
        ...

    def get_many(self, cell_refs: list[str]) -> dict[str, Any]:
        """Return values for multiple cells.

        Args:
            cell_refs: List of cell addresses in ``"Sheet!A1"`` or ``"A1"`` form.

        Returns:
            Mapping of each cell reference to its value.
        """
        ...


class TableReader(Protocol):
    """Protocol for extracting tabular data from a workbook as Polars DataFrames.

    Must be used as a context manager to ensure the workbook is properly
    opened and closed.
    """

    def __enter__(self) -> Self: ...

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: TracebackType | None,
    ) -> None: ...

    def extract_table_by_column_names(
        self,
        column_names: list[str],
        exact_columns: bool = False,
        unmerge_cells: bool = True,
        fill_forward: bool = True,
    ) -> pl.DataFrame:
        """Search all sheets and return the first table containing *column_names*.

        All listed columns must be present in the same header row.  By default
        (``exact_columns=False``) the returned DataFrame contains every column
        in the detected range, which may include additional columns beyond those
        listed.  Set ``exact_columns=True`` to require the table to have no
        extra columns.

        Args:
            column_names: Column names that must all appear together in a
                header row.
            exact_columns: If ``True``, raise ``TableNotFoundError`` when the
                detected table has columns beyond those in *column_names*.
            unmerge_cells: Expand merged cells before extraction.
            fill_forward: Forward-fill null values after extraction.

        Returns:
            Polars DataFrame of the matched table.

        Raises:
            TableNotFoundError: No sheet contains a row with all *column_names*,
                or ``exact_columns=True`` and no sheet has an exact column set match.
            MultipleTablesFoundError: More than one row matches.
        """
        ...

    def extract_table_by_column_names_from_sheet(
        self,
        column_names: list[str],
        sheet_name: str,
        exact_columns: bool = False,
        unmerge_cells: bool = True,
        fill_forward: bool = True,
    ) -> pl.DataFrame:
        """Extract from a specific sheet by matching column names.

        Args:
            column_names: Column names that must all appear in the header row.
            sheet_name: Sheet to search.
            exact_columns: If ``True``, raise ``TableNotFoundError`` when the
                detected table has columns beyond those in *column_names*.
            unmerge_cells: Expand merged cells before extraction.
            fill_forward: Forward-fill null values after extraction.

        Returns:
            Polars DataFrame of the matched table.

        Raises:
            TableNotFoundError: Sheet does not contain all *column_names*, or
                ``exact_columns=True`` and the table has extra columns.
            MultipleTablesFoundError: More than one header row matches.
            ExcelTableReaderError: Sheet does not exist, or a column appears
                more than once in the header row.
        """
        ...

    def extract_table_by_range(
        self,
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
    ) -> pl.DataFrame:
        """Extract a table from an Excel range string.

        Two modes are supported:

        * **Exact** (``dynamic=False``, default) — fixed rectangle from *range_str*.
        * **Dynamic** (``dynamic=True``) — column span fixed by *range_str*; row
          boundary auto-detected downward.  ``max_empty_rows``, ``stop_at``, and
          ``stop_before`` apply only in this mode.

        Args:
            range_str: Excel range, e.g. ``"A5:C20"`` or ``"A1:D1"``.
            sheet: Sheet name (required).
            has_headers: Whether the first row of the range is a header row.
            column_names: Manual column names when *has_headers* is ``False``.
            dynamic: If ``True``, fix the column span but auto-detect the bottom row.
            unmerge_cells: Expand merged cells before extraction.
            fill_forward: Forward-fill null values after extraction.
            max_empty_rows: (dynamic only) Stop after this many consecutive empty rows.
            stop_at: (dynamic only) Stop when a row contains this value, including it.
            stop_before: (dynamic only) Stop when a row contains this value, excluding it.

        Returns:
            Polars DataFrame.

        Raises:
            ValueError: Both *stop_at* and *stop_before* are provided.
            ExcelTableReaderError: Sheet does not exist.
        """
        ...

    def extract_table_from_cell(
        self,
        start_cell: str,
        sheet: str,
        has_headers: bool = True,
        column_names: list[str] | None = None,
        unmerge_cells: bool = True,
        fill_forward: bool = True,
        max_empty_rows: int = 2,
        stop_at: str | None = None,
        stop_before: str | None = None,
    ) -> pl.DataFrame:
        """Extract a table from *start_cell*, auto-detecting its boundaries.

        Args:
            start_cell: Top-left cell address, e.g. ``"A5"``.
            sheet: Sheet name (required).
            has_headers: Whether the first row is a header row.
            column_names: Manual column names when *has_headers* is ``False``.
            unmerge_cells: Expand merged cells before extraction.
            fill_forward: Forward-fill null values after extraction.
            max_empty_rows: Stop after this many consecutive empty rows
                (ignored when *stop_at* or *stop_before* is set).
            stop_at: Stop when a row contains this value, including that row.
            stop_before: Stop when a row contains this value, excluding that row.

        Returns:
            Polars DataFrame.

        Raises:
            ValueError: Both *stop_at* and *stop_before* are provided.
        """
        ...

    def extract_table_near(
        self,
        column_names: list[str],
        sheet: str,
        ref_cell: str | None = None,
        keyword: str | None = None,
        exact_columns: bool = False,
        unmerge_cells: bool = True,
        fill_forward: bool = True,
        stop_at: str | None = None,
        stop_before: str | None = None,
    ) -> pl.DataFrame:
        """Find and extract a table by scanning from an anchor for matching column headers.

        Exactly one of *ref_cell* or *keyword* must be provided.

        Args:
            column_names: Column names that identify the header row.
            sheet: Sheet name (required).
            ref_cell: A1 address to start scanning from.
            keyword: Exact cell value used as the anchor.
            exact_columns: If ``True``, raise ``TableNotFoundError`` when the
                detected table has columns beyond those in *column_names*.
            unmerge_cells: Expand merged cells before extraction.
            fill_forward: Forward-fill null values after extraction.
            stop_at: Stop when a row contains this value, including that row.
            stop_before: Stop when a row contains this value, excluding that row.

        Returns:
            Polars DataFrame.

        Raises:
            ValueError: Both or neither of *ref_cell* / *keyword* are given, or
                both *stop_at* and *stop_before* are given.
            TableNotFoundError: Anchor or column headers not found, or
                ``exact_columns=True`` and the table has extra columns.
            MultipleTablesFoundError: *keyword* or header row matches more than once.
        """
        ...

