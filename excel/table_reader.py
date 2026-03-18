from openpyxl.utils import coordinate_to_tuple, range_boundaries
import polars as pl
from typing import cast
from types import TracebackType
from excel.exceptions import (
    ExcelTableReaderError,
    TableNotFoundError,
    MultipleTablesFoundError,
)
from excel.utils import load_excel_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell


class ExcelTableReader:
    def __init__(self, filepath: str):
        self.filepath = filepath
        self._wb: Workbook | None = None

    @property
    def wb(self) -> Workbook:
        """
        Get loaded workbook.

        Returns:
            Workbook object

        Raises:
            ExcelTableReaderError: If workbook not loaded
        """
        if self._wb is None:
            raise ExcelTableReaderError(
                "Workbook not loaded. Use ExcelTableReader as context manager:\n"
                "    with ExcelTableReader('file.xlsx') as reader:\n"
                "        df = reader.extract_table_by_column_names([...])"
            )
        return self._wb

    def __enter__(self) -> "ExcelTableReader":
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

    def extract_table_by_column_names(
        self,
        required_columns: list[str],
        unmerge_cells: bool = True,
        fill_forward: bool = True,
    ) -> pl.DataFrame:

        for sheet_name in self.wb.sheetnames:
            try:
                df = self.extract_table_by_column_names_from_sheet(
                    required_columns,
                    sheet_name,
                    unmerge_cells,
                    fill_forward,
                )
            except TableNotFoundError:
                continue
            else:
                return df

        raise TableNotFoundError(
            f"Could not find table with columns {required_columns} "
            f"in any sheet of {self.filepath}"
        )

    def extract_table_by_column_names_from_sheet(
        self,
        required_columns: list[str],
        sheet_name: str,
        unmerge_cells: bool = True,
        fill_forward: bool = True,
    ) -> pl.DataFrame:
        """
        Internal method to extract from a specific sheet.

        Args:
            sheet_name: Name of sheet to extract from
            required_columns: Columns to find
            unmerge_cells: Unmerge cells
            fill_forward: Fill forward

        Returns:
            Polars DataFrame

        Raises:
            TableNotFoundError: If table not found in this sheet
        """
        sheet = self._get_sheet(sheet_name)

        if unmerge_cells:
            self._unmerge_and_fill_sheet(sheet)

        header_row = self._find_header_row_in_sheet(
            sheet,
            required_columns,
        )

        # Find the leftmost column that contains any cell value in the header row
        start_col = next(
            cell.column
            for cell in sheet[header_row]
            if cell.value is not None and cell.value != ""
        )
        boundaries = self._detect_boundaries_in_sheet(
            sheet, header_row, start_col, has_headers=True
        )

        df = self._extract_range_from_sheet(
            sheet,
            *boundaries,
            has_headers=True,
            column_names=None,
        )

        if fill_forward and len(df) > 0:
            df = df.with_columns([pl.col(col).forward_fill() for col in df.columns])

        return df

    def extract_table_by_range(
        self,
        range_str: str,
        sheet: str,
        has_headers: bool = True,
        column_names: list[str] | None = None,
        unmerge_cells: bool = True,
        fill_forward: bool = True,
    ) -> pl.DataFrame:
        """
        Extract table from explicit range.

        Args:
            range_str: Excel range (e.g., "A5:C20")
            sheet: Sheet name (REQUIRED for range extraction)
            has_headers: First row contains headers
            column_names: Manual column names if has_headers=False
            unmerge_cells: Unmerge cells
            fill_forward: Fill forward

        Returns:
            Polars DataFrame

        Example:
            >>> df = extractor.extract_by_range('A5:C20', sheet='Data')
            >>> df = extractor.extract_by_range(
            ...     'B10:D30',
            ...     sheet='Sheet1',
            ...     has_headers=False,
            ...     column_names=['ID', 'Name', 'Value']
            ... )
        """
        sheet_obj = self._get_sheet(sheet)

        if unmerge_cells:
            self._unmerge_and_fill_sheet(sheet_obj)

        min_col, min_row, max_col, max_row = range_boundaries(range_str)

        assert min_row is not None
        assert min_col is not None
        assert max_row is not None
        assert max_col is not None

        df = self._extract_range_from_sheet(
            sheet_obj,
            min_row,
            min_col,
            max_row,
            max_col,
            has_headers=has_headers,
            column_names=column_names,
        )

        if fill_forward and len(df) > 0:
            df = df.with_columns([pl.col(col).forward_fill() for col in df.columns])

        return df

    def extract_table_from_cell(
        self,
        start_cell: str,
        sheet: str,
        has_headers: bool = True,
        column_names: list[str] | None = None,
        unmerge_cells: bool = True,
        fill_forward: bool = True,
        max_empty_rows: int = 2,
    ) -> pl.DataFrame:
        """
        Extract table starting from cell, auto-detecting boundaries.

        Args:
            start_cell: Starting cell (e.g., "A5")
            sheet: Sheet name (REQUIRED)
            has_headers: First row is headers
            column_names: Manual column names
            unmerge_cells: Unmerge cells
            fill_forward: Fill forward
            max_empty_rows: Stop threshold

        Returns:
            Polars DataFrame

        Example:
            >>> df = extractor.extract_from_cell('A5', sheet='Data')
        """
        # Get sheet
        sheet_obj = self._get_sheet(sheet)

        # Step 1: Unmerge cells if requested
        if unmerge_cells:
            self._unmerge_and_fill_sheet(sheet_obj)

        # Step 2: Parse start cell
        start_row, start_col = coordinate_to_tuple(start_cell)

        # Step 3: Detect boundaries
        boundaries = self._detect_boundaries_in_sheet(
            sheet_obj, start_row, start_col, has_headers, max_empty_rows
        )

        # Step 4: Extract
        df = self._extract_range_from_sheet(
            sheet_obj,
            *boundaries,
            has_headers=has_headers,
            column_names=column_names,
        )

        # Step 5: Fill forward
        if fill_forward and len(df) > 0:
            df = df.with_columns([pl.col(col).forward_fill() for col in df.columns])

        return df

    # ==================== Utility Methods (INTERNAL) ====================

    def _get_sheet(self, sheet_name: str) -> Worksheet:
        """
        Get sheet, with validation.

        Args:
            sheet_name: Sheet name (if None, uses active sheet)

        Returns:
            Worksheet object

        Raises:
            ExcelDataExtractorError: If sheet not found
        """

        if sheet_name not in self.wb.sheetnames:
            raise ExcelTableReaderError(
                f"Sheet '{sheet_name}' not found in {self.filepath}. "
                f"Available sheets: {self.wb.sheetnames}"
            )

        return self.wb[sheet_name]

    def _unmerge_and_fill_sheet(self, sheet: Worksheet) -> None:
        """
        Unmerge all merged cells in the sheet and fill values forward.

        Args:
            sheet: Worksheet object
        """
        merged_ranges = list(sheet.merged_cells.ranges)

        for merged_range in merged_ranges:
            # Get bounds of merged range
            min_col, min_row, max_col, max_row = merged_range.bounds

            # Get value from top-left cell (where value is stored)
            value = sheet.cell(min_row, min_col).value

            # Unmerge
            sheet.unmerge_cells(str(merged_range))

            # Fill value to all cells in the range
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    cast(Cell, sheet.cell(row, col)).value = value

    def _detect_boundaries_in_sheet(
        self,
        sheet: Worksheet,
        start_row: int,
        start_col: int,
        has_headers: bool = True,
        max_empty_rows: int = 2,
    ) -> tuple[int, int, int, int]:
        """
        Auto-detect table boundaries from a starting point.

        Args:
            sheet: Worksheet object
            start_row: Starting row number (1-indexed)
            start_col: Starting column number (1-indexed)
            has_headers: Whether first row is headers
            max_empty_rows: Stop after this many empty rows

        Returns:
            Tuple of (min_row, min_col, max_row, max_col)
        """
        # Detect right boundary (columns)
        max_col = start_col
        for col in range(start_col, min(start_col + 100, sheet.max_column + 1)):
            # Check if this column has data in header row
            cell_value = sheet.cell(start_row, col).value

            if cell_value is None or cell_value == "":
                # Empty column - stop here
                break

            max_col = col

        # Detect bottom boundary (rows)
        data_start_row = start_row + 1 if has_headers else start_row
        max_row = start_row
        consecutive_empty = 0

        for row in range(data_start_row, sheet.max_row + 1):
            # Check if row has any data in our column range
            has_data = any(
                sheet.cell(row, col).value is not None
                and sheet.cell(row, col).value != ""
                for col in range(start_col, max_col + 1)
            )

            if has_data:
                max_row = row
                consecutive_empty = 0
            else:
                consecutive_empty += 1
                if consecutive_empty >= max_empty_rows:
                    # Too many empty rows, stop
                    break

        return (start_row, start_col, max_row, max_col)

    def _extract_range_from_sheet(
        self,
        sheet: Worksheet,
        min_row: int,
        min_col: int,
        max_row: int,
        max_col: int,
        has_headers: bool = True,
        column_names: list[str] | None = None,
    ) -> pl.DataFrame:
        """
        Extract data from specific cell range.

        Args:
            sheet: Worksheet object
            min_row, min_col: Top-left corner (1-indexed)
            max_row, max_col: Bottom-right corner (1-indexed)
            has_headers: Whether first row contains headers
            column_names: Manual column names if has_headers=False

        Returns:
            Polars DataFrame
        """
        # Step 1: Get headers
        headers: list[str]
        if has_headers:
            headers = [
                str(sheet.cell(min_row, col).value)
                for col in range(min_col, max_col + 1)
            ]
            data_start = min_row + 1
        else:
            # Use provided column names or generate default
            if column_names:
                headers = column_names
            else:
                headers = [f"col_{i}" for i in range(max_col - min_col + 1)]
            data_start = min_row

        # Step 2: Get data rows
        data = []
        for row in range(data_start, max_row + 1):
            row_data = [
                sheet.cell(row, col).value for col in range(min_col, max_col + 1)
            ]
            data.append(row_data)

        # Step 3: Create DataFrame
        if not data:
            # Empty table - return empty DataFrame with schema
            return pl.DataFrame(schema={h: pl.Utf8 for h in headers})

        return pl.DataFrame(data, schema=headers, orient="row")

    def _find_header_row_in_sheet(
        self,
        sheet: Worksheet,
        required_columns: list[str],
    ) -> int:
        """
        Find row containing all required column headers.

        Args:
            sheet: Worksheet object
            required_columns: List of column names to find

        Returns:
            Row number (1-indexed)

        Raises:
            TableNotFoundError: If no row contains all required columns
            MultipleTablesFoundError: If multiple rows match the required columns
            ExcelDataExtractorError: If required columns appear more than once in the matched row
        """
        matching_rows: list[int] = []

        for row_idx in range(1, sheet.max_row + 1):
            row_values = [
                cell.value for cell in sheet[row_idx] if cell.value is not None
            ]
            if all(col in row_values for col in required_columns):
                matching_rows.append(row_idx)

        if not matching_rows:
            raise TableNotFoundError(
                f"Could not find table with columns: {required_columns} "
                f"in sheet '{sheet.title}'"
            )

        if len(matching_rows) > 1:
            raise MultipleTablesFoundError(
                f"Required columns {required_columns} were found in multiple rows: {matching_rows}. "
                f"Cannot determine which table to extract.",
                found_in=[str(r) for r in matching_rows],
            )

        matched_row = matching_rows[0]
        row_values = [
            cell.value for cell in sheet[matched_row] if cell.value is not None
        ]
        duplicates = [col for col in required_columns if row_values.count(col) > 1]
        if duplicates:
            raise ExcelTableReaderError(
                f"Columns {duplicates} appear more than once in row {matched_row}. "
                f"Cannot determine which column to use."
            )

        return matched_row
