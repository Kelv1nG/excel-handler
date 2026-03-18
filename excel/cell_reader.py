from typing import Any
from types import TracebackType
from openpyxl.utils import coordinate_to_tuple
from openpyxl.workbook.workbook import Workbook

from excel.exceptions import ExcelTableReaderError
from excel.utils import load_excel_workbook


class ExcelCellReader:
    """Read individual cell values from an Excel file."""

    def __init__(self, filepath: str):
        self.filepath = filepath
        self._wb: Workbook | None = None

    @property
    def wb(self) -> Workbook:
        if self._wb is None:
            raise ExcelTableReaderError(
                "Workbook not loaded. Use ExcelCellReader as context manager:\n"
                "    with ExcelCellReader('file.xlsx') as reader:\n"
                "        value = reader.get('Sheet1!B5')"
            )
        return self._wb

    def __enter__(self) -> "ExcelCellReader":
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

    def get(self, cell_ref: str) -> Any:
        """
        Get the value of a single cell.

        Args:
            cell_ref: Cell reference, either "Sheet1!B5" or "B5" (uses active sheet)

        Returns:
            Cell value (str, int, float, datetime, or None)

        Example:
            >>> with ExcelCellReader('config.xlsx') as reader:
            ...     rate = reader.get('Config!B5')
            ...     name = reader.get('A1')  # active sheet
        """
        sheet, cell_addr = self._parse_ref(cell_ref)
        row, col = coordinate_to_tuple(cell_addr)
        return sheet.cell(row, col).value

    def get_many(self, cell_refs: list[str]) -> dict[str, Any]:
        """
        Get values for multiple cell references.

        Args:
            cell_refs: List of cell references (e.g., ["Sheet1!B5", "Sheet2!C10"])

        Returns:
            Dict mapping each reference to its value

        Example:
            >>> with ExcelCellReader('config.xlsx') as reader:
            ...     values = reader.get_many(['Config!B5', 'Config!B6'])
            ...     # {'Config!B5': 0.15, 'Config!B6': 'USD'}
        """
        return {ref: self.get(ref) for ref in cell_refs}

    def _parse_ref(self, cell_ref: str):
        """Parse 'Sheet1!B5' into (sheet, 'B5'), falling back to active sheet."""
        if "!" in cell_ref:
            sheet_name, cell_addr = cell_ref.split("!", 1)
            if sheet_name not in self.wb.sheetnames:
                raise ExcelTableReaderError(
                    f"Sheet '{sheet_name}' not found in {self.filepath}. "
                    f"Available sheets: {self.wb.sheetnames}"
                )
            return self.wb[sheet_name], cell_addr
        else:
            active = self.wb.active
            if active is None:
                raise ExcelTableReaderError(
                    f"No active sheet found in {self.filepath}."
                )
            return active, cell_ref
