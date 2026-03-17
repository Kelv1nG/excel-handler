from os import PathLike
from typing import Any
from openpyxl.workbook.workbook import Workbook

from abstract import TemplateReader
from excel.exceptions import ExcelFileNotFoundError, TemplateReadError
from excel.utils import load_excel_workbook


class MarkedCell:
    name: str
    cell_addr: str
    metadata: str

    def parse_metadata(self) -> dict[str, Any]: ...


type Worksheet = str
type WorksheetMarkedCells = dict[Worksheet, list[MarkedCell]]


class ExcelTemplateReader[WorksheetMarkedCells](TemplateReader):
    def read(self, file: str | PathLike[str]) -> WorksheetMarkedCells:
        """Read template and return marked cells structure."""

        try:
            wb = load_excel_workbook(file, read_only=True)
        except ExcelFileNotFoundError as e:
            raise TemplateReadError(str(e)) from e
        except Exception as e:
            raise TemplateReadError(f"Failed to read template: {file}") from e

        result = self._process_workbook(wb)

        wb.close()
        return result

    def _process_workbook(self, workbook: Workbook): ...
