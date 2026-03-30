from __future__ import annotations
import re
from dataclasses import dataclass, field
from os import PathLike
from typing import Any

from openpyxl.workbook.workbook import Workbook

from excel.protocols import TemplateReader
from excel.exceptions import ExcelFileNotFoundError, TemplateReadError
from excel._utils import load_excel_workbook

# Matches the first {{ ... }} tag in a cell value.
# Captures everything between the outer braces (lazy).
_TAG_RE = re.compile(r"\{\{\s*(.+?)\s*\}\}")

# Matches the function-call metadata form:  type(key=value, ...)
# Group 1 → type name, Group 2 → inner key=value content
_CALL_RE = re.compile(r"^(\w+)\((.*)\)$", re.DOTALL)


def _coerce_value(v: str) -> bool | int | float | str:
    """Convert a metadata value string to its most specific Python type."""
    if v.lower() == "true":
        return True
    if v.lower() == "false":
        return False
    try:
        return int(v)
    except ValueError:
        pass
    try:
        return float(v)
    except ValueError:
        pass
    return v


@dataclass
class MarkedCell:
    """A single {{...}} tag found inside a template workbook cell."""

    name: str
    """Variable name from the tag, e.g. ``revenue`` from ``{{ revenue }}``."""

    sheet: str
    """Worksheet name where this cell lives."""

    cell_addr: str
    """Cell address in A1 notation, e.g. ``"B5"``."""

    raw: str
    """Full original tag text including braces, e.g. ``"{{ revenue | loop }}"``."""

    metadata: str = field(default="")
    """Raw string after the ``|`` separator (stripped).
    Empty string when no metadata is present.
    Example: ``"orientation=horizontal, expand=True"``
    """

    def parse_metadata(self) -> dict[str, Any]:
        """Parse the metadata string into a typed dict.

        Two forms are supported:

        **Function-call form** (preferred)::

            {{ variable | type(key=value, key2=value2) }}

        The type name becomes ``result["type"]`` and each ``key=value`` pair
        inside the parentheses is added to the dict.

        **Flat key=value form** (for simple scalar metadata)::

            {{ variable | key=value, key2=value2 }}

        Values are coerced in both forms:

        * ``True`` / ``False``  ->  ``bool``
        * All-digit strings     ->  ``int``
        * Decimal strings       ->  ``float``
        * Everything else       ->  ``str``

        Returns an empty dict when there is no metadata.

        Examples::

            # Function-call form
            MarkedCell(..., metadata="table(join=outer, on=Sector)").parse_metadata()
            # {"type": "table", "join": "outer", "on": "Sector"}

            MarkedCell(..., metadata="loop()").parse_metadata()
            # {"type": "loop"}

            # Flat form
            MarkedCell(..., metadata="skip=2, flag=True").parse_metadata()
            # {"skip": 2, "flag": True}

            MarkedCell(..., metadata="").parse_metadata()
            # {}

        Raises:
            TemplateReadError: A metadata fragment is not in ``key=value`` form.
        """
        if not self.metadata:
            return {}

        call_match = _CALL_RE.match(self.metadata.strip())
        if call_match:
            result: dict[str, Any] = {"type": call_match.group(1)}
            inner = call_match.group(2).strip()
            if inner:
                for part in inner.split(","):
                    part = part.strip()
                    if not part:
                        continue
                    if "=" not in part:
                        raise TemplateReadError(
                            f"Invalid metadata fragment {part!r} in tag {self.raw!r} "
                            f"at {self.sheet}!{self.cell_addr}. Expected key=value format."
                        )
                    key, _, value = part.partition("=")
                    result[key.strip()] = _coerce_value(value.strip())
            return result

        # Flat key=value form
        result = {}
        for part in self.metadata.split(","):
            part = part.strip()
            if not part:
                continue
            if "=" not in part:
                raise TemplateReadError(
                    f"Invalid metadata fragment {part!r} in tag {self.raw!r} "
                    f"at {self.sheet}!{self.cell_addr}. Expected key=value format."
                )
            key, _, value = part.partition("=")
            result[key.strip()] = _coerce_value(value.strip())
        return result

type Worksheet = str
type WorksheetMarkedCells = dict[Worksheet, list[MarkedCell]]


class ExcelTemplateReader(TemplateReader[WorksheetMarkedCells]):
    """Scan an Excel workbook for ``{{ variable }}`` tags and return their locations.

    Tags may carry optional metadata after a ``|`` separator::

        {{ variable_name }}
        {{ variable_name | key=value, key2=value2 }}

    Usage::

        reader = ExcelTemplateReader()
        structure = reader.read("template.xlsx")
        # {"Sheet1": [MarkedCell(name="revenue", cell_addr="B3", ...), ...]}
    """

    def read(self, file: str | PathLike[str] | bytes) -> WorksheetMarkedCells:
        """Read *file* and return all tagged cells grouped by worksheet.

        Args:
            file: Path to the ``.xlsx`` template file.

        Returns:
            ``dict[sheet_name, list[MarkedCell]]``.
            Only sheets containing at least one tag are included.

        Raises:
            TemplateReadError: File not found, unreadable, or a tag has an
                empty variable name (e.g. ``{{ }}``, ``{{ | key=value }}``).
        """
        try:
            wb = load_excel_workbook(file, read_only=True)
        except ExcelFileNotFoundError as e:
            raise TemplateReadError(str(e)) from e
        except Exception as e:
            raise TemplateReadError(f"Failed to read template: {file}") from e

        try:
            result = self._process_workbook(wb)
        finally:
            wb.close()

        return result

    def _process_workbook(self, workbook: Workbook) -> WorksheetMarkedCells:
        """Scan every cell in every sheet for ``{{ }}`` tags.

        Args:
            workbook: An already-opened openpyxl Workbook.

        Returns:
            Populated ``WorksheetMarkedCells`` dict.

        Raises:
            TemplateReadError: A tag with an empty or missing variable name
                was found.
        """
        result: WorksheetMarkedCells = {}

        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
            marked: list[MarkedCell] = []

            for row in ws.iter_rows():
                for cell in row:
                    if not isinstance(cell.value, str):
                        continue
                    match = _TAG_RE.search(cell.value)
                    if match is None:
                        continue

                    content = match.group(1)
                    if "|" in content:
                        name, _, metadata = content.partition("|")
                        name = name.strip()
                        metadata = metadata.strip()
                    else:
                        name = content.strip()
                        metadata = ""

                    if not name:
                        raise TemplateReadError(
                            f"Empty variable name in tag {match.group(0)!r} "
                            f"at {sheet_name}!{cell.coordinate}"
                        )

                    marked.append(
                        MarkedCell(
                            name=name,
                            sheet=sheet_name,
                            cell_addr=cell.coordinate,
                            raw=match.group(0),
                            metadata=metadata,
                        )
                    )

            if marked:
                result[sheet_name] = marked

        return result
