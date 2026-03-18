# loading workbook related errors
class ExcelError(Exception):
    """Base exception for Excel operations."""

    pass


class ExcelFileNotFoundError(ExcelError):
    """Excel file not found."""

    pass


class ExcelPermissionError(ExcelError):
    """No permission to access Excel file."""

    pass


class ExcelCorruptedError(ExcelError):
    """Excel file is corrupted or invalid."""

    pass


# template reading related errors
class TemplateReadError(Exception):
    """When reading marked cells"""

    pass


# table related errors
class ExcelTableReaderError(ExcelError):
    """Base exception for ExcelTableReader."""

    pass


class TableNotFoundError(ExcelTableReaderError):
    """Table with specified criteria not found."""

    pass


class MultipleTablesFoundError(ExcelTableReaderError):
    """Multiple table match found"""

    def __init__(self, message: str, found_in: list[str]): ...
