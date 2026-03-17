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


# table related errors
class ExcelDataExtractorError(ExcelError):
    """Base exception for ExcelDataExtractor."""

    pass


class TableNotFoundError(ExcelDataExtractorError):
    """Table with specified criteria not found."""

    pass
