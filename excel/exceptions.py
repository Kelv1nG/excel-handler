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


class ExcelSheetNotFoundError(ExcelError):
    """Named worksheet does not exist in the workbook."""

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

    def __init__(self, message: str, found_in: list[str]):
        super().__init__(message)
        self.found_in = found_in


class KeywordNotFoundError(ExcelError):
    """Keyword label was not found in any searched sheet."""

    pass


class MultipleKeywordsFoundError(ExcelError):
    """Keyword label was found in more than one cell."""

    def __init__(self, message: str, found_in: list[str]):
        super().__init__(message)
        self.found_in = found_in


class ColumnNamesMismatchError(ExcelTableReaderError):
    """Number of provided column_names does not match the number of columns in the range."""

    def __init__(self, message: str, expected: int, got: int):
        super().__init__(message)
        self.expected = expected
        self.got = got
