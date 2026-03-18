import pytest
from pathlib import Path

FIXTURES_DIR = Path(__file__).parent / "fixtures"


@pytest.fixture(scope="session")
def simple_table_path():
    """Single clean table at A1 - Name / Amount / Category, 3 data rows."""
    return str(FIXTURES_DIR / "simple_table.xlsx")


@pytest.fixture(scope="session")
def no_headers_path():
    """Pure data rows with no header row - 3 rows of (ID, Name, Value)."""
    return str(FIXTURES_DIR / "no_headers.xlsx")


@pytest.fixture(scope="session")
def merged_cells_path():
    """Hierarchical table: Region column is merged across two rows per region."""
    return str(FIXTURES_DIR / "merged_cells.xlsx")


@pytest.fixture(scope="session")
def multiple_sheets_path():
    """Same columns (Name, Amount) appear in both Sheet1 and Sheet2."""
    return str(FIXTURES_DIR / "multiple_sheets.xlsx")


@pytest.fixture(scope="session")
def multiple_tables_path():
    """Two tables with identical headers (Name, Amount) on the same sheet."""
    return str(FIXTURES_DIR / "multiple_tables.xlsx")


@pytest.fixture(scope="session")
def empty_table_path():
    """Headers present (Name, Amount) but no data rows below them."""
    return str(FIXTURES_DIR / "empty_table.xlsx")


@pytest.fixture(scope="session")
def offset_table_path():
    """Table starting at C3 - rows 1-2 and columns A-B are empty."""
    return str(FIXTURES_DIR / "offset_table.xlsx")
