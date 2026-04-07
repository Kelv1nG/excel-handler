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


@pytest.fixture(scope="session")
def template_path():
    """Template with {{ }} tags on Sheet1 and Sheet2; EmptySheet has none."""
    return str(FIXTURES_DIR / "template.xlsx")


@pytest.fixture(scope="session")
def template_merge_preservation_path():
    """Template with merged title (A1:C1) and footer (A6:C6) around a table.
    Used to verify merged cells outside the insertion zone survive outer-join fills.
    """
    return str(FIXTURES_DIR / "template_merge_preservation.xlsx")


@pytest.fixture(scope="session")
def template_complex_merges_path():
    """Template with A1:C1 title, a 2x2 merge (A6:B7), and 1x3 footer (A9:C9) below the table."""
    return str(FIXTURES_DIR / "template_complex_merges.xlsx")


@pytest.fixture(scope="session")
def template_tight_footer_path():
    """Template where the footer merge (A4:C4) sits immediately below the last data row (no separator)."""
    return str(FIXTURES_DIR / "template_tight_footer.xlsx")


@pytest.fixture(scope="session")
def template_left_join_path():
    """Template with left join tag — no extra rows inserted, merges must be untouched."""
    return str(FIXTURES_DIR / "template_left_join.xlsx")


@pytest.fixture(scope="session")
def template_vertical_merge_path():
    """Template with a 3-row vertical merge (A5:A7) below last data row, with blank separator."""
    return str(FIXTURES_DIR / "template_vertical_merge.xlsx")


@pytest.fixture(scope="session")
def template_vertical_merge_adjacent_path():
    """Template with a 3-row vertical merge (A4:A6) immediately below the last data row."""
    return str(FIXTURES_DIR / "template_vertical_merge_adjacent.xlsx")


@pytest.fixture(scope="session")
def template_data_col_vertical_merge_path():
    """Template with a 3-row vertical merge in a DATA column (B4:B6) at the boundary row.

    Row 4 has A4='Status' (not in DF) and B4:B6 merged as 'Section Header' italic.
    Tests that _find_last_data_row stops before any multi-row merge (not just
    join-column merges), so the merge is treated as 'below' and shifts intact.
    """
    return str(FIXTURES_DIR / "template_data_col_vertical_merge.xlsx")


@pytest.fixture(scope="session")
def template_positional_fill_path():
    """Template with a {{ data | table(positional=True) }} tag for positional (no-join) filling.

    Sheet has a 2x3 region starting at B3 tagged with table(positional=True).
    No headers, no join column — the DataFrame is written positionally.
    """
    return str(FIXTURES_DIR / "template_positional_fill.xlsx")


@pytest.fixture(scope="session")
def template_collision_path():
    """Template with two {{ table(positional=True) }} tags whose regions overlap.

    Used to verify that ValueError is raised on collision detection.
    """
    return str(FIXTURES_DIR / "template_collision.xlsx")


@pytest.fixture(scope="session")
def template_record_path():
    """Template with dot-notation tags for record (single-row DataFrame) access.

    Sheet1 has {{ result.Company }}, {{ result.Revenue }}, {{ other.Quarter }},
    and a plain scalar {{ title }} for mixed-mode testing.
    """
    return str(FIXTURES_DIR / "template_record.xlsx")
