import pytest
from pathlib import Path

FIXTURES_DIR = Path(__file__).parent / "fixtures"
OUTPUT_DIR = Path(__file__).parent / "output"


def pytest_configure(config):
    """Ensure output directory exists for test results and artifacts."""
    OUTPUT_DIR.mkdir(exist_ok=True)
    # Configure pytest to use output directory for temporary files during test runs
    config.option.basetemp = OUTPUT_DIR / ".pytest_tmp"


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
def anchored_cells_path():
    """Labeled cells for get_relative / get_many_relative tests.

    Sheet1: A1='Revenue Label' B1=5000 C1='USD', A2='Tax Label' B2=250
    Sheet2: A1='Revenue Label' B1=9999  (duplicate for multi-keyword error tests)
    """
    return str(FIXTURES_DIR / "anchored_cells.xlsx")


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


@pytest.fixture(scope="session")
def template_sorted_outer_asc_path():
    """Template with table(join=outer, order_by=asc) — 2 upper template rows, Option A end_table."""
    return str(FIXTURES_DIR / "template_sorted_outer_asc.xlsx")


@pytest.fixture(scope="session")
def template_sorted_outer_desc_path():
    """Template with table(join=outer, order_by=desc) — 2 upper template rows, Option A end_table."""
    return str(FIXTURES_DIR / "template_sorted_outer_desc.xlsx")


@pytest.fixture(scope="session")
def template_sorted_outer_fixed_path():
    """Template with table(join=outer, order_by=asc) and a fixed lower zone via {{ insert_data }}."""
    return str(FIXTURES_DIR / "template_sorted_outer_fixed.xlsx")


@pytest.fixture(scope="session")
def template_sorted_outer_by_col_path():
    """Template with table(join=outer, order_by=Value:desc) — sort by non-join column."""
    return str(FIXTURES_DIR / "template_sorted_outer_by_col.xlsx")


@pytest.fixture(scope="session")
def template_sorted_outer_shorter_path():
    """Template with 3 upper zone template rows but df provides only 2 rows — template-only rows preserved."""
    return str(FIXTURES_DIR / "template_sorted_outer_shorter.xlsx")


@pytest.fixture(scope="session")
def template_sorted_outer_tmpl_rows_path():
    """Template with upper zone rows (foo, bar) absent from the df — must appear in sorted output."""
    return str(FIXTURES_DIR / "template_sorted_outer_tmpl_rows.xlsx")


@pytest.fixture(scope="session")
def template_fill_global_path():
    """Template with table(join=outer, fill=0) — every null in output gets 0."""
    return str(FIXTURES_DIR / "template_fill_global.xlsx")


@pytest.fixture(scope="session")
def template_fill_per_col_path():
    """Template with table(join=outer, fill=col1:0;col2:N/A) — per-column fill values."""
    return str(FIXTURES_DIR / "template_fill_per_col.xlsx")


@pytest.fixture(scope="session")
def template_fill_lower_zone_path():
    """Template with fill=0, insert_data, and an unmatched lower zone row — fill applies everywhere."""
    return str(FIXTURES_DIR / "template_fill_lower_zone.xlsx")


@pytest.fixture(scope="session")
def template_fill_sorted_outer_lower_zone_path():
    """Sorted outer join with fill=0 and an unmatched lower zone row (No Sector) — must get fill."""
    return str(FIXTURES_DIR / "template_fill_sorted_outer_lower_zone.xlsx")


@pytest.fixture(scope="session")
def bug_on_insert_path():
    """Template with a scalar tag ({{ some_value }}) one row below {{ end_table }}.
    Used to reproduce the stale-address bug where scalars below an expanding
    outer join table landed on the wrong row.
    """
    return str(FIXTURES_DIR / "bug-on-insert.xlsx")


@pytest.fixture(scope="session")
def template_placeholder_outer_path():
    """Template with placeholder=true tag and end_table|insert=above on Total row.
    Row 1: headers. Row 2: blank join col + tag (plain style). Row 3: Total (bold, yellow).
    """
    return str(FIXTURES_DIR / "template_placeholder_outer.xlsx")


@pytest.fixture(scope="session")
def template_style_src_last_path():
    """Template with default style (style=last). Row 2: plain tag row. Row 3: Total (bold, yellow).
    Used to verify inserted rows inherit the last template row's style (bold + yellow).
    """
    return str(FIXTURES_DIR / "template_style_src_last.xlsx")


@pytest.fixture(scope="session")
def template_style_src_first_path():
    """Template with style=first. Row 2: plain tag row. Row 3: Total (bold, yellow).
    Used to verify inserted rows inherit the tag row's plain style, NOT the Total row style.
    """
    return str(FIXTURES_DIR / "template_style_src_first.xlsx")


@pytest.fixture(scope="session")
def template_empty_outer_style_first_path():
    """Template with empty table: outer join + style=first + end_table|insert=above.
    Row 1: headers. Row 2: plain tag row (blank join col). Row 3: styled end_table marker.
    Used to verify: when DataFrame is filled, inserted rows copy style from tag row (plain),
    NOT from end_table row (bold + yellow).  Tests the edge case of empty template table.
    """
    return str(FIXTURES_DIR / "template_empty_outer_style_first.xlsx")


@pytest.fixture(scope="session")
def template_combo_outer_merges_below_path():
    """Outer join (placeholder=True) + 2 adjacent same-span merges below the table.
    DF a/b/c+Total → net +2 shift → merges land at A7:B7 and A8:B8.
    Primary regression test for the stale-merge-registry shift bug.
    """
    return str(FIXTURES_DIR / "template_combo_outer_merges_below.xlsx")


@pytest.fixture(scope="session")
def template_combo_left_with_merges_path():
    """Left join (no row insertion) + 2 adjacent merges below.
    Merges must be completely untouched since no rows are inserted.
    """
    return str(FIXTURES_DIR / "template_combo_left_with_merges.xlsx")


@pytest.fixture(scope="session")
def template_combo_scalar_with_outer_path():
    """Outer join (placeholder=True) + scalar cells below the table.
    Scalars must shift to their correct row after table expansion.
    """
    return str(FIXTURES_DIR / "template_combo_scalar_with_outer.xlsx")


@pytest.fixture(scope="session")
def template_combo_triple_adjacent_merges_path():
    """Three adjacent same-span (A:D) merges below an outer-join table.
    All three must survive after the +2 row-shift — direct replication of the
    stale-registry bug that silently dropped merges during _copy_row_styles.
    """
    return str(FIXTURES_DIR / "template_combo_triple_adjacent_merges.xlsx")


@pytest.fixture(scope="session")
def template_combo_two_outer_tables_path():
    """Two stacked outer-join tables (each inserting 1 row) + adjacent merge pair below.
    Merges must shift by the cumulative +2 (one from each table expansion).
    """
    return str(FIXTURES_DIR / "template_combo_two_outer_tables.xlsx")
