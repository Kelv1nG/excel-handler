"""
Run this script to (re)generate all Excel test fixtures.

    uv run python tests/fixtures/create_fixtures.py
"""

from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

FIXTURES_DIR = Path(__file__).parent


def _header_style(ws, row: int, cols: range):
    """Bold + light-blue fill for header cells."""
    fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    for col in cols:
        cell = ws.cell(row, col)
        cell.font = Font(bold=True)
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")


# ---------------------------------------------------------------------------
# simple_table.xlsx
# Single clean table at A1 — Name / Amount / Category
# Tests: extract_table_by_column_names, partial columns, unordered columns,
#        extract_table_by_range, extract_table_from_cell, extract_table_near_cell
# ---------------------------------------------------------------------------
def create_simple_table():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "Amount", "Category"])
    ws.append(["Alice", 100, "Food"])
    ws.append(["Bob", 200, "Travel"])
    ws.append(["Carol", 300, "Food"])
    _header_style(ws, 1, range(1, 4))
    wb.save(FIXTURES_DIR / "simple_table.xlsx")
    wb.close()
    print("  simple_table.xlsx")


# ---------------------------------------------------------------------------
# no_headers.xlsx
# Pure data rows, no header row
# Tests: extract_table_by_range(has_headers=False, column_names=[...])
#        extract_table_by_range(has_headers=False)  → auto col_0, col_1, ...
# ---------------------------------------------------------------------------
def create_no_headers():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append([1, "Alice", 100])
    ws.append([2, "Bob", 200])
    ws.append([3, "Carol", 300])
    wb.save(FIXTURES_DIR / "no_headers.xlsx")
    wb.close()
    print("  no_headers.xlsx")


# ---------------------------------------------------------------------------
# merged_cells.xlsx
# Region column is merged across two rows per region
# Tests: unmerge_cells=True + fill_forward=True removes nulls
#        unmerge_cells=False + fill_forward=False leaves nulls
# ---------------------------------------------------------------------------
def create_merged_cells():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Region"
    ws["B1"] = "Country"
    ws["C1"] = "Sales"
    ws["A2"] = "Europe"
    ws["B2"] = "France"
    ws["C2"] = 100
    ws["B3"] = "Germany"
    ws["C3"] = 150
    ws["A4"] = "Asia"
    ws["B4"] = "Japan"
    ws["C4"] = 200
    ws["B5"] = "China"
    ws["C5"] = 250
    ws.merge_cells("A2:A3")  # "Europe" spans rows 2-3
    ws.merge_cells("A4:A5")  # "Asia"   spans rows 4-5
    _header_style(ws, 1, range(1, 4))
    wb.save(FIXTURES_DIR / "merged_cells.xlsx")
    wb.close()
    print("  merged_cells.xlsx")


# ---------------------------------------------------------------------------
# multiple_sheets.xlsx
# Same columns (Name, Amount) exist in Sheet1 AND Sheet2
# Tests: extract_table_by_column_names → returns first sheet match (Sheet1)
#        extract_table_by_column_names_from_sheet → targets Sheet2 specifically
# ---------------------------------------------------------------------------
def create_multiple_sheets():
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Sheet1"
    ws1.append(["Name", "Amount"])
    ws1.append(["Alice", 100])
    _header_style(ws1, 1, range(1, 3))

    ws2 = wb.create_sheet("Sheet2")
    ws2.append(["Name", "Amount"])
    ws2.append(["Bob", 200])
    _header_style(ws2, 1, range(1, 3))

    wb.save(FIXTURES_DIR / "multiple_sheets.xlsx")
    wb.close()
    print("  multiple_sheets.xlsx")


# ---------------------------------------------------------------------------
# multiple_tables.xlsx
# Two tables with identical headers on the same sheet (rows 1 and 3)
# Tests: extract_table_by_column_names → raises MultipleTablesFoundError
# ---------------------------------------------------------------------------
def create_multiple_tables():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "Amount"])  # row 1 — table 1 header
    ws.append(["Alice", 100])      # row 2 — table 1 data
    ws.append(["Name", "Amount"])  # row 3 — table 2 header
    ws.append(["Bob", 200])        # row 4 — table 2 data
    _header_style(ws, 1, range(1, 3))
    _header_style(ws, 3, range(1, 3))
    wb.save(FIXTURES_DIR / "multiple_tables.xlsx")
    wb.close()
    print("  multiple_tables.xlsx")


# ---------------------------------------------------------------------------
# empty_table.xlsx
# Headers present (Name / Amount) but no data rows below
# Tests: extract_table_by_column_names → returns empty DataFrame with schema
#        extract_table_near_cell from below headers → raises TableNotFoundError
# ---------------------------------------------------------------------------
def create_empty_table():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "Amount"])
    _header_style(ws, 1, range(1, 3))
    wb.save(FIXTURES_DIR / "empty_table.xlsx")
    wb.close()
    print("  empty_table.xlsx")


# ---------------------------------------------------------------------------
# offset_table.xlsx
# Table starting at C3 — rows 1-2 and columns A-B are empty
# Tests: extract_table_from_cell("C3") → finds table directly
#        extract_table_near_cell("A1") → scans right+down to C3
# ---------------------------------------------------------------------------
def create_offset_table():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["C3"] = "Name"
    ws["D3"] = "Amount"
    ws["E3"] = "Category"
    ws["C4"] = "Alice"
    ws["D4"] = 100
    ws["E4"] = "Food"
    ws["C5"] = "Bob"
    ws["D5"] = 200
    ws["E5"] = "Travel"
    _header_style(ws, 3, range(3, 6))
    wb.save(FIXTURES_DIR / "offset_table.xlsx")
    wb.close()
    print("  offset_table.xlsx")


# ---------------------------------------------------------------------------
# template.xlsx
# Template with {{ variable }} tags spread across two sheets.
# A third sheet (EmptySheet) has no tags — it must be excluded from results.
#
# Sheet1 layout:
#   B2  → {{ revenue }}
#   C4  → {{ title | orientation=horizontal }}
#   D6  → {{ count | skip=2, flag=True }}
#   A8  → plain text (no tag) — must not appear in results
#   B8  → 42  (numeric) — must not appear in results
#
# Sheet2 layout:
#   A1  → {{ summary }}
#
# Tests: ExcelTemplateReader.read() happy path + error cases,
#        MarkedCell.parse_metadata() coercion
# ---------------------------------------------------------------------------
def create_template():
    wb = Workbook()

    ws1 = wb.active
    ws1.title = "Sheet1"
    ws1["B2"] = "{{ revenue }}"
    ws1["C4"] = "{{ title | orientation=horizontal }}"
    ws1["D6"] = "{{ count | skip=2, flag=True }}"
    ws1["A8"] = "plain text — no tag here"
    ws1["B8"] = 42

    ws2 = wb.create_sheet("Sheet2")
    ws2["A1"] = "{{ summary }}"

    wb.create_sheet("EmptySheet")  # no tags — must be excluded from results

    wb.save(FIXTURES_DIR / "template.xlsx")
    wb.close()
    print("  template.xlsx")


if __name__ == "__main__":
    print("Creating fixtures in", FIXTURES_DIR)
    create_simple_table()
    create_no_headers()
    create_merged_cells()
    create_multiple_sheets()
    create_multiple_tables()
    create_empty_table()
    create_offset_table()
    create_template()
    print("Done.")
