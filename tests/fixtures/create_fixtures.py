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




# ---------------------------------------------------------------------------
# template_merge_preservation.xlsx
# Template with:
#   - A1:C1 merged title cell (bold, centered, wrap_text) ABOVE the table
#   - A table at rows 3-4 with an outer-join tag
#   - A6:C6 merged footer cell (italic) BELOW the table
# Used to verify that merged cells outside the insertion zone survive intact
# after ExcelTemplateWriter inserts extra outer-join rows.
# ---------------------------------------------------------------------------
def create_template_merge_preservation():
    from openpyxl.styles import Border, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Title: A1:C1 merged — bold, centered, wrap_text, light-blue fill
    ws["A1"] = "Report Title"
    title_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    thin = Side(style="thin")
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].fill = title_fill
    ws["A1"].alignment = Alignment(horizontal="center", wrap_text=True)
    ws.merge_cells("A1:C1")

    # Headers row 2
    ws["A2"] = "Index"
    ws["B2"] = "col1"
    ws["C2"] = "col2"
    _header_style(ws, 2, range(1, 4))

    # Data rows 3-4 with outer-join tag in B3
    ws["A3"] = "a"
    ws["B3"] = "{{ data | table(join=outer, on=Index) }}"
    ws["A4"] = "b"

    # Empty row 5 (separator)

    # Footer: A6:C6 merged — italic
    ws["A6"] = "Footer Note"
    ws["A6"].font = Font(italic=True)
    ws["A6"].alignment = Alignment(horizontal="center")
    ws.merge_cells("A6:C6")

    wb.save(FIXTURES_DIR / "template_merge_preservation.xlsx")
    wb.close()
    print("  template_merge_preservation.xlsx")

# ---------------------------------------------------------------------------
# template_complex_merges.xlsx
# Template with:
#   - A1:C1 merged title ABOVE table (1x3)
#   - A table at rows 3-4 with outer-join tag
#   - A6:B7 merged 2x2 "BigNote" BELOW table
#   - A9:C9 merged 1x3 "Footer2" BELOW 2x2 merge
# Tests: multiple merges of different shapes all shift correctly on multi-insert
# ---------------------------------------------------------------------------
def create_template_complex_merges():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Complex Report"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center", wrap_text=True)
    ws.merge_cells("A1:C1")

    ws["A2"] = "Index"
    ws["B2"] = "col1"
    ws["C2"] = "col2"
    _header_style(ws, 2, range(1, 4))

    ws["A3"] = "a"
    ws["B3"] = "{{ data | table(join=outer, on=Index) }}"
    ws["A4"] = "b"

    # 2x2 merge at A6:B7
    ws["A6"] = "BigNote"
    ws["A6"].font = Font(bold=True, italic=True)
    ws["A6"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.merge_cells("A6:B7")

    # Footer at A9:C9
    ws["A9"] = "Footer2"
    ws["A9"].font = Font(italic=True, color="FF0000")
    ws["A9"].alignment = Alignment(horizontal="center")
    ws.merge_cells("A9:C9")

    wb.save(FIXTURES_DIR / "template_complex_merges.xlsx")
    wb.close()
    print("  template_complex_merges.xlsx")


# ---------------------------------------------------------------------------
# template_tight_footer.xlsx
# Footer merge immediately below last data row (no blank separator).
# ---------------------------------------------------------------------------
def create_template_tight_footer():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Index"
    ws["B1"] = "col1"
    ws["C1"] = "col2"
    _header_style(ws, 1, range(1, 4))

    ws["A2"] = "a"
    ws["B2"] = "{{ data | table(join=outer, on=Index) }}"
    ws["A3"] = "b"

    # Footer immediately at row 4 (no separator)
    ws["A4"] = "Tight Footer"
    ws["A4"].font = Font(italic=True)
    ws["A4"].alignment = Alignment(horizontal="center")
    ws.merge_cells("A4:C4")

    wb.save(FIXTURES_DIR / "template_tight_footer.xlsx")
    wb.close()
    print("  template_tight_footer.xlsx")


# ---------------------------------------------------------------------------
# template_left_join.xlsx
# Left join — no row insertion happens. Merges must be completely untouched.
# ---------------------------------------------------------------------------
def create_template_left_join():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Left Join Report"
    ws["A1"].font = Font(bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", wrap_text=True)
    ws.merge_cells("A1:C1")

    ws["A2"] = "Index"
    ws["B2"] = "col1"
    ws["C2"] = "col2"
    _header_style(ws, 2, range(1, 4))

    ws["A3"] = "a"
    ws["B3"] = "{{ data | table(join=left, on=Index) }}"
    ws["A4"] = "b"

    ws["A6"] = "Left Join Footer"
    ws["A6"].font = Font(italic=True, bold=True)
    ws["A6"].alignment = Alignment(horizontal="center")
    ws.merge_cells("A6:C6")

    wb.save(FIXTURES_DIR / "template_left_join.xlsx")
    wb.close()
    print("  template_left_join.xlsx")


# ---------------------------------------------------------------------------
# template_vertical_merge.xlsx
# Table with a 3-row vertical merge (A5:A7) immediately below the last data
# row (A4 is blank separator, A5:A7 = "Status" italic).
# Used to verify that the vertical merge is correctly shifted down (not split)
# when outer join inserts extra rows.
# Variant 2: merge IMMEDIATELY adjacent (no blank separator) at A4:A6.
# ---------------------------------------------------------------------------
def create_template_vertical_merge():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Headers
    ws["A1"] = "Index"
    ws["B1"] = "col1"
    ws["C1"] = "col2"
    _header_style(ws, 1, range(1, 4))

    # Data rows
    ws["A2"] = "a"
    ws["B2"] = "{{ data | table(join=outer, on=Index) }}"
    ws["A3"] = "b"

    # Row 4: blank separator

    # Vertical 3-row merge below the table
    ws["A5"] = "Status"
    ws["A5"].font = Font(italic=True)
    ws["A5"].alignment = Alignment(horizontal="center", wrap_text=True)
    ws.merge_cells("A5:A7")

    wb.save(FIXTURES_DIR / "template_vertical_merge.xlsx")
    wb.close()
    print("  template_vertical_merge.xlsx")


def create_template_vertical_merge_adjacent():
    """Vertical merge immediately below the last data row (no blank separator)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Index"
    ws["B1"] = "col1"
    ws["C1"] = "col2"
    _header_style(ws, 1, range(1, 4))

    ws["A2"] = "a"
    ws["B2"] = "{{ data | table(join=outer, on=Index) }}"
    ws["A3"] = "b"

    # Vertical 3-row merge IMMEDIATELY after last data row
    ws["A4"] = "Status"
    ws["A4"].font = Font(italic=True)
    ws["A4"].alignment = Alignment(horizontal="center", wrap_text=True)
    ws.merge_cells("A4:A6")

    wb.save(FIXTURES_DIR / "template_vertical_merge_adjacent.xlsx")
    wb.close()
    print("  template_vertical_merge_adjacent.xlsx")


def create_template_data_col_vertical_merge():
    """Vertical 3-row merge in a DATA column (not join col) at the boundary row.

    Template:
      Row 1: headers
      Row 2: a  {{tag|outer}}
      Row 3: b
      Row 4: Status (A4, not in DF)  +  B4:B6 = "Section Header" italic merge

    _find_last_data_row must stop at row 3 (before the merge at row 4) so
    that _copy_row_styles uses source_row=3, making B4:B6 a "below" merge
    that shifts down intact (to B5:B7 after 1 insertion, or B6:B8 for 2, etc.)
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Index"
    ws["B1"] = "col1"
    ws["C1"] = "col2"
    _header_style(ws, 1, range(1, 4))

    ws["A2"] = "a"
    ws["B2"] = "{{ data | table(join=outer, on=Index) }}"
    ws["A3"] = "b"

    # Row 4: join col has a value NOT in DF, data col has 3-row merge
    ws["A4"] = "Status"
    ws["B4"] = "Section Header"
    ws["B4"].font = Font(italic=True)
    ws["B4"].alignment = Alignment(horizontal="center", wrap_text=True)
    ws.merge_cells("B4:B6")

    wb.save(FIXTURES_DIR / "template_data_col_vertical_merge.xlsx")
    wb.close()
    print("  template_data_col_vertical_merge.xlsx")


# ---------------------------------------------------------------------------
# template_record.xlsx
# Template for testing record (dot-notation) variable access.
#
# Sheet1:
#   Row 1: label/value header
#   Row 2: "Company"  | {{ result.Company }}
#   Row 3: "Revenue"  | {{ result.Revenue }}
#   Row 4: "Quarter"  | {{ other.Quarter }}
#   Row 5: "Title"    | {{ title }}          <- plain scalar for mixed test
# ---------------------------------------------------------------------------
def create_template_record():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Field"
    ws["B1"] = "Value"
    _header_style(ws, 1, range(1, 3))

    ws["A2"] = "Company"
    ws["B2"] = "{{ result.Company }}"

    ws["A3"] = "Revenue"
    ws["B3"] = "{{ result.Revenue }}"

    ws["A4"] = "Quarter"
    ws["B4"] = "{{ other.Quarter }}"

    ws["A5"] = "Title"
    ws["B5"] = "{{ title }}"

    wb.save(FIXTURES_DIR / "template_record.xlsx")
    wb.close()
    print("  template_record.xlsx")


# ---------------------------------------------------------------------------
# template_sorted_outer_asc.xlsx
# Sorted outer join — ascending, no fixed zone, Option A end_table.
# Row 1: Index, Value  (headers)
# Row 2: c,  {{ data | table(join=outer, on=Index, order_by=asc) }}
# Row 3: a,  (empty)
# Row 4: (empty), {{ end_table }}   ← Option A: own row, no join value
# ---------------------------------------------------------------------------
def create_template_sorted_outer_asc():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Index"
    ws["B1"] = "Value"
    _header_style(ws, 1, range(1, 3))

    ws["A2"] = "c"
    ws["B2"] = "{{ data | table(join=outer, on=Index, order_by=asc) }}"
    ws["A3"] = "a"

    ws["B4"] = "{{ end_table }}"

    wb.save(FIXTURES_DIR / "template_sorted_outer_asc.xlsx")
    wb.close()
    print("  template_sorted_outer_asc.xlsx")


# ---------------------------------------------------------------------------
# template_sorted_outer_desc.xlsx
# Same layout as asc but tag uses order_by=desc.
# ---------------------------------------------------------------------------
def create_template_sorted_outer_desc():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Index"
    ws["B1"] = "Value"
    _header_style(ws, 1, range(1, 3))

    ws["A2"] = "c"
    ws["B2"] = "{{ data | table(join=outer, on=Index, order_by=desc) }}"
    ws["A3"] = "a"

    ws["B4"] = "{{ end_table }}"

    wb.save(FIXTURES_DIR / "template_sorted_outer_desc.xlsx")
    wb.close()
    print("  template_sorted_outer_desc.xlsx")


# ---------------------------------------------------------------------------
# template_sorted_outer_fixed.xlsx
# Sorted outer join with a fixed lower zone via {{ insert_data }}.
# Row 1: Index, Value  (headers)
# Row 2: c,  {{ data | table(join=outer, on=Index, order_by=asc) }}
# Row 3: a,  (empty)
# Row 4: (empty), {{ insert_data }}   ← insertion point marker
# Row 5: total, {{ end_table }}       ← Option B: on data row
# ---------------------------------------------------------------------------
def create_template_sorted_outer_fixed():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Index"
    ws["B1"] = "Value"
    _header_style(ws, 1, range(1, 3))

    ws["A2"] = "c"
    ws["B2"] = "{{ data | table(join=outer, on=Index, order_by=asc) }}"
    ws["A3"] = "a"

    ws["B4"] = "{{ insert_data }}"

    ws["A5"] = "total"
    ws["B5"] = "{{ end_table }}"

    wb.save(FIXTURES_DIR / "template_sorted_outer_fixed.xlsx")
    wb.close()
    print("  template_sorted_outer_fixed.xlsx")


# ---------------------------------------------------------------------------
# template_sorted_outer_by_col.xlsx
# Sorted outer join — sort by a non-join column (Value:desc).
# Row 1: Index, Value  (headers)
# Row 2: c,  {{ data | table(join=outer, on=Index, order_by=Value:desc) }}
# Row 3: a,  (empty)
# Row 4: (empty), {{ end_table }}
# ---------------------------------------------------------------------------
def create_template_sorted_outer_by_col():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Index"
    ws["B1"] = "Value"
    _header_style(ws, 1, range(1, 3))

    ws["A2"] = "c"
    ws["B2"] = "{{ data | table(join=outer, on=Index, order_by=Value:desc) }}"
    ws["A3"] = "a"

    ws["B4"] = "{{ end_table }}"

    wb.save(FIXTURES_DIR / "template_sorted_outer_by_col.xlsx")
    wb.close()
    print("  template_sorted_outer_by_col.xlsx")


# ---------------------------------------------------------------------------
# template_sorted_outer_shorter.xlsx
# Sorted outer join where df has fewer rows than upper zone template slots.
# Row 1: Index, Value  (headers)
# Row 2: c,  {{ data | table(join=outer, on=Index, order_by=asc) }}
# Row 3: a,  (empty)
# Row 4: b,  (empty)
# Row 5: (empty), {{ end_table }}   ← 3 upper slots, but df only provides 2 rows
# ---------------------------------------------------------------------------
def create_template_sorted_outer_shorter():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Index"
    ws["B1"] = "Value"
    _header_style(ws, 1, range(1, 3))

    ws["A2"] = "c"
    ws["B2"] = "{{ data | table(join=outer, on=Index, order_by=asc) }}"
    ws["A3"] = "a"
    ws["A4"] = "b"

    ws["B5"] = "{{ end_table }}"

    wb.save(FIXTURES_DIR / "template_sorted_outer_shorter.xlsx")
    wb.close()
    print("  template_sorted_outer_shorter.xlsx")


# ---------------------------------------------------------------------------
# template_sorted_outer_tmpl_rows.xlsx
# Sorted outer join where the upper zone has template-only rows (foo, bar)
# that are NOT present in the DataFrame — they must survive the sort.
# Row 1: Index, Value  (headers)
# Row 2: foo, {{ data | table(join=outer, on=Index, order_by=asc) }}
# Row 3: bar
# Row 4: (B4={{ insert_data }})         ← insertion point marker
# Row 5: No Sector                       ← lower zone (fixed)
# Row 6: Total, {{ end_table }}          ← lower zone fixed row + Option B end_table
# ---------------------------------------------------------------------------
def create_template_sorted_outer_tmpl_rows():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Index"
    ws["B1"] = "Value"
    _header_style(ws, 1, range(1, 3))

    ws["A2"] = "foo"
    ws["B2"] = "{{ data | table(join=outer, on=Index, order_by=asc) }}"
    ws["A3"] = "bar"

    ws["B4"] = "{{ insert_data }}"

    ws["A5"] = "No Sector"
    ws["A6"] = "Total"
    ws["B6"] = "{{ end_table }}"

    wb.save(FIXTURES_DIR / "template_sorted_outer_tmpl_rows.xlsx")
    wb.close()
    print("  template_sorted_outer_tmpl_rows.xlsx")


# ---------------------------------------------------------------------------
# template_fill_global.xlsx
# Outer join with fill=0 — every null data column in the output gets 0.
# Row 1: Index, col1, col2  (headers)
# Row 2: a,  {{ data | table(join=outer, on=Index, fill=0) }}
# Row 3: x  (template-only, not in df)
# Row 4: (empty), {{ end_table }}   ← Option A
# ---------------------------------------------------------------------------
def create_template_fill_global():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Index"
    ws["B1"] = "col1"
    ws["C1"] = "col2"
    _header_style(ws, 1, range(1, 4))

    ws["A2"] = "a"
    ws["B2"] = "{{ data | table(join=outer, on=Index, fill=0) }}"
    ws["A3"] = "x"

    ws["B4"] = "{{ end_table }}"

    wb.save(FIXTURES_DIR / "template_fill_global.xlsx")
    wb.close()
    print("  template_fill_global.xlsx")


# ---------------------------------------------------------------------------
# template_fill_per_col.xlsx
# Outer join with fill=col1:0;col2:N/A — per-column fill values.
# Same layout as template_fill_global.xlsx.
# ---------------------------------------------------------------------------
def create_template_fill_per_col():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Index"
    ws["B1"] = "col1"
    ws["C1"] = "col2"
    _header_style(ws, 1, range(1, 4))

    ws["A2"] = "a"
    ws["B2"] = "{{ data | table(join=outer, on=Index, fill=col1:0;col2:N/A) }}"
    ws["A3"] = "x"

    ws["B4"] = "{{ end_table }}"

    wb.save(FIXTURES_DIR / "template_fill_per_col.xlsx")
    wb.close()
    print("  template_fill_per_col.xlsx")


# ---------------------------------------------------------------------------
# template_fill_lower_zone.xlsx
# Outer join with fill=0, insert_data marker, and an unmatched lower zone row.
# The lower zone row has no df match — fill=0 must still apply to its nulls.
# Row 1: Index, col1  (headers)
# Row 2: a,  {{ data | table(join=outer, on=Index, fill=0) }}
# Row 3:     {{ insert_data }}    ← insertion point
# Row 4: Total                   ← lower zone (NOT in df; col1 → 0 via fill)
# Row 5:     {{ end_table }}     ← Option A
# ---------------------------------------------------------------------------
def create_template_fill_lower_zone():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Index"
    ws["B1"] = "col1"
    _header_style(ws, 1, range(1, 3))

    ws["A2"] = "a"
    ws["B2"] = "{{ data | table(join=outer, on=Index, fill=0) }}"

    ws["B3"] = "{{ insert_data }}"

    ws["A4"] = "Total"

    ws["B5"] = "{{ end_table }}"

    wb.save(FIXTURES_DIR / "template_fill_lower_zone.xlsx")
    wb.close()
    print("  template_fill_lower_zone.xlsx")


# ---------------------------------------------------------------------------
# template_fill_sorted_outer_lower_zone.xlsx
# Sorted outer join with fill=0, insert_data, and an unmatched lower zone row.
# The unmatched lower zone row must get fill=0, not stay blank.
# Row 1: Sector, ColValue  (headers)
# Row 2: Foo,  {{ data | table(join=outer, on=Sector, order_by=asc, fill=0) }}
# Row 3: Bar
# Row 4:       {{ insert_data }}    ← insertion point
# Row 5: No Sector                  ← lower zone: NOT in df → ColValue must be 0
# Row 6: Total  {{ end_table }}     ← lower zone: matched in df
# ---------------------------------------------------------------------------
def create_template_fill_sorted_outer_lower_zone():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Sector"
    ws["B1"] = "ColValue"
    _header_style(ws, 1, range(1, 3))

    ws["A2"] = "Foo"
    ws["B2"] = "{{ data | table(join=outer, on=Sector, order_by=asc, fill=0) }}"
    ws["A3"] = "Bar"

    ws["B4"] = "{{ insert_data }}"

    ws["A5"] = "No Sector"
    ws["A6"] = "Total"
    ws["B6"] = "{{ end_table }}"

    wb.save(FIXTURES_DIR / "template_fill_sorted_outer_lower_zone.xlsx")
    wb.close()
    print("  template_fill_sorted_outer_lower_zone.xlsx")


# ---------------------------------------------------------------------------
# anchored_cells.xlsx
# Fixture for ExcelCellReader.get_relative / get_many_relative tests.
#
# Sheet1 layout:
#   A1 = "Revenue Label"   B1 = 5000   C1 = "USD"
#   A2 = "Tax Label"       B2 = 250
#   A3 = "Note"
#
# Sheet2 layout:
#   A1 = "Revenue Label"   B1 = 9999   ← duplicate keyword for error tests
#
# Tests:
#   - get_relative(cell_ref="Sheet1!A1", right=1)   → 5000
#   - get_relative(cell_ref="Sheet1!A1", right=2)   → "USD"
#   - get_relative(cell_ref="Sheet1!B1", left=1)    → "Revenue Label"
#   - get_relative(cell_ref="Sheet1!A2", up=1)      → "Revenue Label"
#   - get_relative(keyword="Revenue Label", sheet="Sheet1", right=1)           → 5000
#   - get_relative(keyword="Tax Label",    sheet="Sheet1", right=1)            → 250
#   - get_relative(keyword="Revenue Label", sheet="Sheet1", down=1)            → "Tax Label"
#   - get_relative(keyword="Revenue Label")  → raises MultipleKeywordsFoundError
#   - get_relative(keyword="No Such Label") → raises KeywordNotFoundError
#   - get_many_relative(cell_ref="Sheet1!A1", offsets={...})  → dict
#   - get_many_relative(keyword="Tax Label", sheet="Sheet1", offsets={...}) → dict
# ---------------------------------------------------------------------------
def create_anchored_cells():
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Sheet1"
    ws1["A1"] = "Revenue Label"
    ws1["B1"] = 5000
    ws1["C1"] = "USD"
    ws1["A2"] = "Tax Label"
    ws1["B2"] = 250
    ws1["A3"] = "Note"

    ws2 = wb.create_sheet("Sheet2")
    ws2["A1"] = "Revenue Label"
    ws2["B1"] = 9999

    wb.save(FIXTURES_DIR / "anchored_cells.xlsx")
    wb.close()
    print("  anchored_cells.xlsx")


# ---------------------------------------------------------------------------
# template_placeholder_outer.xlsx
# Template for placeholder=true + end_table|insert=above tests.
# Row 1: headers (Index, Value)
# Row 2: blank Index, {{ data | table(join=outer, placeholder=true) }} — plain
# Row 3: Total, {{ end_table | insert=above }} — bold + yellow fill
# Tests: placeholder row deleted when unmatched; Total pinned via Option C
# ---------------------------------------------------------------------------
def create_template_placeholder_outer():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Index"
    ws["B1"] = "Value"
    _header_style(ws, 1, range(1, 3))

    # Row 2: blank join col, tag cell (plain style — no bold, no fill)
    ws["B2"] = "{{ data | table(join=outer, placeholder=true) }}"

    # Row 3: Total row — bold + yellow fill
    total_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws["A3"] = "Total"
    ws["B3"] = "{{ end_table | insert=above }}"
    for col in range(1, 3):
        ws.cell(3, col).font = Font(bold=True)
        ws.cell(3, col).fill = total_fill

    wb.save(FIXTURES_DIR / "template_placeholder_outer.xlsx")
    wb.close()
    print("  template_placeholder_outer.xlsx")


# ---------------------------------------------------------------------------
# template_style_src_last.xlsx — style=last (default)
# template_style_src_first.xlsx — style=first
# Both: Row 1 headers, Row 2 plain data row (tag), Row 3 Total (bold + yellow)
# Tests: style=last → inserted rows inherit bold/yellow;
#        style=first → inserted rows inherit plain style
# ---------------------------------------------------------------------------
def create_template_style_src_last():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Index"
    ws["B1"] = "Value"
    _header_style(ws, 1, range(1, 3))

    ws["A2"] = "a"
    ws["B2"] = "{{ data | table(join=outer) }}"

    total_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws["A3"] = "Total"
    ws["B3"] = 100
    for col in range(1, 3):
        ws.cell(3, col).font = Font(bold=True)
        ws.cell(3, col).fill = total_fill

    wb.save(FIXTURES_DIR / "template_style_src_last.xlsx")
    wb.close()
    print("  template_style_src_last.xlsx")


def create_template_style_src_first():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "Index"
    ws["B1"] = "Value"
    _header_style(ws, 1, range(1, 3))

    ws["A2"] = "a"
    ws["B2"] = "{{ data | table(join=outer, style=first) }}"

    total_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws["A3"] = "Total"
    ws["B3"] = 100
    for col in range(1, 3):
        ws.cell(3, col).font = Font(bold=True)
        ws.cell(3, col).fill = total_fill

    wb.save(FIXTURES_DIR / "template_style_src_first.xlsx")
    wb.close()
    print("  template_style_src_first.xlsx")


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
    create_template_merge_preservation()
    create_template_complex_merges()
    create_template_tight_footer()
    create_template_left_join()
    create_template_vertical_merge()
    create_template_vertical_merge_adjacent()
    create_template_data_col_vertical_merge()
    create_template_record()
    create_template_sorted_outer_asc()
    create_template_sorted_outer_desc()
    create_template_sorted_outer_fixed()
    create_template_sorted_outer_by_col()
    create_template_sorted_outer_shorter()
    create_template_sorted_outer_tmpl_rows()
    create_template_fill_global()
    create_template_fill_per_col()
    create_template_fill_lower_zone()
    create_template_fill_sorted_outer_lower_zone()
    create_anchored_cells()
    create_template_placeholder_outer()
    create_template_style_src_last()
    create_template_style_src_first()
    print("Done.")
