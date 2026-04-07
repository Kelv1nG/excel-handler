"""Tests for ExcelTemplateWriter.

Focuses on merge cell preservation during table fills with row insertion
(outer join), which is the scenario most likely to corrupt cells outside
the insertion zone.
"""
import pytest
import polars as pl
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

from excel.template_writer import ExcelTemplateWriter
from excel.protocols import TypedValue


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _run_outer_join(template_path, tmp_path):
    """Fill template_merge_preservation.xlsx with a 3-row outer-join DF.

    Template has rows a, b.  DF has a, b, c — so row c is extra and triggers
    _copy_row_styles (one row insertion).

    Returns the opened output workbook (caller must close it).
    """
    df = pl.DataFrame({
        "Index": ["a", "b", "c"],
        "col1": [10, 20, 30],
        "col2": [100, 200, 300],
    })
    output = str(tmp_path / "output.xlsx")
    writer = ExcelTemplateWriter(template_path)
    writer.write({"data": TypedValue(df, "table")}, output)
    return load_workbook(output), output


def _merge_ranges(ws):
    """Return the set of merge range strings for a worksheet."""
    return {str(m) for m in ws.merged_cells.ranges}


# ---------------------------------------------------------------------------
# Merged title cell (A1:C1) — sits ABOVE the table and insertion zone
# ---------------------------------------------------------------------------

class TestMergedTitleSurvives:
    def test_title_merge_range_intact(self, template_merge_preservation_path, tmp_path):
        wb, _ = _run_outer_join(template_merge_preservation_path, tmp_path)
        ws = wb.active
        assert "A1:C1" in _merge_ranges(ws), (
            "Title merge A1:C1 was destroyed by row insertion"
        )

    def test_title_value_intact(self, template_merge_preservation_path, tmp_path):
        wb, _ = _run_outer_join(template_merge_preservation_path, tmp_path)
        ws = wb.active
        assert ws["A1"].value == "Report Title", (
            "Title cell value was wiped by row insertion"
        )

    def test_title_font_bold_intact(self, template_merge_preservation_path, tmp_path):
        wb, _ = _run_outer_join(template_merge_preservation_path, tmp_path)
        ws = wb.active
        assert ws["A1"].font.bold is True, (
            "Title bold font was stripped by row insertion"
        )

    def test_title_alignment_center_intact(self, template_merge_preservation_path, tmp_path):
        wb, _ = _run_outer_join(template_merge_preservation_path, tmp_path)
        ws = wb.active
        assert ws["A1"].alignment.horizontal == "center", (
            "Title alignment was stripped by row insertion"
        )

    def test_title_wrap_text_intact(self, template_merge_preservation_path, tmp_path):
        wb, _ = _run_outer_join(template_merge_preservation_path, tmp_path)
        ws = wb.active
        assert ws["A1"].alignment.wrap_text is True, (
            "Title wrap_text was stripped by row insertion"
        )


# ---------------------------------------------------------------------------
# Merged footer cell (A6:C6 in template) — sits BELOW the insertion zone
# One extra row is inserted, so footer shifts to A7:C7.
# ---------------------------------------------------------------------------

class TestMergedFooterShifts:
    def test_footer_shifts_to_row_7(self, template_merge_preservation_path, tmp_path):
        wb, _ = _run_outer_join(template_merge_preservation_path, tmp_path)
        ws = wb.active
        assert "A7:C7" in _merge_ranges(ws), (
            "Footer merge did not shift to A7:C7 after one row insertion"
        )

    def test_original_footer_row_is_gone(self, template_merge_preservation_path, tmp_path):
        wb, _ = _run_outer_join(template_merge_preservation_path, tmp_path)
        ws = wb.active
        assert "A6:C6" not in _merge_ranges(ws), (
            "Footer merge A6:C6 still present — should have shifted to A7:C7"
        )

    def test_footer_value_intact(self, template_merge_preservation_path, tmp_path):
        wb, _ = _run_outer_join(template_merge_preservation_path, tmp_path)
        ws = wb.active
        assert ws["A7"].value == "Footer Note", (
            "Footer value was lost after merge shift"
        )

    def test_footer_font_italic_intact(self, template_merge_preservation_path, tmp_path):
        wb, _ = _run_outer_join(template_merge_preservation_path, tmp_path)
        ws = wb.active
        assert ws["A7"].font.italic is True, (
            "Footer italic font was stripped after merge shift"
        )


# ---------------------------------------------------------------------------
# Data correctness — all three rows (a, b, c) should be filled
# ---------------------------------------------------------------------------

class TestOuterJoinDataCorrect:
    def test_row_a_col1(self, template_merge_preservation_path, tmp_path):
        wb, _ = _run_outer_join(template_merge_preservation_path, tmp_path)
        ws = wb.active
        assert ws["B3"].value == 10

    def test_row_a_col2(self, template_merge_preservation_path, tmp_path):
        wb, _ = _run_outer_join(template_merge_preservation_path, tmp_path)
        ws = wb.active
        assert ws["C3"].value == 100

    def test_row_b_col1(self, template_merge_preservation_path, tmp_path):
        wb, _ = _run_outer_join(template_merge_preservation_path, tmp_path)
        ws = wb.active
        assert ws["B4"].value == 20

    def test_row_b_col2(self, template_merge_preservation_path, tmp_path):
        wb, _ = _run_outer_join(template_merge_preservation_path, tmp_path)
        ws = wb.active
        assert ws["C4"].value == 200

    def test_extra_row_c_index(self, template_merge_preservation_path, tmp_path):
        wb, _ = _run_outer_join(template_merge_preservation_path, tmp_path)
        ws = wb.active
        assert ws["A5"].value == "c"

    def test_extra_row_c_col1(self, template_merge_preservation_path, tmp_path):
        wb, _ = _run_outer_join(template_merge_preservation_path, tmp_path)
        ws = wb.active
        assert ws["B5"].value == 30

    def test_extra_row_c_col2(self, template_merge_preservation_path, tmp_path):
        wb, _ = _run_outer_join(template_merge_preservation_path, tmp_path)
        ws = wb.active
        assert ws["C5"].value == 300


# ---------------------------------------------------------------------------
# No phantom merges on inserted rows
# ---------------------------------------------------------------------------

class TestNoPhantomMergesOnInsertedRows:
    def test_inserted_row_5_has_no_merge(self, template_merge_preservation_path, tmp_path):
        wb, _ = _run_outer_join(template_merge_preservation_path, tmp_path)
        ws = wb.active
        # Row 5 is the inserted extra row — it must not be inside any merge
        for m in ws.merged_cells.ranges:
            assert not (m.min_row <= 5 <= m.max_row), (
                f"Inserted row 5 is inside merge {m} — phantom merge not cleaned up"
            )

    def test_data_rows_are_real_cells_not_merged(self, template_merge_preservation_path, tmp_path):
        wb, _ = _run_outer_join(template_merge_preservation_path, tmp_path)
        ws = wb.active
        # B3, B4, B5 (data cols) must be real Cell objects, not MergedCell proxies
        for row in (3, 4, 5):
            cell = ws.cell(row, 2)  # col B
            assert not isinstance(cell, MergedCell), (
                f"B{row} is a MergedCell proxy — phantom merge ghost not purged"
            )


# ---------------------------------------------------------------------------
# Multiple insertions + 2x2 and 1x3 merges below table
# Template: A1:C1 title, table at rows 3-4, 2x2 merge A6:B7, footer A9:C9
# DF has a, b, c, d → inserts 2 extra rows → merges shift by 2
# ---------------------------------------------------------------------------

def _run_complex_merges(template_path, tmp_path):
    """3 extras (c, d, e) → 3 insertions → A6:B7 shifts to A9:B10, A9:C9 → A12:C12."""
    df = pl.DataFrame({
        "Index": ["a", "b", "c", "d", "e"],
        "col1": [1, 2, 3, 4, 5],
        "col2": [10, 20, 30, 40, 50],
    })
    output = str(tmp_path / "complex_output.xlsx")
    writer = ExcelTemplateWriter(template_path)
    writer.write({"data": TypedValue(df, "table")}, output)
    return load_workbook(output), output


class TestComplexMergesShift:
    def test_title_range_intact(self, template_complex_merges_path, tmp_path):
        wb, _ = _run_complex_merges(template_complex_merges_path, tmp_path)
        ws = wb.active
        assert "A1:C1" in _merge_ranges(ws), "Title A1:C1 was destroyed"

    def test_title_value_intact(self, template_complex_merges_path, tmp_path):
        wb, _ = _run_complex_merges(template_complex_merges_path, tmp_path)
        ws = wb.active
        assert ws["A1"].value == "Complex Report"

    def test_2x2_merge_shifts_by_3(self, template_complex_merges_path, tmp_path):
        # Original A6:B7; 3 extras inserted → A9:B10
        wb, _ = _run_complex_merges(template_complex_merges_path, tmp_path)
        ws = wb.active
        assert "A9:B10" in _merge_ranges(ws), (
            f"2x2 merge did not shift to A9:B10. Found: {_merge_ranges(ws)}"
        )

    def test_2x2_merge_old_range_gone(self, template_complex_merges_path, tmp_path):
        wb, _ = _run_complex_merges(template_complex_merges_path, tmp_path)
        ws = wb.active
        assert "A6:B7" not in _merge_ranges(ws), "Stale A6:B7 still present"

    def test_2x2_merge_value_preserved(self, template_complex_merges_path, tmp_path):
        wb, _ = _run_complex_merges(template_complex_merges_path, tmp_path)
        ws = wb.active
        assert ws["A9"].value == "BigNote", (
            f"2x2 merge value lost after shift. Got: {ws['A9'].value!r}"
        )

    def test_2x2_merge_bold_preserved(self, template_complex_merges_path, tmp_path):
        wb, _ = _run_complex_merges(template_complex_merges_path, tmp_path)
        ws = wb.active
        assert ws["A9"].font.bold is True, "2x2 merge bold font lost after shift"

    def test_2x2_merge_italic_preserved(self, template_complex_merges_path, tmp_path):
        wb, _ = _run_complex_merges(template_complex_merges_path, tmp_path)
        ws = wb.active
        assert ws["A9"].font.italic is True, "2x2 merge italic font lost after shift"

    def test_footer2_shifts_by_3(self, template_complex_merges_path, tmp_path):
        # Original A9:C9; 3 extras inserted → A12:C12
        wb, _ = _run_complex_merges(template_complex_merges_path, tmp_path)
        ws = wb.active
        assert "A12:C12" in _merge_ranges(ws), (
            f"Footer2 did not shift to A12:C12. Found: {_merge_ranges(ws)}"
        )

    def test_footer2_value_preserved(self, template_complex_merges_path, tmp_path):
        wb, _ = _run_complex_merges(template_complex_merges_path, tmp_path)
        ws = wb.active
        assert ws["A12"].value == "Footer2", (
            f"Footer2 value lost. Got: {ws['A12'].value!r}"
        )

    def test_all_5_data_rows_filled(self, template_complex_merges_path, tmp_path):
        wb, _ = _run_complex_merges(template_complex_merges_path, tmp_path)
        ws = wb.active
        assert ws["A3"].value == "a"
        assert ws["B3"].value == 1
        assert ws["A4"].value == "b"
        assert ws["B4"].value == 2
        # Extra rows at 5, 6, 7
        assert ws["A5"].value == "c"
        assert ws["B5"].value == 3
        assert ws["A6"].value == "d"
        assert ws["B6"].value == 4
        assert ws["A7"].value == "e"
        assert ws["B7"].value == 5


# ---------------------------------------------------------------------------
# Tight footer — merge immediately below last data row (no separator)
# Template: headers row 1, data rows 2-3, footer A4:C4
# DF has a, b, c → 1 extra → footer shifts to A5:C5
# ---------------------------------------------------------------------------

def _run_tight_footer(template_path, tmp_path):
    df = pl.DataFrame({
        "Index": ["a", "b", "c"],
        "col1": [10, 20, 30],
        "col2": [100, 200, 300],
    })
    output = str(tmp_path / "tight_output.xlsx")
    writer = ExcelTemplateWriter(template_path)
    writer.write({"data": TypedValue(df, "table")}, output)
    return load_workbook(output), output


class TestTightFooterShifts:
    """Without {{ end_table }}, _find_last_data_row includes the footer row
    (it has a non-empty join column) as last_tmpl_row.  The extra row 'c' is
    inserted AFTER the footer (at row 5), not before it.  This documents the
    expected heuristic behaviour: users must use {{ end_table }} to separate
    a footer from the table when there is no blank separator row.
    """

    def test_footer_stays_at_row_4(self, template_tight_footer_path, tmp_path):
        # last_tmpl_row=4 (footer included), extra inserted at row 5 after it
        wb, _ = _run_tight_footer(template_tight_footer_path, tmp_path)
        ws = wb.active
        assert "A4:C4" in _merge_ranges(ws), (
            f"Footer merge was unexpectedly moved. Found: {_merge_ranges(ws)}"
        )

    def test_footer_value_at_row_4(self, template_tight_footer_path, tmp_path):
        wb, _ = _run_tight_footer(template_tight_footer_path, tmp_path)
        ws = wb.active
        assert ws["A4"].value == "Tight Footer"

    def test_footer_italic_at_row_4(self, template_tight_footer_path, tmp_path):
        wb, _ = _run_tight_footer(template_tight_footer_path, tmp_path)
        ws = wb.active
        assert ws["A4"].font.italic is True, "Footer italic font missing at A4"

    def test_extra_row_c_at_row_5(self, template_tight_footer_path, tmp_path):
        # Extra row goes AFTER the footer because footer is treated as table row
        wb, _ = _run_tight_footer(template_tight_footer_path, tmp_path)
        ws = wb.active
        assert ws["A5"].value == "c"

    def test_data_rows_a_b_correct(self, template_tight_footer_path, tmp_path):
        wb, _ = _run_tight_footer(template_tight_footer_path, tmp_path)
        ws = wb.active
        assert ws["A2"].value == "a" and ws["B2"].value == 10
        assert ws["A3"].value == "b" and ws["B3"].value == 20


# ---------------------------------------------------------------------------
# Left join — no row insertion. Merges must be COMPLETELY untouched.
# ---------------------------------------------------------------------------

def _run_left_join(template_path, tmp_path):
    # DF has exactly a and b — same as template, no extras
    df = pl.DataFrame({
        "Index": ["a", "b"],
        "col1": [10, 20],
        "col2": [100, 200],
    })
    output = str(tmp_path / "left_output.xlsx")
    writer = ExcelTemplateWriter(template_path)
    writer.write({"data": TypedValue(df, "table")}, output)
    return load_workbook(output), output


class TestLeftJoinMergesUntouched:
    def test_title_range_unchanged(self, template_left_join_path, tmp_path):
        wb, _ = _run_left_join(template_left_join_path, tmp_path)
        ws = wb.active
        assert "A1:C1" in _merge_ranges(ws), "Title merge was modified on left join"

    def test_title_value_unchanged(self, template_left_join_path, tmp_path):
        wb, _ = _run_left_join(template_left_join_path, tmp_path)
        ws = wb.active
        assert ws["A1"].value == "Left Join Report"

    def test_title_bold_unchanged(self, template_left_join_path, tmp_path):
        wb, _ = _run_left_join(template_left_join_path, tmp_path)
        ws = wb.active
        assert ws["A1"].font.bold is True, "Title bold lost on left join"

    def test_footer_range_unchanged(self, template_left_join_path, tmp_path):
        wb, _ = _run_left_join(template_left_join_path, tmp_path)
        ws = wb.active
        assert "A6:C6" in _merge_ranges(ws), "Footer range shifted when it shouldn't"

    def test_footer_value_unchanged(self, template_left_join_path, tmp_path):
        wb, _ = _run_left_join(template_left_join_path, tmp_path)
        ws = wb.active
        assert ws["A6"].value == "Left Join Footer"

    def test_footer_italic_bold_unchanged(self, template_left_join_path, tmp_path):
        wb, _ = _run_left_join(template_left_join_path, tmp_path)
        ws = wb.active
        assert ws["A6"].font.italic is True
        assert ws["A6"].font.bold is True

    def test_data_filled_correctly(self, template_left_join_path, tmp_path):
        wb, _ = _run_left_join(template_left_join_path, tmp_path)
        ws = wb.active
        assert ws["B3"].value == 10
        assert ws["C3"].value == 100
        assert ws["B4"].value == 20
        assert ws["C4"].value == 200


# ---------------------------------------------------------------------------
# Vertical merge below table: with blank separator (A5:A7)
# Template: rows 2-3 are data (a,b), row 4 blank, rows 5-7 vertical merge "Status"
# DF has a, b, c — one extra row triggers one insertion at row 4 (after row 3)
# Expected: merge shifts from A5:A7 → A6:A8; value and italic preserved
# ---------------------------------------------------------------------------

def _run_vertical_merge(template_path, tmp_path, extra_indices=None):
    """Fill the vertical merge template with optional extra rows."""
    if extra_indices is None:
        extra_indices = ["c"]
    indices = ["a", "b"] + extra_indices
    df = pl.DataFrame({
        "Index": indices,
        "col1": list(range(len(indices))),
        "col2": list(range(100, 100 + len(indices))),
    })
    output = str(tmp_path / "output_vm.xlsx")
    writer = ExcelTemplateWriter(template_path)
    writer.write({"data": TypedValue(df, "table")}, output)
    return load_workbook(output), output


class TestVerticalMergeShiftsWithSeparator:
    """Vertical 3-row merge (A5:A7) with blank separator — 1 extra row inserted."""

    def test_merge_shifts_to_correct_range(self, template_vertical_merge_path, tmp_path):
        wb, _ = _run_vertical_merge(template_vertical_merge_path, tmp_path)
        ws = wb.active
        assert "A6:A8" in _merge_ranges(ws), (
            f"Expected A6:A8 after 1-row shift, got {_merge_ranges(ws)}"
        )

    def test_no_stale_merge_at_original_position(self, template_vertical_merge_path, tmp_path):
        wb, _ = _run_vertical_merge(template_vertical_merge_path, tmp_path)
        ws = wb.active
        assert "A5:A7" not in _merge_ranges(ws), "Stale A5:A7 merge still present"

    def test_value_at_new_top_left(self, template_vertical_merge_path, tmp_path):
        wb, _ = _run_vertical_merge(template_vertical_merge_path, tmp_path)
        ws = wb.active
        assert ws["A6"].value == "Status"

    def test_italic_preserved(self, template_vertical_merge_path, tmp_path):
        wb, _ = _run_vertical_merge(template_vertical_merge_path, tmp_path)
        ws = wb.active
        assert ws["A6"].font.italic is True

    def test_data_rows_correct(self, template_vertical_merge_path, tmp_path):
        wb, _ = _run_vertical_merge(template_vertical_merge_path, tmp_path)
        ws = wb.active
        assert ws["A2"].value == "a"
        assert ws["A3"].value == "b"
        assert ws["A4"].value == "c"  # extra row inserted at row 4


class TestVerticalMergeMultipleInserts:
    """Vertical 3-row merge (A5:A7) with 3 extra rows — shifts by 3."""

    def test_merge_shifts_by_three(self, template_vertical_merge_path, tmp_path):
        wb, _ = _run_vertical_merge(
            template_vertical_merge_path, tmp_path, extra_indices=["c", "d", "e"]
        )
        ws = wb.active
        assert "A8:A10" in _merge_ranges(ws), (
            f"Expected A8:A10 after 3-row shift, got {_merge_ranges(ws)}"
        )

    def test_value_correct_after_multi_shift(self, template_vertical_merge_path, tmp_path):
        wb, _ = _run_vertical_merge(
            template_vertical_merge_path, tmp_path, extra_indices=["c", "d", "e"]
        )
        ws = wb.active
        assert ws["A8"].value == "Status"

    def test_italic_correct_after_multi_shift(self, template_vertical_merge_path, tmp_path):
        wb, _ = _run_vertical_merge(
            template_vertical_merge_path, tmp_path, extra_indices=["c", "d", "e"]
        )
        ws = wb.active
        assert ws["A8"].font.italic is True


# ---------------------------------------------------------------------------
# Vertical merge immediately adjacent (A4:A6, no blank separator)
# Template: rows 2-3 are data (a,b), rows 4-6 are the vertical merge "Status"
# DF has a, b, c — extra row inserted at row 4 (after row 3)
# Expected: merge shifts from A4:A6 → A5:A7; value and italic preserved;
#           extra row "c" appears at row 4 (the insertion slot)
# ---------------------------------------------------------------------------

class TestVerticalMergeAdjacentShifts:
    """Vertical 3-row merge (A4:A6) immediately adjacent to last data row."""

    def test_merge_shifts_to_correct_range(self, template_vertical_merge_adjacent_path, tmp_path):
        wb, _ = _run_vertical_merge(template_vertical_merge_adjacent_path, tmp_path)
        ws = wb.active
        assert "A5:A7" in _merge_ranges(ws), (
            f"Expected A5:A7 after 1-row shift, got {_merge_ranges(ws)}"
        )

    def test_no_stale_merge_at_original_position(self, template_vertical_merge_adjacent_path, tmp_path):
        wb, _ = _run_vertical_merge(template_vertical_merge_adjacent_path, tmp_path)
        ws = wb.active
        assert "A4:A6" not in _merge_ranges(ws), "Stale A4:A6 merge still present"

    def test_value_at_new_top_left(self, template_vertical_merge_adjacent_path, tmp_path):
        wb, _ = _run_vertical_merge(template_vertical_merge_adjacent_path, tmp_path)
        ws = wb.active
        assert ws["A5"].value == "Status"

    def test_italic_preserved(self, template_vertical_merge_adjacent_path, tmp_path):
        wb, _ = _run_vertical_merge(template_vertical_merge_adjacent_path, tmp_path)
        ws = wb.active
        assert ws["A5"].font.italic is True

    def test_extra_row_at_insertion_slot(self, template_vertical_merge_adjacent_path, tmp_path):
        wb, _ = _run_vertical_merge(template_vertical_merge_adjacent_path, tmp_path)
        ws = wb.active
        assert ws["A4"].value == "c", (
            "Extra row 'c' should appear at row 4 (the insertion point)"
        )

    def test_merge_not_split(self, template_vertical_merge_adjacent_path, tmp_path):
        """Merge must NOT be split into a 1-row cell + 2-row merge below it."""
        wb, _ = _run_vertical_merge(template_vertical_merge_adjacent_path, tmp_path)
        ws = wb.active
        ranges = _merge_ranges(ws)
        # The only merge involving cols A should be A5:A7 (3 rows intact)
        a_merges = [r for r in ranges if r.startswith("A")]
        assert a_merges == ["A5:A7"], f"Unexpected merge state: {a_merges}"

    def test_adjacent_multiple_inserts(self, template_vertical_merge_adjacent_path, tmp_path):
        """3 extra rows: merge shifts to A7:A9."""
        wb, _ = _run_vertical_merge(
            template_vertical_merge_adjacent_path, tmp_path, extra_indices=["c", "d", "e"]
        )
        ws = wb.active
        assert "A7:A9" in _merge_ranges(ws), (
            f"Expected A7:A9 after 3-row shift, got {_merge_ranges(ws)}"
        )
        assert ws["A7"].value == "Status"
        assert ws["A7"].font.italic is True


# ---------------------------------------------------------------------------
# Data-column vertical merge at boundary row
# Template: rows 2-3 data (a,b), row 4 has A4='Status' (not in DF) + B4:B6 merged
# DF has a, b, c — extra 'c' inserted at row 4, pushing row 4 down to row 5
# Expected:
#   - Merge B4:B6 shifts to B5:B7 (3-row, intact) with 'Section Header' + italic
#   - Extra 'c' at row 4 with correct DF values
#   - A5='Status' stays (unmatched key, left join behaviour)
# ---------------------------------------------------------------------------

def _run_data_col_merge(template_path, tmp_path, extra_indices=None):
    if extra_indices is None:
        extra_indices = ["c"]
    indices = ["a", "b"] + extra_indices
    df = pl.DataFrame({
        "Index": indices,
        "col1": list(range(len(indices))),
        "col2": list(range(100, 100 + len(indices))),
    })
    output = str(tmp_path / "output_dcm.xlsx")
    writer = ExcelTemplateWriter(template_path)
    writer.write({"data": TypedValue(df, "table")}, output)
    return load_workbook(output), output


class TestDataColVerticalMergeShifts:
    """3-row vertical merge in a DATA column (B4:B6) shifts intact after outer join."""

    def test_merge_shifts_to_correct_range(self, template_data_col_vertical_merge_path, tmp_path):
        wb, _ = _run_data_col_merge(template_data_col_vertical_merge_path, tmp_path)
        ws = wb.active
        assert "B5:B7" in _merge_ranges(ws), (
            f"Expected B5:B7 after 1-row shift, got {_merge_ranges(ws)}"
        )

    def test_no_stale_merge_at_original_position(self, template_data_col_vertical_merge_path, tmp_path):
        wb, _ = _run_data_col_merge(template_data_col_vertical_merge_path, tmp_path)
        ws = wb.active
        assert "B4:B6" not in _merge_ranges(ws), "Stale B4:B6 still present"

    def test_value_at_new_merge_top(self, template_data_col_vertical_merge_path, tmp_path):
        wb, _ = _run_data_col_merge(template_data_col_vertical_merge_path, tmp_path)
        ws = wb.active
        assert ws["B5"].value == "Section Header"

    def test_italic_preserved(self, template_data_col_vertical_merge_path, tmp_path):
        wb, _ = _run_data_col_merge(template_data_col_vertical_merge_path, tmp_path)
        ws = wb.active
        assert ws["B5"].font.italic is True

    def test_extra_row_data_correct(self, template_data_col_vertical_merge_path, tmp_path):
        wb, _ = _run_data_col_merge(template_data_col_vertical_merge_path, tmp_path)
        ws = wb.active
        # extra 'c' inserted at row 4
        assert ws["A4"].value == "c"
        assert ws["B4"].value == 2  # col1 index 2
        assert ws["C4"].value == 102  # col2

    def test_three_extra_rows_shift_by_three(self, template_data_col_vertical_merge_path, tmp_path):
        wb, _ = _run_data_col_merge(
            template_data_col_vertical_merge_path, tmp_path, extra_indices=["c", "d", "e"]
        )
        ws = wb.active
        assert "B7:B9" in _merge_ranges(ws), (
            f"Expected B7:B9 after 3-row shift, got {_merge_ranges(ws)}"
        )
        assert ws["B7"].value == "Section Header"
        assert ws["B7"].font.italic is True


# ---------------------------------------------------------------------------
# Positional fill() — no join column, data written by position
# ---------------------------------------------------------------------------

def _run_positional_fill(template_path, tmp_path, df, extra_df=None):
    """Fill template_positional_fill.xlsx with df via table(positional=True) tag.

    If extra_df is provided it is written under a second 'other' table tag.
    Returns the opened output workbook.
    """
    output = str(tmp_path / "out_fill.xlsx")
    writer = ExcelTemplateWriter(template_path)
    variables = {"data": TypedValue(df, "table")}
    if extra_df is not None:
        variables["other"] = TypedValue(extra_df, "table")
    writer.write(variables, output)
    return load_workbook(output)


class TestPositionalFill:
    """{{ data | table(positional=True) }} writes the DataFrame positionally from the tag cell."""

    def test_values_written_correctly(self, template_positional_fill_path, tmp_path):
        df = pl.DataFrame({"col1": [1, 2, 3], "col2": [4, 5, 6]})
        wb = _run_positional_fill(template_positional_fill_path, tmp_path, df)
        ws = wb.active
        # B3 = tag cell, written with df[0,0]
        assert ws["B3"].value == 1
        assert ws["C3"].value == 4
        assert ws["B4"].value == 2
        assert ws["C4"].value == 5
        assert ws["B5"].value == 3
        assert ws["C5"].value == 6

    def test_title_merge_untouched(self, template_positional_fill_path, tmp_path):
        df = pl.DataFrame({"col1": [1, 2, 3], "col2": [4, 5, 6]})
        wb = _run_positional_fill(template_positional_fill_path, tmp_path, df)
        ws = wb.active
        assert "A1:C1" in _merge_ranges(ws)
        assert ws["A1"].value == "My Report"
        assert ws["A1"].font.bold is True

    def test_footer_merge_untouched(self, template_positional_fill_path, tmp_path):
        df = pl.DataFrame({"col1": [1, 2, 3], "col2": [4, 5, 6]})
        wb = _run_positional_fill(template_positional_fill_path, tmp_path, df)
        ws = wb.active
        assert "A7:C7" in _merge_ranges(ws)
        assert ws["A7"].value == "Footer note"
        assert ws["A7"].font.italic is True

    def test_single_row_df(self, template_positional_fill_path, tmp_path):
        df = pl.DataFrame({"x": [99], "y": [88]})
        wb = _run_positional_fill(template_positional_fill_path, tmp_path, df)
        ws = wb.active
        assert ws["B3"].value == 99
        assert ws["C3"].value == 88

    def test_empty_df_raises(self, template_positional_fill_path, tmp_path):
        df = pl.DataFrame({"x": [], "y": []})
        writer = ExcelTemplateWriter(template_positional_fill_path)
        with pytest.raises(ValueError, match="non-empty"):
            writer.write({"data": TypedValue(df, "table")}, str(tmp_path / "err.xlsx"))


# ---------------------------------------------------------------------------
# Collision detection — overlapping fill regions raise ValueError
# ---------------------------------------------------------------------------

class TestCollisionDetection:
    """Overlapping table(positional=True) regions on the same sheet raise ValueError."""

    def test_two_overlapping_positional_tables_raise(self, template_collision_path, tmp_path):
        # first: B2:C4 (3 rows x 2 cols), second: C3:D4 (2 rows x 2 cols) — overlap at C3:C4
        first  = pl.DataFrame({"a": [1, 2, 3], "b": [4, 5, 6]})
        second = pl.DataFrame({"x": [7, 8], "y": [9, 10]})
        writer = ExcelTemplateWriter(template_collision_path)
        with pytest.raises(ValueError, match="collision"):
            writer.write(
                {
                    "first":  TypedValue(first,  "table"),
                    "second": TypedValue(second, "table"),
                },
                str(tmp_path / "err_collision.xlsx"),
            )

    def test_non_overlapping_positional_ok(self, template_positional_fill_path, tmp_path):
        # Only one table tag in this template — no collision possible
        df = pl.DataFrame({"col1": [1, 2], "col2": [3, 4]})
        writer = ExcelTemplateWriter(template_positional_fill_path)
        # Should NOT raise
        writer.write({"data": TypedValue(df, "table")}, str(tmp_path / "ok.xlsx"))


# ---------------------------------------------------------------------------
# Record (dot-notation) — {{ var_name.ColumnName }} for single-row DataFrames
#
# Template (template_record.xlsx, Sheet1):
#   B2: {{ result.Company }}
#   B3: {{ result.Revenue }}
#   B4: {{ other.Quarter }}
#   B5: {{ title }}            ← plain scalar
# ---------------------------------------------------------------------------

def _run_record_fill(template_path, tmp_path, result_df, other_df, title="Q1"):
    output = str(tmp_path / "out_record.xlsx")
    writer = ExcelTemplateWriter(template_path)
    writer.write(
        {
            "result": TypedValue(result_df, "record"),
            "other":  TypedValue(other_df,  "record"),
            "title":  TypedValue(title,      "single"),
        },
        output,
    )
    return load_workbook(output)


class TestRecordBasic:
    """Single-row DataFrames are accessed by column name via dot-notation."""

    def test_company_written(self, template_record_path, tmp_path):
        result = pl.DataFrame({"Company": ["Acme"], "Revenue": [999]})
        other  = pl.DataFrame({"Quarter": ["Q1"]})
        wb = _run_record_fill(template_record_path, tmp_path, result, other)
        ws = wb.active
        assert ws["B2"].value == "Acme"

    def test_revenue_written(self, template_record_path, tmp_path):
        result = pl.DataFrame({"Company": ["Acme"], "Revenue": [999]})
        other  = pl.DataFrame({"Quarter": ["Q1"]})
        wb = _run_record_fill(template_record_path, tmp_path, result, other)
        ws = wb.active
        assert ws["B3"].value == 999

    def test_other_namespace_written(self, template_record_path, tmp_path):
        result = pl.DataFrame({"Company": ["Acme"], "Revenue": [999]})
        other  = pl.DataFrame({"Quarter": ["Q2"]})
        wb = _run_record_fill(template_record_path, tmp_path, result, other)
        ws = wb.active
        assert ws["B4"].value == "Q2"

    def test_namespaces_independent(self, template_record_path, tmp_path):
        """Two record vars with the same column name resolve independently."""
        result = pl.DataFrame({"Company": ["Alpha"], "Revenue": [1]})
        other  = pl.DataFrame({"Quarter": ["Q3"]})
        wb = _run_record_fill(template_record_path, tmp_path, result, other)
        ws = wb.active
        assert ws["B2"].value == "Alpha"
        assert ws["B4"].value == "Q3"


class TestRecordMixedWithScalar:
    """Record vars and plain scalars coexist in the same template."""

    def test_plain_scalar_still_written(self, template_record_path, tmp_path):
        result = pl.DataFrame({"Company": ["Beta"], "Revenue": [42]})
        other  = pl.DataFrame({"Quarter": ["Q4"]})
        wb = _run_record_fill(template_record_path, tmp_path, result, other, title="Annual")
        ws = wb.active
        assert ws["B5"].value == "Annual"

    def test_all_cells_correct(self, template_record_path, tmp_path):
        result = pl.DataFrame({"Company": ["Gamma"], "Revenue": [7]})
        other  = pl.DataFrame({"Quarter": ["Q1"]})
        wb = _run_record_fill(template_record_path, tmp_path, result, other, title="Summary")
        ws = wb.active
        assert ws["B2"].value == "Gamma"
        assert ws["B3"].value == 7
        assert ws["B4"].value == "Q1"
        assert ws["B5"].value == "Summary"


class TestRecordMultiRowRaises:
    """Passing a DataFrame with more than one row raises ValueError."""

    def test_raises_on_two_rows(self, template_record_path, tmp_path):
        result = pl.DataFrame({"Company": ["A", "B"], "Revenue": [1, 2]})
        other  = pl.DataFrame({"Quarter": ["Q1"]})
        writer = ExcelTemplateWriter(template_record_path)
        with pytest.raises(ValueError, match="result"):
            writer.write(
                {
                    "result": TypedValue(result, "record"),
                    "other":  TypedValue(other,  "record"),
                    "title":  TypedValue("T",    "single"),
                },
                str(tmp_path / "err.xlsx"),
            )

    def test_error_message_contains_row_count(self, template_record_path, tmp_path):
        result = pl.DataFrame({"Company": ["A", "B", "C"], "Revenue": [1, 2, 3]})
        other  = pl.DataFrame({"Quarter": ["Q1"]})
        writer = ExcelTemplateWriter(template_record_path)
        with pytest.raises(ValueError, match="3"):
            writer.write(
                {
                    "result": TypedValue(result, "record"),
                    "other":  TypedValue(other,  "record"),
                    "title":  TypedValue("T",    "single"),
                },
                str(tmp_path / "err.xlsx"),
            )


# ---------------------------------------------------------------------------
# Sorted outer join — all rows (matched + inserted) sorted in the upper zone
#
# Fixtures share the same df helper: Index=[c,a,b,d], Value=[30,10,20,40].
# Templates have 2 upper zone rows (c, a); outer join adds b and d.
# ---------------------------------------------------------------------------

def _run_sorted_outer(template_path, tmp_path, df=None):
    """Fill a sorted-outer template with *df* (default: 4-row Index/Value DataFrame)."""
    if df is None:
        df = pl.DataFrame({
            "Index": ["c", "a", "b", "d"],
            "Value": [30, 10, 20, 40],
        })
    output = str(tmp_path / "out_sorted.xlsx")
    writer = ExcelTemplateWriter(template_path)
    writer.write({"data": TypedValue(df, "table")}, output)
    return load_workbook(output)


class TestSortedOuterAsc:
    """order_by=asc — all 4 df rows sorted ascending by join column (Index)."""

    def test_row_order_ascending(self, template_sorted_outer_asc_path, tmp_path):
        wb = _run_sorted_outer(template_sorted_outer_asc_path, tmp_path)
        ws = wb.active
        assert ws["A2"].value == "a"
        assert ws["A3"].value == "b"
        assert ws["A4"].value == "c"
        assert ws["A5"].value == "d"

    def test_values_match_sorted_rows(self, template_sorted_outer_asc_path, tmp_path):
        wb = _run_sorted_outer(template_sorted_outer_asc_path, tmp_path)
        ws = wb.active
        assert ws["B2"].value == 10  # a
        assert ws["B3"].value == 20  # b
        assert ws["B4"].value == 30  # c
        assert ws["B5"].value == 40  # d

    def test_end_table_row_deleted(self, template_sorted_outer_asc_path, tmp_path):
        """The {{ end_table }} marker row (Option A) must be deleted."""
        wb = _run_sorted_outer(template_sorted_outer_asc_path, tmp_path)
        ws = wb.active
        # Row 6 and beyond should be empty (end_table row deleted, nothing follows)
        assert ws["A6"].value is None
        assert ws["B6"].value is None


class TestSortedOuterDesc:
    """order_by=desc — all 4 df rows sorted descending by join column (Index)."""

    def test_row_order_descending(self, template_sorted_outer_desc_path, tmp_path):
        wb = _run_sorted_outer(template_sorted_outer_desc_path, tmp_path)
        ws = wb.active
        assert ws["A2"].value == "d"
        assert ws["A3"].value == "c"
        assert ws["A4"].value == "b"
        assert ws["A5"].value == "a"

    def test_values_match_sorted_rows(self, template_sorted_outer_desc_path, tmp_path):
        wb = _run_sorted_outer(template_sorted_outer_desc_path, tmp_path)
        ws = wb.active
        assert ws["B2"].value == 40  # d
        assert ws["B3"].value == 30  # c
        assert ws["B4"].value == 20  # b
        assert ws["B5"].value == 10  # a


class TestSortedOuterFixed:
    """order_by=asc with {{ insert_data }} marker — upper zone sorted, lower zone fixed."""

    def test_upper_zone_sorted_ascending(self, template_sorted_outer_fixed_path, tmp_path):
        wb = _run_sorted_outer(template_sorted_outer_fixed_path, tmp_path)
        ws = wb.active
        assert ws["A2"].value == "a"
        assert ws["A3"].value == "b"
        assert ws["A4"].value == "c"
        assert ws["A5"].value == "d"

    def test_upper_zone_values_correct(self, template_sorted_outer_fixed_path, tmp_path):
        wb = _run_sorted_outer(template_sorted_outer_fixed_path, tmp_path)
        ws = wb.active
        assert ws["B2"].value == 10
        assert ws["B3"].value == 20
        assert ws["B4"].value == 30
        assert ws["B5"].value == 40

    def test_lower_zone_fixed_key_preserved(self, template_sorted_outer_fixed_path, tmp_path):
        """Row 'total' is in the lower zone; its join column value must be retained."""
        wb = _run_sorted_outer(template_sorted_outer_fixed_path, tmp_path)
        ws = wb.active
        assert ws["A6"].value == "total"

    def test_lower_zone_value_empty_for_unmatched_key(self, template_sorted_outer_fixed_path, tmp_path):
        """'total' is not in the df so its Value column stays as-is (None from template)."""
        wb = _run_sorted_outer(template_sorted_outer_fixed_path, tmp_path)
        ws = wb.active
        assert ws["B6"].value is None


class TestSortedOuterByCol:
    """order_by=Value:desc — sorted by the Value column, not the join column."""

    def test_row_order_by_value_descending(self, template_sorted_outer_by_col_path, tmp_path):
        # df: Index=[c,a,b,d], Value=[30,10,20,40] → sorted by Value desc: d,c,b,a
        wb = _run_sorted_outer(template_sorted_outer_by_col_path, tmp_path)
        ws = wb.active
        assert ws["A2"].value == "d"  # Value=40
        assert ws["A3"].value == "c"  # Value=30
        assert ws["A4"].value == "b"  # Value=20
        assert ws["A5"].value == "a"  # Value=10

    def test_values_correct_after_col_sort(self, template_sorted_outer_by_col_path, tmp_path):
        wb = _run_sorted_outer(template_sorted_outer_by_col_path, tmp_path)
        ws = wb.active
        assert ws["B2"].value == 40
        assert ws["B3"].value == 30
        assert ws["B4"].value == 20
        assert ws["B5"].value == 10


class TestSortedOuterShorter:
    """df has 2 rows (a, d) but template has 3 upper zone slots (c, a, b).
    Remaining slots must be cleared after writing the sorted rows.
    """

    def test_rows_written_sorted(self, template_sorted_outer_shorter_path, tmp_path):
        df = pl.DataFrame({"Index": ["a", "d"], "Value": [10, 40]})
        wb = _run_sorted_outer(template_sorted_outer_shorter_path, tmp_path, df)
        ws = wb.active
        assert ws["A2"].value == "a"
        assert ws["B2"].value == 10
        assert ws["A3"].value == "d"
        assert ws["B3"].value == 40

    def test_leftover_slot_cleared(self, template_sorted_outer_shorter_path, tmp_path):
        """Row 4 was template slot 'b', unreachable from 2-row df — must be cleared."""
        df = pl.DataFrame({"Index": ["a", "d"], "Value": [10, 40]})
        wb = _run_sorted_outer(template_sorted_outer_shorter_path, tmp_path, df)
        ws = wb.active
        assert ws["A4"].value is None
        assert ws["B4"].value is None
