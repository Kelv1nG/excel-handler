import pytest
from excel.cell_reader import ExcelCellReader
from excel.exceptions import (
    ExcelError,
    ExcelSheetNotFoundError,
    KeywordNotFoundError,
    MultipleKeywordsFoundError,
)


# ---------------------------------------------------------------------------
# get / get_many (existing behaviour — regression guard)
# ---------------------------------------------------------------------------

class TestGetAndGetMany:
    def test_get_known_cell(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            assert r.get("Sheet1!B1") == 5000

    def test_get_active_sheet(self, anchored_cells_path):
        # Sheet1 is active (first sheet); A1 = "Revenue Label"
        with ExcelCellReader(anchored_cells_path) as r:
            assert r.get("A1") == "Revenue Label"

    def test_get_many_returns_dict(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            result = r.get_many(["Sheet1!A1", "Sheet1!B1"])
        assert result == {"Sheet1!A1": "Revenue Label", "Sheet1!B1": 5000}

    def test_get_unknown_sheet_raises(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            with pytest.raises(ExcelSheetNotFoundError, match="NoSheet"):
                r.get("NoSheet!A1")

    def test_context_manager_required(self, anchored_cells_path):
        r = ExcelCellReader(anchored_cells_path)
        with pytest.raises(ExcelError):
            r.get("A1")


# ---------------------------------------------------------------------------
# get_relative — cell_ref anchor
# ---------------------------------------------------------------------------

class TestGetRelativeCellRef:
    def test_right_offset(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            assert r.get_relative(cell_ref="Sheet1!A1", right=1) == 5000

    def test_right_offset_two_steps(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            assert r.get_relative(cell_ref="Sheet1!A1", right=2) == "USD"

    def test_left_offset(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            assert r.get_relative(cell_ref="Sheet1!B1", left=1) == "Revenue Label"

    def test_down_offset(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            assert r.get_relative(cell_ref="Sheet1!A1", down=1) == "Tax Label"

    def test_up_offset(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            assert r.get_relative(cell_ref="Sheet1!A2", up=1) == "Revenue Label"

    def test_combined_offset(self, anchored_cells_path):
        # A1 + down=1, right=1 → B2 = 250
        with ExcelCellReader(anchored_cells_path) as r:
            assert r.get_relative(cell_ref="Sheet1!A1", down=1, right=1) == 250

    def test_zero_offset_returns_anchor_value(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            assert r.get_relative(cell_ref="Sheet1!A1") == "Revenue Label"


# ---------------------------------------------------------------------------
# get_relative — keyword anchor
# ---------------------------------------------------------------------------

class TestGetRelativeKeyword:
    def test_keyword_right(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            assert r.get_relative(keyword="Revenue Label", sheet="Sheet1", right=1) == 5000

    def test_keyword_case_insensitive(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            assert r.get_relative(keyword="revenue label", sheet="Sheet1", right=1) == 5000

    def test_keyword_with_extra_spaces(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            assert r.get_relative(keyword="  Revenue Label  ", sheet="Sheet1", right=1) == 5000

    def test_keyword_down(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            assert r.get_relative(keyword="Revenue Label", sheet="Sheet1", down=1) == "Tax Label"

    def test_keyword_different_label(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            assert r.get_relative(keyword="Tax Label", sheet="Sheet1", right=1) == 250

    def test_keyword_cross_sheet_search(self, anchored_cells_path):
        # "Tax Label" only exists in Sheet1 — no sheet= needed
        with ExcelCellReader(anchored_cells_path) as r:
            assert r.get_relative(keyword="Tax Label", right=1) == 250

    def test_keyword_on_sheet2(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            assert r.get_relative(keyword="Revenue Label", sheet="Sheet2", right=1) == 9999

    def test_keyword_not_found_raises(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            with pytest.raises(KeywordNotFoundError):
                r.get_relative(keyword="No Such Label", sheet="Sheet1")

    def test_keyword_not_found_any_sheet_raises(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            with pytest.raises(KeywordNotFoundError):
                r.get_relative(keyword="Completely Missing")

    def test_multiple_keywords_raises(self, anchored_cells_path):
        # "Revenue Label" appears in both Sheet1 and Sheet2
        with ExcelCellReader(anchored_cells_path) as r:
            with pytest.raises(MultipleKeywordsFoundError) as exc_info:
                r.get_relative(keyword="Revenue Label")
        assert len(exc_info.value.found_in) == 2

    def test_multiple_keywords_found_in_contains_locations(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            with pytest.raises(MultipleKeywordsFoundError) as exc_info:
                r.get_relative(keyword="Revenue Label")
        locations = exc_info.value.found_in
        assert any("Sheet1" in loc for loc in locations)
        assert any("Sheet2" in loc for loc in locations)

    def test_unknown_sheet_arg_raises(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            with pytest.raises(ExcelSheetNotFoundError, match="NoSheet"):
                r.get_relative(keyword="Revenue Label", sheet="NoSheet")


# ---------------------------------------------------------------------------
# get_relative — argument validation
# ---------------------------------------------------------------------------

class TestGetRelativeValidation:
    def test_neither_arg_raises(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            with pytest.raises(ValueError, match="Specify either"):
                r.get_relative()

    def test_both_args_raise(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            with pytest.raises(ValueError, match="Cannot specify both"):
                r.get_relative(cell_ref="Sheet1!A1", keyword="Revenue Label")


# ---------------------------------------------------------------------------
# get_many_relative — cell_ref anchor
# ---------------------------------------------------------------------------

class TestGetManyRelativeCellRef:
    def test_multiple_offsets(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            result = r.get_many_relative(
                cell_ref="Sheet1!A1",
                offsets={
                    "value": {"right": 1},
                    "currency": {"right": 2},
                    "tax_label": {"down": 1},
                    "tax_value": {"down": 1, "right": 1},
                },
            )
        assert result == {
            "value": 5000,
            "currency": "USD",
            "tax_label": "Tax Label",
            "tax_value": 250,
        }

    def test_empty_offsets(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            result = r.get_many_relative(cell_ref="Sheet1!A1", offsets={})
        assert result == {}

    def test_none_offsets(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            result = r.get_many_relative(cell_ref="Sheet1!A1")
        assert result == {}


# ---------------------------------------------------------------------------
# get_many_relative — keyword anchor
# ---------------------------------------------------------------------------

class TestGetManyRelativeKeyword:
    def test_keyword_multiple_offsets(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            result = r.get_many_relative(
                keyword="Tax Label",
                sheet="Sheet1",
                offsets={
                    "value": {"right": 1},
                    "label_above": {"up": 1},
                },
            )
        assert result == {"value": 250, "label_above": "Revenue Label"}

    def test_keyword_not_found_raises(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            with pytest.raises(KeywordNotFoundError):
                r.get_many_relative(
                    keyword="Ghost",
                    offsets={"x": {"right": 1}},
                )

    def test_multiple_keywords_raises(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            with pytest.raises(MultipleKeywordsFoundError):
                r.get_many_relative(
                    keyword="Revenue Label",
                    offsets={"v": {"right": 1}},
                )


# ---------------------------------------------------------------------------
# get_many_relative — argument validation
# ---------------------------------------------------------------------------

class TestGetManyRelativeValidation:
    def test_neither_arg_raises(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            with pytest.raises(ValueError, match="Specify either"):
                r.get_many_relative(offsets={"x": {"right": 1}})

    def test_both_args_raise(self, anchored_cells_path):
        with ExcelCellReader(anchored_cells_path) as r:
            with pytest.raises(ValueError, match="Cannot specify both"):
                r.get_many_relative(
                    cell_ref="Sheet1!A1",
                    keyword="Revenue Label",
                    offsets={"x": {"right": 1}},
                )
