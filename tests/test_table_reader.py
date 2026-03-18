import pytest
from excel.table_reader import ExcelTableReader
from excel.exceptions import (
    ExcelTableReaderError,
    ExcelFileNotFoundError,
    ExcelCorruptedError,
    TableNotFoundError,
    MultipleTablesFoundError,
)


# ==================== extract_table_by_column_names ====================


class TestExtractByColumnNames:
    def test_returns_correct_columns_and_rows(self, simple_table_path):
        with ExcelTableReader(simple_table_path) as reader:
            df = reader.extract_table_by_column_names(["Name", "Amount", "Category"])
        assert list(df.columns) == ["Name", "Amount", "Category"]
        assert len(df) == 3

    def test_partial_columns_returns_full_table(self, simple_table_path):
        """A subset of column names still returns all columns of the table."""
        with ExcelTableReader(simple_table_path) as reader:
            df = reader.extract_table_by_column_names(["Name"])
        assert "Amount" in df.columns
        assert "Category" in df.columns
        assert len(df) == 3

    def test_unordered_columns_finds_table(self, simple_table_path):
        """Column order in required_columns does not need to match the sheet."""
        with ExcelTableReader(simple_table_path) as reader:
            df = reader.extract_table_by_column_names(["Category", "Name"])
        assert "Name" in df.columns
        assert "Amount" in df.columns

    def test_data_values(self, simple_table_path):
        with ExcelTableReader(simple_table_path) as reader:
            df = reader.extract_table_by_column_names(["Name", "Amount"])
        assert df["Name"].to_list() == ["Alice", "Bob", "Carol"]

    def test_first_sheet_returned_when_columns_in_multiple_sheets(
        self, multiple_sheets_path
    ):
        """When the same columns exist in multiple sheets, the first match is returned."""
        with ExcelTableReader(multiple_sheets_path) as reader:
            df = reader.extract_table_by_column_names(["Name", "Amount"])
        assert df["Name"].to_list() == ["Alice"]

    def test_multiple_tables_same_sheet_raises(self, multiple_tables_path):
        with ExcelTableReader(multiple_tables_path) as reader:
            with pytest.raises(MultipleTablesFoundError) as exc_info:
                reader.extract_table_by_column_names(["Name", "Amount"])
        assert len(exc_info.value.found_in) == 2

    def test_table_not_found_raises(self, simple_table_path):
        with ExcelTableReader(simple_table_path) as reader:
            with pytest.raises(TableNotFoundError):
                reader.extract_table_by_column_names(["NonExistent", "Column"])


# ==================== extract_table_by_column_names_from_sheet ====================


class TestExtractByColumnNamesFromSheet:
    def test_extracts_from_named_sheet(self, multiple_sheets_path):
        with ExcelTableReader(multiple_sheets_path) as reader:
            df = reader.extract_table_by_column_names_from_sheet(
                ["Name", "Amount"], sheet_name="Sheet2"
            )
        assert df["Name"].to_list() == ["Bob"]

    def test_table_not_found_in_sheet_raises(self, simple_table_path):
        with ExcelTableReader(simple_table_path) as reader:
            with pytest.raises(TableNotFoundError):
                reader.extract_table_by_column_names_from_sheet(
                    ["NonExistent"], sheet_name="Sheet1"
                )

    def test_multiple_matches_in_sheet_raises(self, multiple_tables_path):
        with ExcelTableReader(multiple_tables_path) as reader:
            with pytest.raises(MultipleTablesFoundError) as exc_info:
                reader.extract_table_by_column_names_from_sheet(
                    ["Name", "Amount"], sheet_name="Sheet1"
                )
        assert len(exc_info.value.found_in) == 2


# ==================== extract_table_by_range ====================


class TestExtractByRange:
    def test_exact_range_with_headers(self, simple_table_path):
        with ExcelTableReader(simple_table_path) as reader:
            df = reader.extract_table_by_range("A1:C4", sheet="Sheet1")
        assert df.shape == (3, 3)
        assert list(df.columns) == ["Name", "Amount", "Category"]

    def test_no_headers_with_column_names(self, no_headers_path):
        with ExcelTableReader(no_headers_path) as reader:
            df = reader.extract_table_by_range(
                "A1:C3",
                sheet="Sheet1",
                has_headers=False,
                column_names=["ID", "Name", "Value"],
            )
        assert list(df.columns) == ["ID", "Name", "Value"]
        assert len(df) == 3

    def test_no_headers_auto_names(self, no_headers_path):
        with ExcelTableReader(no_headers_path) as reader:
            df = reader.extract_table_by_range(
                "A1:C3", sheet="Sheet1", has_headers=False
            )
        assert list(df.columns) == ["col_0", "col_1", "col_2"]

    def test_single_data_row(self, simple_table_path):
        with ExcelTableReader(simple_table_path) as reader:
            df = reader.extract_table_by_range("A1:C2", sheet="Sheet1")
        assert df.shape == (1, 3)
        assert df["Name"].to_list() == ["Alice"]

    def test_sheet_not_found_raises(self, simple_table_path):
        with ExcelTableReader(simple_table_path) as reader:
            with pytest.raises(ExcelTableReaderError):
                reader.extract_table_by_range("A1:C5", sheet="DoesNotExist")


# ==================== extract_table_from_cell ====================


class TestExtractFromCell:
    def test_start_at_a1(self, simple_table_path):
        with ExcelTableReader(simple_table_path) as reader:
            df = reader.extract_table_from_cell("A1", sheet="Sheet1")
        assert list(df.columns) == ["Name", "Amount", "Category"]
        assert len(df) == 3

    def test_offset_table_start(self, offset_table_path):
        with ExcelTableReader(offset_table_path) as reader:
            df = reader.extract_table_from_cell("C3", sheet="Sheet1")
        assert list(df.columns) == ["Name", "Amount", "Category"]
        assert len(df) == 2

    def test_data_values(self, simple_table_path):
        with ExcelTableReader(simple_table_path) as reader:
            df = reader.extract_table_from_cell("A1", sheet="Sheet1")
        assert df["Name"].to_list() == ["Alice", "Bob", "Carol"]

    def test_sheet_not_found_raises(self, simple_table_path):
        with ExcelTableReader(simple_table_path) as reader:
            with pytest.raises(ExcelTableReaderError):
                reader.extract_table_from_cell("A1", sheet="DoesNotExist")


# ==================== extract_table_near_cell ====================


class TestExtractNearCell:
    def test_ref_cell_is_table_start(self, simple_table_path):
        with ExcelTableReader(simple_table_path) as reader:
            df = reader.extract_table_near_cell("A1", sheet="Sheet1")
        assert list(df.columns) == ["Name", "Amount", "Category"]
        assert len(df) == 3

    def test_scans_to_find_offset_table(self, offset_table_path):
        """Searching from A1 scans right and down until it finds the table at C3."""
        with ExcelTableReader(offset_table_path) as reader:
            df = reader.extract_table_near_cell("A1", sheet="Sheet1")
        assert "Name" in df.columns
        assert len(df) == 2

    def test_no_table_found_raises(self, empty_table_path):
        """Searching from a row below all content raises TableNotFoundError."""
        with ExcelTableReader(empty_table_path) as reader:
            with pytest.raises(TableNotFoundError):
                reader.extract_table_near_cell("A3", sheet="Sheet1")

    def test_sheet_not_found_raises(self, simple_table_path):
        with ExcelTableReader(simple_table_path) as reader:
            with pytest.raises(ExcelTableReaderError):
                reader.extract_table_near_cell("A1", sheet="DoesNotExist")


# ==================== Merged Cells ====================


class TestMergedCells:
    def test_unmerge_and_fill_removes_all_nulls(self, merged_cells_path):
        with ExcelTableReader(merged_cells_path) as reader:
            df = reader.extract_table_by_column_names(
                ["Region", "Country", "Sales"],
                unmerge_cells=True,
                fill_forward=True,
            )
        assert df["Region"].null_count() == 0
        assert df["Region"].to_list() == ["Europe", "Europe", "Asia", "Asia"]

    def test_no_unmerge_no_fill_leaves_nulls(self, merged_cells_path):
        with ExcelTableReader(merged_cells_path) as reader:
            df = reader.extract_table_by_column_names(
                ["Region", "Country", "Sales"],
                unmerge_cells=False,
                fill_forward=False,
            )
        assert df["Region"].null_count() > 0

    def test_fill_forward_without_unmerge_fills_via_polars(self, merged_cells_path):
        """fill_forward=True alone forward-fills nulls left by merged cells."""
        with ExcelTableReader(merged_cells_path) as reader:
            df = reader.extract_table_by_column_names(
                ["Region", "Country", "Sales"],
                unmerge_cells=False,
                fill_forward=True,
            )
        assert df["Region"].null_count() == 0
        assert df["Region"].to_list() == ["Europe", "Europe", "Asia", "Asia"]


# ==================== Edge Cases ====================


class TestEdgeCases:
    def test_empty_table_returns_empty_dataframe(self, empty_table_path):
        with ExcelTableReader(empty_table_path) as reader:
            df = reader.extract_table_by_column_names(["Name", "Amount"])
        assert len(df) == 0
        assert list(df.columns) == ["Name", "Amount"]


# ==================== File-level Errors ====================


class TestFileErrors:
    def test_context_manager_required(self):
        reader = ExcelTableReader("any.xlsx")
        with pytest.raises(ExcelTableReaderError):
            reader.extract_table_by_column_names(["Col"])

    def test_file_not_found(self):
        with pytest.raises(ExcelFileNotFoundError):
            with ExcelTableReader("does_not_exist.xlsx") as reader:
                reader.extract_table_by_column_names(["Col"])

    def test_corrupted_file(self, tmp_path):
        bad = tmp_path / "bad.xlsx"
        bad.write_text("not an excel file")
        with pytest.raises(ExcelCorruptedError):
            with ExcelTableReader(str(bad)) as reader:
                reader.extract_table_by_column_names(["Col"])
