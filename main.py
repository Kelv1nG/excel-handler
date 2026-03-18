from excel.data_extractor import ExcelDataExtractor


if __name__ == "__main__":
    with ExcelDataExtractor("sample.xlsx") as x:
        df = x.extract_table_by_column_names_from_sheet(["A", "C", "D", "E", "F"], "Sheet1", unmerge_cells=False, fill_forward=False)
        print(df)

        df = x.extract_table_by_range("C16:G21", "Sheet1", unmerge_cells=False, fill_forward=False)
        print(df)