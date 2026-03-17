from excel.data_extractor import ExcelDataExtractor


if __name__ == "__main__":
    with ExcelDataExtractor("sample.xlsx") as x:
        df = x.extract_table_by_range("G9:K13", sheet="Sheet1")
        print(df)
