import os
from typing import Generator

import pandas as pd
from pandas import DataFrame, ExcelWriter

import analysis

FOLDER_ROOT_LOCATION = "/Users/xiwang/Downloads/PHASE ONE CODED"

SHEET_NAME = "Master"

MERGED_FILE_NAME = "POC.xlsx"


def get_all_excel_files(directory: str) -> Generator[str, None, None]:
    for dir_path, _, filenames in os.walk(directory):
        for f in filenames:
            if f.endswith(".xlsx"):
                yield os.path.abspath(os.path.join(dir_path, f))


def read_sheet(excel_file: str, sheet_name: str) -> DataFrame:
    return pd.read_excel(excel_file, sheet_name=sheet_name)


def write_sheet(data: DataFrame, path: str, sheet_name: str) -> None:
    with ExcelWriter(path, engine='xlsxwriter') as writer:
        data.to_excel(writer, sheet_name=sheet_name, index=False)


def to_column_string(n) -> str:
    """
    Convert column index to Excel column string.
    The column index should start from 0.
    """
    column_string = ""
    col_idx = n + 1
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        column_string = chr(65 + remainder) + column_string
    return column_string


if __name__ == '__main__':
    # Get all data
    all_data = DataFrame()
    for single_excel_file in get_all_excel_files(FOLDER_ROOT_LOCATION):
        print(f"Begin to process excel file -> {single_excel_file}")
        sheet_data = read_sheet(single_excel_file, SHEET_NAME)
        all_data = all_data.append(sheet_data)

    # Output to a file
    records_count = len(all_data.index)
    output_path = os.path.abspath(os.path.join(FOLDER_ROOT_LOCATION, "..", MERGED_FILE_NAME))
    print(f"\nOutput {records_count} records to file {output_path}")
    write_sheet(all_data, output_path, SHEET_NAME)

    # Create analysis sheet
    print("Write analysis result...")
    analysis.create_analysis_sheet(output_path, records_count)
