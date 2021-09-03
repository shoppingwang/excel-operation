import os
from typing import Generator

import pandas as pd
from pandas import DataFrame, ExcelWriter

SHEET_NAME = "Master"

MERGED_FILE_NAME = "POC.xlsx"


def _get_all_excel_files(directory: str) -> Generator[str, None, None]:
    for dir_path, _, filenames in os.walk(directory):
        for f in filenames:
            if f.endswith(".xlsx"):
                yield os.path.abspath(os.path.join(dir_path, f))


def _read_sheet(excel_file: str, sheet_name: str) -> DataFrame:
    return pd.read_excel(excel_file, sheet_name=sheet_name)


def _write_sheet(data: DataFrame, path: str, sheet_name: str) -> None:
    with ExcelWriter(path, engine='xlsxwriter') as writer:
        data.to_excel(writer, sheet_name=sheet_name, index=False)


def merge_excel_sheet(base_path: str) -> (str, int):
    # Get all data
    all_data = DataFrame()
    for single_excel_file in _get_all_excel_files(base_path):
        print(f"Begin to process excel file -> {single_excel_file}")
        sheet_data = _read_sheet(single_excel_file, SHEET_NAME)
        all_data = all_data.append(sheet_data)

    # Output to a file
    records_count = len(all_data.index)
    output_path = os.path.abspath(os.path.join(base_path, "..", MERGED_FILE_NAME))
    print(f"\nOutput {records_count} records to file {output_path}")
    _write_sheet(all_data, output_path, SHEET_NAME)

    return output_path, records_count
