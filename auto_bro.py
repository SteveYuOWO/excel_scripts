import os
import glob
import pandas as pd
from typing import List


def get_header(file_path: str) -> List[str]:
    _data = pd.read_excel(file_path, header=None)
    data = _data.iloc[:, :8]  # remove after col 8
    header = data.iloc[3].copy()
    header_list = header.to_list().append("XX")
    return header_list


def resolve_file(file_path: str) -> pd.DataFrame:
    print(f"Staring resolve {file_path}")
    _data = pd.read_excel(file_path, header=None)
    data = _data.iloc[:, :8]  # remove after col 8
    # date_time = data.iloc[0, 1] # date time row 0, col 1
    institution = data.iloc[1, 1]  # institution row 1, col 1
    value = data.iloc[4:].copy()
    value[8] = [institution for _ in range(len(value))]  # add institution to last column
    return value


def resolve_files(file_paths: List[str]) -> pd.DataFrame:
    data: pd.DataFrame = None
    for file_path in file_paths:
        if data is None:
            data = resolve_file(file_path)
        else:
            data = pd.concat([data, resolve_file(file_path)])
    return data


def format_df(df: pd.DataFrame) -> pd.DataFrame:
    ids = [i for i in range(1, len(df[0]) + 1)]  # from 1 to len
    df.index = ids
    df.iloc[:, 0] = ids  # column 0 add 1 to len
    return df


def format_writer(writer: pd.ExcelWriter, row: int, column: int) -> None:
    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]

    # header
    center_format = workbook.add_format({
        'align': 'center'
    })
    header_format = workbook.add_format({
        'bold': True,
        'font_color': 'black',
        'font_size': 20,
        'align': 'center'
    })
    header_format.set_align('vcenter')
    worksheet.set_row(row=0, height=14, cell_format=header_format)
    for i in range(1, row):
        worksheet.set_row(row=i,cell_format=center_format)

if __name__ == '__main__':
    base_dir = "/Users/steveyu/Documents/a.nosync/random"
    location = f"{base_dir}/*.xlsx"
    _file_paths = glob.glob(location)
    file_paths = list(filter(lambda path: "~$" not in path, _file_paths)) # filter ~$xx.xlsx
    print(f"Computing {len(file_paths)} files...")
    df = format_df(resolve_files(file_paths))
    writer = pd.ExcelWriter("output.xlsx", engine="xlsxwriter")
    df.to_excel(writer, index=False, sheet_name="Sheet1", header=[
        "序号", "客户名称",
        "新客/存量", "业务版块",
        "交流内容", "是否需要分行陪访",
        "是否需要总行参加", "客户经理",
        "未命名"])
    format_writer(writer, row=df.size, column=df.iloc[0].size)
    writer.save()