"""
summary_files - 汇总数据到一张表里

Author: JiangHai江海
Date： 2023/4/6
"""
from pathlib import Path
import openpyxl


def sum_csv():
    pass


def sum_xls():
    pass


def sum_xlsx(list):
    for xlsx_file in list:
        pass


def get_files_from_folder(input_path):
    folder_src = input_path  # type: Path
    files = folder_src.glob("*.*")
    return list(files)


file_type ={
    "1": sum_csv,
    "2": sum_xls,
    "3": sum_xlsx
}

if __name__ == '__main__':
    while True:
        source_src = input("请输入源数据文件夹路径：")
        source_src = Path(source_src)
        files_list = get_files_from_folder(source_src)
        if source_src.exists():
            break
        else:
            print("路径有误，请重新输入--->\n")
