"""
summary_fts - 汇总代理商数据

Author: JiangHai江海
Date： 2023/4/25
"""

from pathlib import Path
import openpyxl
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell import Cell
import xlrd
from bs4 import BeautifulSoup
import time


def get_files_from_folder(input_path):
    folder_src = input_path  # type: Path
    files = folder_src.glob("*.*")
    return list(files)


def read_config():
    pass


def main():
    while True:
        source_src = Path(input("请输入文件夹路径："))
        #  C:\Users\JiangHai江海\Desktop\工作\04.数据导出\BI\source_BI
        files_list = get_files_from_folder(source_src)
        if source_src.exists():
            break
        else:
            print("路径有误，请重新输入--->\n")
    for file in files_list:
        print(file)
        file_type = str(file).split(".")[-1]
        if file_type == "xlsx":
            pass
        else:
            pass


if __name__ == '__main__':
    main()
