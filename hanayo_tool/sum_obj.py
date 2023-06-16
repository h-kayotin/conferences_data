"""
sum_mult - 合并文件，面向对象的优化

Author: hanayo
Date： 2023/6/14
"""

from openpyxl.worksheet.worksheet import Worksheet
import time
import openpyxl
from bs4 import BeautifulSoup
from pathlib import Path


class SumTool(object):
    def __init__(self):
        self.row_num = 1
        self.col_num = 1
        self.path = Path(input("请输入源数据文件夹路径："))
        self.files_list = list(self.path.glob("*.xls"))
        self.data_list = list()

    def sum_html(self, file):
        print(f"正在读取{file}\n", end="")
        with open(file, "rb") as file:
            html = file.read()
            bs_html = BeautifulSoup(html, "html.parser")
            tar_table = bs_html.find("table")
            rows = tar_table.find_all("tr")
            for row in rows:
                cols = row.find_all("td")
                if len(cols) > 0:
                    row_list = list()
                    for col in cols:
                        row_list.append(col.text)
                    self.data_list.append(row_list)
                    self.row_num += 1
                    if self.row_num % 1000 == 0:
                        print(f"已写入{self.row_num}行数据--->\n", end="")

    def save_file(self):
        start = time.time()
        for file in self.files_list:
            self.sum_html(file)
        workbook = openpyxl.Workbook()
        worksheet = workbook.worksheets[0]  # type:Worksheet
        for row in self.data_list:
            worksheet.append(row)
        workbook.save(f"{self.path}/summary_excel.xlsx")
        end = time.time()
        print(f"共耗时：{end - start:.2f}秒，合并了{len(self.data_list)}条数据。")


if __name__ == '__main__':
    # C:\Users\JiangHai江海\Desktop\工作\04.数据导出、合并\Oracle\tset

    SumTool().save_file()




