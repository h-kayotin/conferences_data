"""
sum_mult - 合并文件的多线程版本

Author: hanayo
Date： 2023/6/14
"""

from concurrent.futures import ThreadPoolExecutor
from threading import RLock
from openpyxl.worksheet.worksheet import Worksheet
import time
import openpyxl
from bs4 import BeautifulSoup
from pathlib import Path


class SumTool(object):
    def __init__(self):
        self.row_num = 1
        self.path = Path(input("请输入源数据文件夹路径："))
        self.files_list = list(self.path.glob("*"))
        self.data_list = []
        self.lock = RLock()

    def sum_csv(self, files_list, start_num):
        pass

    def sum_xls(self, files_list, start_num):
        pass

    def sum_html(self, file):
        print(f"正在读取{file}\n", end="")
        self.lock.acquire()
        with open(file, "rb") as file:
            bs_html = BeautifulSoup(file.read(), "html/parser")
            tar_table = bs_html.find("table")
            rows = tar_table.find_all("tr")
            for row in rows:
                cols = row.find_all("td")
                if len(cols) > 0:
                    self.data_list.append(row)
                    self.row_num += 1
                    if self.row_num % 1000 == 0:
                        print(f"已写入{self.row_num}行数据--->\n", end="")
        self.lock.release()

    def save_file(self):
        with ThreadPoolExecutor(max_workers=4) as pool:
            for file in self.files_list:
                pool.submit(self.sum_html, file=file)
        workbook = openpyxl.Workbook()
        worksheet = workbook.worksheets[0]  # type:Worksheet
        for row in self.data_list:
            worksheet.append(row)
        workbook.save("summary_excel.xlsx")


if __name__ == '__main__':
    # C:\Users\JiangHai江海\Desktop\工作\04.数据导出、合并\Oracle\3月

    SumTool().save_file()




