"""
summary_fts - 汇总代理商数据

Author: JiangHai江海
Date： 2023/4/25
"""

from pathlib import Path
import openpyxl
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter, column_index_from_string
import datetime
from openpyxl.cell import Cell
import xlrd
from bs4 import BeautifulSoup
import time


def get_files_from_folder(input_path):
    folder_src = input_path  # type: Path
    files = folder_src.glob("*.*")
    return list(files)


def read_config():
    # config_dic = {
    #     "大连康乐美": {
    #         "sheet_name": "商品(进、销、存)滚动表",
    #         "data_start_col": 3,
    #         "static_start": "A",
    #         "static_end": "H",
    #         "data_field": [
    #             {
    #                 "router": "",
    #                 "start_row": "",
    #                 "end_row": ""
    #             }
    #         ]
    #     }
    # }

    config_dic = {}
    con_wb = openpyxl.load_workbook("sources/config.xlsx")  # type: Workbook
    con_sheet = con_wb.worksheets[0]  # type: Worksheet
    for row in range(2, con_sheet.max_row + 1):
        config_dic[con_sheet.cell(row, 1).value] = {
            "sheet_name": con_sheet.cell(row, 2).value,
            "data_start_row": int(con_sheet.cell(row, 3).value),
            "static_start": con_sheet.cell(row, 4).value,
            "static_end": con_sheet.cell(row, 5).value,
            "data_list": []
        }
        for col in range(6, con_sheet.max_column + 1):
            if con_sheet.cell(row, col).value is None:
                continue
            cell_values = con_sheet.cell(row, col).value.split(",")
            data = {
                "router": cell_values[0],
                "start_row": cell_values[1],
                "end_row": cell_values[2]
            }
            config_dic[con_sheet.cell(row, 1).value]["data_list"].append(data)
    return config_dic


def read_xlsx(file_src, file_config, file_name):
    wb = openpyxl.load_workbook(file_src, data_only=True)  # type: Workbook
    sheet = wb[file_config["sheet_name"]]  # type: Worksheet

    start_row = file_config["data_start_row"]
    static_start = file_config["static_start"]
    static_end = file_config["static_end"]
    data_list = file_config["data_list"]

    processed_data_list = []
    for rout_index in range(len(data_list)):
        for row in range(start_row, sheet.max_row + 1):
            if sheet.cell(row, 1) == "合计":
                break
            # 如果销售额是0，跳过
            if sheet.cell(row, column_index_from_string(data_list[rout_index]["start_row"])).value == 0:
                continue

            processed_row = []
            for col in range(column_index_from_string(static_start), column_index_from_string(static_end)):
                processed_row.append(sheet.cell(row, col).value)
            processed_row.append(sheet.cell(row, column_index_from_string(data_list[rout_index]["start_row"])).value)
            processed_row.append(sheet.cell(row, column_index_from_string(data_list[rout_index]["end_row"])).value)
            processed_row.append(data_list[rout_index]["router"])
            processed_row.append(datetime.datetime.today().strftime("%Y年%m月%d日"))
            processed_row.append(file_name)
            processed_row.append("一般贸易")
            processed_row.append("ST(销)")
            processed_data_list.append(processed_row)

    return processed_data_list


def read_xls(file_src, file_config, file_name):
    pass


def main():
    config_file = read_config()  # 读取配置文件

    while True:
        source_src = Path(input("请输入文件夹路径："))
        #  C:\Users\JiangHai江海\Desktop\工作\04.数据导出\BI\source_BI
        files_list = get_files_from_folder(source_src)
        if source_src.exists():
            break
        else:
            print("路径有误，请重新输入--->\n")

    for file in files_list:
        file_name = Path(file).stem
        file_type = str(file).split(".")[-1]
        if file_type == "xlsx":
            read_xlsx(file, config_file[file_name], file_name)
        else:
            read_xls(file, config_file[file_name], file_name)


if __name__ == '__main__':
    config_file = read_config()
    res_list = read_xlsx("sources/上海路捷.xlsx", config_file["上海路捷"], "上海路捷")
    print(len(res_list))
    print(res_list[0])


