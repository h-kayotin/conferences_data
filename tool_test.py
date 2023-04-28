"""
test - 

Author: JiangHai江海
Date： 2023/4/28
"""

import hanayo_tool
from pathlib import Path

get_files_from_folder = hanayo_tool.hana_filename.get_files_from_folder
get_filenames = hanayo_tool.hana_filename.get_filenames

while True:
    source_src = Path(input("请输入文件夹路径："))
    files_list = get_files_from_folder(source_src)
    if source_src.exists():
        break
    else:
        print("路径有误，请重新输入--->\n")
print(files_list)
print(get_filenames(files_list))

