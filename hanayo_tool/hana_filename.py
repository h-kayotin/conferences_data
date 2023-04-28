"""
hana_filename - 

Author: JiangHai江海
Date： 2023/4/28
"""
from pathlib import Path


def get_files_from_folder(input_path):
    folder_src = input_path  # type: Path
    files = folder_src.glob("*.*")
    return list(files)


def get_filenames(files):
    file_names = []
    for file in files:
        file_names.append(Path(file).stem)
    return file_names


if __name__ == '__main__':
    while True:
        source_src = Path(input("请输入文件夹路径："))
        files_list = get_files_from_folder(source_src)
        if source_src.exists():
            break
        else:
            print("路径有误，请重新输入--->\n")
    print(files_list)
    print(get_filenames(files_list))