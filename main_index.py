"""
main_index - 程序首页

Author: JiangHai江海
Date： 2023/5/9
"""

import summary_files
import summary_fts


index_operations = {
    "1": summary_files.main,
    "2": summary_fts.main
}

if __name__ == '__main__':
    print("""
    请选择您要进行哪种操作：
    1：合并文件\n
    2：占位符\n
    3：占位符\n
    4：占位符\n
    """)
    op_type = input("请输入数字：")
    index_operations[op_type]()
