# kayotin_excel

### 功能描述

批量合并文件，目前支持csv,xlsx,xls,html
公司经常有这种蠢工作，需要把多个Excel文件合并成一个，

希望使用该工具可以节省大家的时间。

### 注意事项
html文件目前只针对里面仅有一个table标签， 不太有泛用性；

其他可以通过参数输入开始的行和列。


### 读取方式
数据有三种类型，分别是xlsx/csv/html

所以分别用了：
- openpyxl来读取xlsx
- xlrd来读取xls
- csv来读取csv
- BeautifulSoup来读取html文件

一对比就发现openpyxl的效率是最低的~可能他的优点就是比较直观吧

### How to use
运行summary_file.py,根据提示输入就行了，就可以合并数据了







