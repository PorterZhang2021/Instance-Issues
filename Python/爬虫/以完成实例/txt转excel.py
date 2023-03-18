"""
从txt中获取需要的json格式数据转换成excel保存
"""
from openpyxl import Workbook

# 文件地址
file_path = r'E:\Project\Instance Issues\Python\爬虫\实例1\book.txt'
# 存放地址
file_save_path = r'E:\Project\Instance Issues\Python\爬虫\实例1\当当Top500.xlsx'
# 打开文件
book_file = open(file_path, 'r', encoding='utf-8')
# 行列表
line_list = book_file.readlines()
# 行字典列表
line_dict_list = []
# 将元素转换成字典列表
for line in line_list:
    # 转换成list
    line = list(line)
    # 将末尾的元素'\n'去除
    line = line[:-1]
    # 重新转换成字符串
    line = ''.join(line)
    # 转换成字典
    line = eval(line)
    # 将每行加入字典中
    line_dict_list.append(line)
# 图书字典
book_dict = {}
# 行元素
rows = []
# 键
key_list = []
# 创建图书字典
for key in line_dict_list[0].keys():
    key_list.append(key)

# 加入行元素中
rows.append(key_list)
# 图书字典元素存入
for line_dict in line_dict_list:
    # 行元素
    row = []
    for key in key_list:
        row.append(line_dict[key])
    # 将元素加入行中
    rows.append(row)
            

# 将元素存入
wb = Workbook()
sheet = wb.active
sheet.title = '当当Top500'

for row in rows:
     sheet.append(row)

wb.save(file_save_path)





