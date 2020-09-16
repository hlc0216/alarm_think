import pandas as pd
import xlrd
import xlwt
import numpy as np
import openpyxl




def append_excel_xlsx(file_path,sheet_name,values):
    wb = openpyxl.load_workbook(file_path)
    sheetnames = wb.sheetnames
    print(sheetnames)
    # table = data.get_sheet_by_name(sheet_name)
    table = wb[sheet_name]

    table = wb.active
    print(table.title)  # 输出表名
    nrows = table.max_row  # 获得行数
    ncolumns = table.max_column  # 获得列数

    # 注意行业列下标是从1开始的
    for i in range(1, len(values) + 1):
        for j in range(1, len(values[i - 1]) + 1):
            table.cell(nrows + i, j).value = values[i - 1][j - 1]

    wb.save(file_path)
    print("xlsx格式表格追加写入数据成功！")

def write_excel_xlsx(path, sheet_name, value):
    index = len(value)
    workbook = openpyxl.Workbook()  # 新建工作簿（默认有一个sheet？）
    sheet = workbook.active  # 获得当前活跃的工作页，默认为第一个工作页
    sheet.title = sheet_name  # 给sheet页的title赋值
    for i in range(0, index):
        for j in range(0, len(value[i])):
            sheet.cell(row=i + 1, column=j + 1, value=str(value[i][j]))  # 行，列，值 这里是从1开始计数的
    workbook.save(path)  # 一定要保存
    print("xlsx格式表格写入数据成功！")

if __name__ == '__main__':

    book_name_xlsx = r'E:\Pycharm_workspace\alarm_think\venv\Include\test\xlsx格式测试工作簿.xlsx'

    sheet_name_xlsx = 'xlsx格式测试表'

    value3 = [["姓名", "性别", "年龄", "城市", "职业"],
              ["111", "女", "66", "石家庄", "运维工程师"],
              ["222", "男", "55", "南京", "饭店老板"],
              ["333", "女", "27", "苏州", "保安"], ]
    value4 = [["姓名", "性别", "年龄", "城市", "职业"],
              ["444", "女", "66", "石家庄", "运维工程师"],
              ["555", "男", "55", "南京", "饭店老板"],
              ["666", "女", "27", "苏州", "保安"], ]
    values5 = [['E', 'X', 'C', 'E', 'L'],
              [1, 2, 3, 4, 5],
              ['a', 'b', 'c', 'd', 'e']]

    # write_excel_xlsx(book_name_xlsx, sheet_name_xlsx, value4)
    append_excel_xlsx(book_name_xlsx,'Sheet1',value3)
