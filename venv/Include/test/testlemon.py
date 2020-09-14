import pandas as pd
import xlrd
import xlwt
import openpyxl
import pandas as pd
import xlrd
import os
import xlwt
import openpyxl
import report_huawei
from openpyxl import load_workbook

def excel_to_set(file_path,sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df_li =  df.values.tolist()
    print (df_li)

#添加sheet
def add_sheet(sheet_name):
    wb = wb = openpyxl.load_workbook(r'E:\Pycharm_workspace\lemon_copy.xlsx')
    wb.create_sheet(title=sheet_name, index=None)
    wb.save(r'E:\Pycharm_workspace\lemon.xlsx')
    count=1



#添加列
def add_col2():

    df = pd.read_excel(r'E:\Pycharm_workspace\lemon_copy.xlsx')
    total_row = len(df)
    add_col_values=[]
    print('总共包含%d行', total_row)
    for i in range(total_row):
        add_col_values.append('华为')
    df['厂家']=add_col_values

    #写入excel指定sheet
    writer = pd.ExcelWriter(r'E:\Pycharm_workspace\lemon_copy.xlsx')
    # df.to_excel(r'E:\Pycharm_workspace\lemon_copy.xlsx',sheet_name='add_col',index=None)
    newsheet=add_sheet('newsheet2')
    df.to_excel(writer,'newsheet2')

    # df.to_excel(r'E:\Pycharm_workspace\lemon_copy.xlsx',sheet_name='newSheet',index=None)

    print('写入文件成功')
def add_col():
    df=pd.read_excel(r'E:\Pycharm_workspace\lemon.xlsx')
    #创建列名和列数据
 #处理表头
    add_col_name='厂家'
    df_head_li=df.columns.tolist()  #获得表头list
    # df_head_li[df.shape[1]+1]=add_col_name
    df_head_li.append(add_col_name)   #在原先表头的基础上添加一个元素
    print('表头信息为：%s',df_head_li)

#处理表头对应的内容
    #为新添加的表头添加内容

    df_value_li = df_head_li + df.values.tolist()
    add_col_values=[]
    total_row=len(df)
    print('总共包含%d行',total_row)
    for i in range(total_row):
        add_col_values.append('华为')
    # print(add_col_values)
    df[add_col_name]=add_col_values
    df.to_excel(r'E:\Pycharm_workspace\华为HSS U2000历史告警查询结果_20200821.xls',sheet_name='test', columns=[add_col_name] ,index=False, header=0)
if __name__ == '__main__':
    file_path=r"E:\Pycharm_workspace\data\lemon.xlsx"
    sheet_name=hlc_sheet

    # add_col2()
    excel_to_set(file_path,sheet_name)