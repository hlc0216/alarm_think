"""
    Author: Huang liuchao
    Contact: huanglc50@chinaunicom.cn
    Datetime: 2020/9/18 10:19
    Software: PyCharm
    File description：
        
"""
import hlc_common_utils as hcu
import onenet_warning_utils as owu
import os
import openpyxl
from openpyxl import load_workbook
import pandas as pd
from pathlib import Path
import win32com.client as win32
def Match_column_data(file_path,colum_name,to_file_path):
    """
    有两个sheet中存在相同的列名，列数据是包含关系，根据sheet中的列数据抽取sheet2中
    包含sheet1列数据的行数据
    :param file_path: 文件路径（包含sheet1,sheet2）
    :param colum_name: 两个sheet都包含的列名
    :param to_file_path: 要写入的文件路径
    """
    print('开始')
    file_path = file_path
    df1 = pd.read_excel(file_path, sheet_name=0)
    list1 = df1[colum_name].values.tolist()
    df2 = pd.read_excel(file_path, sheet_name=1)
    list2 = df2.values.tolist()
    list3 = []

    for i in range(len(list1)):
        for j in range(len(list2)):
            if list1[i] == list2[j][1]:
                list3.append(list2[j])

    to_file_path = to_file_path
    df3 = pd.DataFrame(list3)
    excelWriter = pd.ExcelWriter(file_path, engine='openpyxl')
    to_sheet_name = 'result'
    hcu._excelAddSheet(df3, excelWriter, to_sheet_name)
    print('匹配列数据完成，并抽出数据写入新的sheet(result)中！')
if __name__ == '__main__':
    print('开始')
    file_path=r'E:\Pycharm_workspace\alarm_think\venv\Include\test\2.xlsx'
    df1=pd.read_excel(file_path,sheet_name=0)
    list1= df1['设备'].values.tolist()
    df2=pd.read_excel(file_path,sheet_name=1,usecols=[0,1,2,3])
    list2=df2.values.tolist()
    list3 = []

    for i in range(len(list1)):
        for j in range(len(list2)):
            if list1[i]==list2[j][1]:
                list3.append(list2[j])

    to_file_path=r'E:\Pycharm_workspace\alarm_think\venv\Include\test\xlsx格式测试工作簿.xlsx'
    df3 = pd.DataFrame(list3)
    excelWriter = pd.ExcelWriter(file_path, engine='openpyxl')
    to_sheet_name = 'result'
    hcu._excelAddSheet(df3, excelWriter, to_sheet_name)
    print('写入成哦')



