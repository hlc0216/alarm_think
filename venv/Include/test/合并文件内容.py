"""
    Author: Huang liuchao
    Contact: huanglc50@chinaunicom.cn
    Datetime: 2020/9/21 9:16
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
def File_content_merge(file_path1,file_path2,to_file_path):
    """
    实现将两个（列名相同的）excel文件的内容合并到一个excel中
    :param file_path1: 文件1路径
    :param file_path2: 文件2路径
    :param to_file_path: 要写入的文件路径
    """
    print('实现将两个（列名相同的）excel文件的内容合并到一个excel中')
    file_path1 = file_path1
    file_path2 = file_path2
    df1 = pd.read_excel(file_path1, sheet_name=0)
    df1_head_list = [df1.columns.tolist()]
    df2 = pd.read_excel(file_path2, sheet_name=0)
    df2_head_list = [df2.columns.tolist()]
    df3 = pd.DataFrame()
    sumlist =  df1_head_list + df1.values.tolist()+df2_head_list + df2.values.tolist()
    to_file_path = to_file_path
    pd.DataFrame(sumlist).to_excel(to_file_path, header=None, index=None)
    print('两个excel文件文本合并完成！')
if __name__ == '__main__':
    file_path1=r'E:\Pycharm_workspace\alarm_think\venv\Include\test\2.xlsx'
    file_path2=r'E:\Pycharm_workspace\alarm_think\venv\Include\test\1.xlsx'
    to_file_path=r'E:\Pycharm_workspace\alarm_think\venv\Include\test\3.xlsx'
    File_content_merge(file_path1,file_path2,to_file_path)




