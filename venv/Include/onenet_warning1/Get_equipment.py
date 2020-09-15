

"""
    Author: Huang liuchao
    Contact: huanglc50@chinaunicom.cn
    Datetime: 2020/9/10 9:49
    Software: PyCharm
    File description：
        处理源数据E:\Pycharm_workspace\alarm_think\data\onenet_warning_data1\OSS2.0告警查询导出20200821-1.xls
        得出各个设备表（sbc、scscf、hss fe等）
"""
import hlc_common_utils as hcu
import onenet_warning_utils as owu
import os
import openpyxl
from openpyxl import load_workbook
import pandas as pd
from pathlib import Path
import win32com.client as win32
def get_equipment(file_path,sheet_name,equip_col_name):
    # 1、按照设备分表
    # 1.1存储各个设备的行数，方便导出表
    print('根据源数据得出各个设备表，并写入新的sheet..........')
    equipments_list = []  # 二维list value1：设备名  value2:行号
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    equipments_list = df[equip_col_name].values.tolist()  # 不含头标题

    unique_list = list(set(equipments_list))#得出所有设备名（不重复）

    unique_rows_values_list = []
    df_head_list = [df.columns.tolist()]
    for i in range(len(unique_list)):
        for j in range(len(equipments_list)):
            if equipments_list[j] == unique_list[i]:
                # 根据行将文件抽取出来存入新的sheet
                unique_rows_values_list.append(df.iloc[j])
        # 添加头标题

        unique_rows_values_list = df_head_list + unique_rows_values_list
        df1 = pd.DataFrame(unique_rows_values_list)
        excelWriter = pd.ExcelWriter(file_path, engine='openpyxl')
        hcu._excelAddSheet(df1, excelWriter, unique_list[i])
        print('%s设备数据抽出，并写入文件成功!!' % unique_list[i])
        unique_rows_values_list.clear()  # 清空列表为下一个循环
    print('\n所有设备数据全部抽出，并写入新的sheet。')


    # #1、先将文件xls转换成xlsx格式
    # if os.path.isfile(file_path) and file_path.endswith('.xls'):
    #     file_path1 = hcu.xls2xlsx_transformat_file(file_path)
    #     print(file_path)
    #     # os.remove(file_path)
    #     print('xls转换xlsx成功！')
    #     file_name+='x'
    # file_path=file_path1
    # if os.path.isfile(file_path):
    #     print('转换xlsx格式后的文件名:%s\n文件路径：%s'%(file_name,file_path))



    # print(len(unique_rows_values_list))
    # print(unique_rows_values_list)
    # print(len(unique_nrows_list))





    # print(len(equipments_list))
    # print(unique_list)
    # print(unique_list[0])
    # print(unique_nrows_list)


