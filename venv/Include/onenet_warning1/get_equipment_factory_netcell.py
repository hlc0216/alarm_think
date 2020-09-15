"""
    Author: Huang liuchao
    Contact: huanglc50@chinaunicom.cn
    Datetime: 2020/9/15 14:48
    Software: PyCharm
    File description：
        在设备-厂家表的基础上，再根据网元细分，得到每个厂家对应的
        网元。

"""
import hlc_common_utils as hcu
import onenet_warning_utils as owu
import os
import openpyxl
from openpyxl import load_workbook
import pandas as pd
from pathlib import Path
import win32com.client as win32

def get_files(dir_path):
    return os.listdir(dir_path)

def get_equipment_factory_netcell(dir_path,file_name):
    """
    某一个设备对三个厂家，根据网元进行细分，写入到新的sheet中
    :param dir_path:
    :param file_name:
    """
    print('在设备-厂家表的基础上，再根据网元细分，得到每个厂家对应的网元.........')
    #获取文件夹下的所有文件
    files_name = get_files(dir_path)
    #1、按照网元分表
    file_path= os.path.join(dir_path, file_name)
    # 读取源文件遍历sheet（厂家）
    df = pd.read_excel(file_path, sheet_name=None)

    # 获得表头，为后边添加表头使用
    df_head = pd.read_excel(file_path, sheet_name=0)
    df_head_list = [df_head.columns.tolist()]
    print('表头为\n', df_head_list)

    #得到所有的sheet
    sheets_name = list(df)
    # 遍历sheet，按照网元再分
    df1 = pd.DataFrame()
    print('所有的sheet',sheets_name)

    for i in range (len(sheets_name)-2):
        print(sheets_name[i])
        #读入文件
        df1 = pd.read_excel(file_path, sheet_name=sheets_name[i])
        netcell_values = df1['网元名称'].values.tolist()
        unique_netcell = list(set(netcell_values))  # 得出所有网元名（不重复）

        for j in range (len(unique_netcell)):
            unique_netcell_values = []
            for k in range (len(netcell_values)):
                if unique_netcell[j] == netcell_values[k]:
                    unique_netcell_values.append(df1.iloc[k])
            #添加头标题
            unique_netcell_values = df_head_list + unique_netcell_values
            df2 = pd.DataFrame(unique_netcell_values)
            #将某个设备-厂家-网元写入新的sheet

            to_file_path = r'E:\Pycharm_workspace\alarm_think\data\onenet_warning_data1\test2.xlsx'
            pd.DataFrame().to_excel(to_file_path)
            excelWriter = pd.ExcelWriter(to_file_path,engine='openpyxl')
            to_sheet_name = 'HSS FE'+sheets_name[i]+unique_netcell[j]
            hcu._excelAddSheet(df2, excelWriter, to_sheet_name)
            print('%s网元数据抽出，并写入文件成功!!' % to_sheet_name)
        print('\nHSS FE %s厂家的网元数据抽出完毕！'%sheets_name[i])


if __name__ == '__main__':
    dir_path=r'E:\Pycharm_workspace\alarm_think\data\onenet_warning_data1\get_equipment_factory'
    file_name='设备HSS FE.xlsx'
    file_path = os.path.join(dir_path, file_name)
    equip_factory_netcell_name='网元名称'
    get_equipment_factory_netcell(dir_path,file_name)

    # get_files(dir_path)
