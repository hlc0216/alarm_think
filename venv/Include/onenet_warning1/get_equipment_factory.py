"""
    Author: Huang liuchao
    Contact: huanglc50@chinaunicom.cn
    Datetime: 2020/9/15 9:36
    Software: PyCharm
    File description：
        在设备表的基础上，将设备表根据厂家再细分为华为、中兴、爱立信.
        读取源文件，遍历sheet（设备），每个sheet按照厂家再分，处理完
        产生新的表，每个表有三个sheet（华为、中兴、爱立信）
"""
import hlc_common_utils as hcu
import onenet_warning_utils as owu
import os
import openpyxl
from openpyxl import load_workbook
import pandas as pd
from pathlib import Path
import win32com.client as win32
def get_equipment_factory (dir_path,file_name):
    print('在设备表的基础上，将设备表根据厂家再细分为华为、中兴、爱立信..............')
    file_path = os.path.join(dir_path, file_name)
    #读取源文件遍历sheet（设备）
    df = pd.read_excel(file_path,sheet_name=None)
    # df_head_list = [df.columns.tolist()]
    df_head = pd.read_excel(file_path,sheet_name=0)
    df_head_list = [df_head.columns.tolist()]

    print('表头为\n',df_head_list)
    sheets_name = list(df)
    #遍历sheet，按照厂家再分
    df_temp=pd.DataFrame()


    for i in range (1,len(sheets_name)):
        huawei_nrows_values = []  # 华为
        zte_nrows_values = []  # 中兴
        ericsson_nrows_values = []  # 爱立信
        # print(sheets_name[i])
        #读入文件
        df_temp=pd.read_excel(file_path,sheet_name=sheets_name[i])
        factory_col_values=df_temp['设备厂家'].values.tolist()
        for j in range (len(factory_col_values)):
            if factory_col_values[j]=='华为':
                huawei_nrows_values.append(df_temp.iloc[j])
            elif factory_col_values[j]=='中兴':
                zte_nrows_values.append(df_temp.iloc[j])
            elif factory_col_values[j]=='爱立信':
                ericsson_nrows_values.append(df_temp.iloc[j])

        #添加头标题
        huawei_nrows_values=df_head_list+huawei_nrows_values
        zte_nrows_values=df_head_list+zte_nrows_values
        ericsson_nrows_values=df_head_list+ericsson_nrows_values
        #创造文件
        df_write_huawei=pd.DataFrame(huawei_nrows_values)
        df_write_zte=pd.DataFrame(zte_nrows_values)
        df_write_ericsson=pd.DataFrame(ericsson_nrows_values)
        #写入文件
        to_dir_path=dir_path+r'\get_equipment_factory'
        to_file_name='设备'+sheets_name[i]+'.xlsx'
        to_file_path=os.path.join(to_dir_path, to_file_name)
        df_write_huawei.to_excel(to_file_path,sheet_name='华为',index=False,header=False)

        excelWriter = pd.ExcelWriter(to_file_path, engine='openpyxl')
        # hcu._excelAddSheet(df_write_huawei, excelWriter,'huawei')
        hcu._excelAddSheet(df_write_zte, excelWriter,'中兴')
        hcu._excelAddSheet(df_write_ericsson, excelWriter,'爱立信')

        print('设备 %s 按照厂家已经抽出，并写入文件成功！！'%sheets_name[i])

    print('\n所有设备已经按照厂家细分完毕！!')









