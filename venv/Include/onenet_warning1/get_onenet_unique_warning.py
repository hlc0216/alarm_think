"""
    Author: Huang liuchao
    Contact: huanglc50@chinaunicom.cn
    Datetime: 2020/9/16 15:55
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
def test_files(dir_path):
    files_name = os.listdir(dir_path)  # 得到所有文件的文件名
    print (files_name)
    # file_path = os.path.join(dir_path, files_name)
    print('在设备-厂家-网元表的基础上，求出单一网元的唯一告警.........')
    print(file_path)
def get_onenet_unique_warning (file_path,to_dir_path):
    # files_name=os.listdir(dir_path)#得到所有文件的文件名
    # file_path=os.path.join(dir_path,files_name)
    print('在设备-厂家-网元表的基础上，求出单一网元的唯一告警.........')
    # print(file_path)
    # for file in files_name:
    #读入文件
    df = pd.read_excel(file_path, sheet_name=None)
    sheets_name = list(df)
    # 获得表头，为后边添加表头使用
    df_head = pd.read_excel(file_path, sheet_name=1)
    df_head_list = [df_head.columns.tolist()]
    remeber_nrows=[]#记录保存的行
    nrows_values = []
    for i in range (1,len(sheets_name)):
        onenet_list = []
        others_list = []

        #获取onenet_list和others_list
        for j in range(1,len(sheets_name)):
            if(i==j):
                df_onenet = pd.read_excel(file_path, sheet_name=i)
                # print('获取onenet_list')
                # print(sheets_name[5])
                onenet_list=df_onenet['告警标题'].values.tolist()
                # print(onenet_list)
            else:
                df_others_net = pd.read_excel(file_path,sheet_name=j)
                # print('获取others_net_list')
                others_list += df_others_net['告警标题'].values.tolist()

        #对比onenet_list和others_list，求取行数
        nrows=[]
        for j in range (len(onenet_list)):
            for k in range (len(others_list)):
                if onenet_list[j] not in others_list:
                    nrows.append(j)
                    remeber_nrows.append(j)
        nrows=list(set(nrows))#得到唯一告警的行数
        remeber_nrows=list(set(remeber_nrows))


        #根据行数开始写文件到第一个sheet中
        if len(nrows)==0:
            print('%s网元不存在单一告警'%sheets_name[i])
        else:
            print('将%s网元存在单一告警并将单一告警追加写入第一个sheet中'%sheets_name[i])
            df_read_net = pd.read_excel(file_path, sheet_name=i)

            for n in range (len(nrows)):
                nrows_values.append(df_read_net.iloc[nrows[n]])
    nrows_values= df_head_list+nrows_values

    df_write_net = pd.DataFrame(nrows_values)
    to_file_name='中兴设备scscf单一告警.xlsx'
    to_file_path = os.path.join(to_dir_path,to_file_name)
    df_write_net.to_excel(to_file_path,index=None,header=None)
    # to_file_path = file_path
    # sheet_name='get_onenet_unique_warning'
    # excelWriter = pd.ExcelWriter(to_file_path, engine='openpyxl')
    # hcu._excelAddSheet(df_write_net, excelWriter, sheet_name)  # 为excel添加新的sheet，并将dataframe写入新的sheet（不覆盖其他sheet）
    # # print('%s网元处理数据完成，写入文件成功！' % file_name)







    print('remeber_nrows=', remeber_nrows)
    # print(nrows_values)
    print('单一告警写入完毕')
    # print('onenet_list的长度',len(onenet_list))
    # print('others_list的长度',len(others_list))









if __name__ == '__main__':
    #主函数测试用
    abspath = os.path.abspath('../../../data')  # 设置相对路径（基准路径）
    # 1、处理源数据得出各个设备表（sbc、scscf、hss fe等）
    dir_path = abspath + r'\onenet_warning_data1\get_equipment_factory_netcell'
    to_dir_path = abspath + r'\onenet_warning_data1\get_onenet_unique_warning'
    file_name = r'中兴设备scscf.xlsx'
    file_path = os.path.join(dir_path, file_name)
    # equip_col_name = '设备类型'
    get_onenet_unique_warning(file_path,to_dir_path)



