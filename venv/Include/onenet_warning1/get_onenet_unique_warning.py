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
def get_onenet_unique_warning(dir_path,to_dir_path):
    files_name=os.listdir(dir_path)#得到所有文件的文件名
    for file_name in files_name:
        file_path=os.path.join(dir_path,file_name)
        print(file_path)
        get_onenet_unique_warning0(file_path,file_name,to_dir_path)
    print('全部单一告警统计完毕！')

def get_onenet_unique_warning0 (file_path,file_name,to_dir_path):
    # files_name=os.listdir(dir_path)#得到所有文件的文件名
    # file_path=os.path.join(dir_path,files_name)
    print('在设备-厂家-网元表的基础上，求出单一网元的唯一告警.........')
    # print(file_path)
    # for file in files_name:
    #读入文件
    df = pd.read_excel(file_path, sheet_name=None)
    sheets_name = list(df)
    # 获得表头，为后边添加表头使用
    # df_head = pd.read_excel(file_path, sheet_name=1)

    df_head_list = [['告警来源', '对象名称', '（告警）地市名称', '区县', '机房名称', '网元名称', '设备类型', '设备厂家', '网管告警级别', '告警标题', '厂家原始告警级别', '告警发生时间', '告警清除时间', '工单号', '派单所使用的规则', '派单状态', '未派单原因', '派单失败原因', '工单状态', '告警备注', '告警指纹fp0', '告警指纹fp1', '告警指纹fp2', '告警指纹fp3', '告警清除状态', '清除告警指纹fp0', '清除告警指纹fp1', '清除告警指纹fp2', '清除告警指纹fp3']]
    # print(df_head_list)
    #告警来源	对象名称	（告警）地市名称	区县		网元名称		设备厂家	网管告警级别	告警标题	厂家原始告警级别	告警发生时间	告警清除时间	工单号	派单所使用的规则	派单状态	未派单原因	派单失败原因	工单状态	告警备注	告警指纹fp0	告警指纹fp1	告警指纹fp2	告警指纹fp3	告警清除状态	清除告警指纹fp0	清除告警指纹fp1	清除告警指纹fp2	清除告警指纹fp3
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
        # if len(others_list)==0:
        #
        #     break
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

    to_file_name = '单一告警'+file_name
    to_file_path = os.path.join(to_dir_path,to_file_name)
    df_write_net.to_excel(to_file_path,index=None,header=None)

    print('remeber_nrows=', remeber_nrows)
    # print(nrows_values)
    print('单一告警写入完毕')
    # print('onenet_list的长度',len(onenet_list))
    # print('others_list的长度',len(others_list))




if __name__ == '__main__':
    #主函数测试用
    abspath = os.path.abspath('../../../data')  # 设置相对路径（基准路径）

    dir_path = abspath + r'\onenet_warning_data1\get_equipment_factory_netcell'
    to_dir_path = abspath + r'\onenet_warning_data1\get_onenet_unique_warning'
    get_onenet_unique_warning(dir_path,to_dir_path)


