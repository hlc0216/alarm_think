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
def test_mkdir():
    """
    测试在指定文件夹下，建立新的文件夹

    """
    dir_path=r'H:\pycharm_workspace\alarm_think\data\onenet_warning_data1\get_equipment_factory_netcell'

    to_dir_path = dir_path + '\\' + '华为'
    os.mkdir(to_dir_path)
    to_file_name = 'test.xlsx'
    to_file_path = os.path.join(to_dir_path, to_file_name)
    pd.DataFrame().to_excel(to_file_path)
def test_str():
    """
    字符串测试
    """
    a='设备Gomc.xlsx'
    print(a[2:-5])
    print('\033[32;1m' + a[2:-5] + '\033[0m')
def get_files(dir_path):
    """
    获得文件夹下的所有文件列表
    :param dir_path:
    :return: 返回文件列表
    """
    return os.listdir(dir_path)


def get_equipment_factory_netcell(dir_path,to_dir_path):
    """
    所有设备对三个厂家，根据网元进行细分，写入到新的sheet中
    :param to_dir_path: 处理完的数据要输入的目录
    :param dir_path: 要处理文件所在的目录
    :param file_name: 要处理的文件名
    """
    print('在设备-厂家表的基础上，再根据网元细分，得到每个厂家对应的网元.........')
    #获取文件夹下的所有文件
    files_name = get_files(dir_path)
    for file_name in files_name:
        #1、按照网元分表
        file_path= os.path.join(dir_path, file_name)
        # 读取源文件遍历sheet（厂家）
        df = pd.read_excel(file_path, sheet_name=None)

        # 获得表头，为后边添加表头使用
        df_head = pd.read_excel(file_path, sheet_name=0)
        df_head_list = [df_head.columns.tolist()]
        # print('表头为\n', df_head_list)

        #得到所有的sheet
        sheets_name = list(df)
        # 遍历sheet，按照网元再分
        df1 = pd.DataFrame()
        # print('所有的sheet',sheets_name)#['华为','中兴','爱立信']

        for i in range (len(sheets_name)):
            # print(sheets_name[i])
            #读入文件
            df1 = pd.read_excel(file_path, sheet_name=sheets_name[i])
            netcell_values = df1['网元名称'].values.tolist()
            unique_netcell = list(set(netcell_values))  # 得出所有网元名（不重复）
            #创建新的excel文件，为后续写文件做铺垫

            # to_file_path = r'H:\pycharm_workspace\alarm_think\data\onenet_warning_data1\test2.xlsx'
            to_file_name = sheets_name[i]+file_name
            to_file_path = os.path.join(to_dir_path,to_file_name)
            pd.DataFrame().to_excel(to_file_path)#创建新excel文件

            for j in range (len(unique_netcell)):
                unique_netcell_values = []

                for k in range (len(netcell_values)):
                    if unique_netcell[j] == netcell_values[k]:
                        unique_netcell_values.append(df1.iloc[k])
                #添加头标题
                unique_netcell_values = df_head_list + unique_netcell_values
                df2 = pd.DataFrame(unique_netcell_values)
                #将某个设备-厂家-网元写入新的sheet

                excelWriter = pd.ExcelWriter(to_file_path,engine='openpyxl')
                to_sheet_name = file_name[2:-5]+sheets_name[i]+unique_netcell[j]
                hcu._excelAddSheet(df2, excelWriter, to_sheet_name)
                print('%s网元数据抽出，并写入文件成功!!' % to_sheet_name)

            print('\033[32;1m%s %s厂家的网元数据抽出完毕!!\033[0m'%(file_name[0:-5],sheets_name[i]))
        #带颜色输出控制台

    print('\033[33;1m' +'所有文件处理完毕，在设备-厂家表的基础上，再根据网元细分，得到每个厂家对应的网元！！'+ '\033[0m')


if __name__ == '__main__':
    dir_path=r'H:\Pycharm_workspace\alarm_think\data\onenet_warning_data1\get_equipment_factory'
    file_name='设备HSS FE.xlsx'
    to_dir_path=r'H:\pycharm_workspace\alarm_think\data\onenet_warning_data1\get_equipment_factory_netcell'
    file_path = os.path.join(dir_path, file_name)
    equip_factory_netcell_name='网元名称'
    get_equipment_factory_netcell(dir_path,to_dir_path)


    # get_files(dir_path)
    # test_str()
    # test_mkdir()
