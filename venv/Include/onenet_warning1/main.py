"""
    Author: Huang liuchao
    Contact: huanglc50@chinaunicom.cn
    Datetime: 2020/9/15 9:27
    Software: PyCharm
    File description：
        主函数
"""
import Get_equipment as ge
import get_equipment_factory as gef
if __name__ == '__main__':

    #1、处理源数据得出各个设备表（sbc、scscf、hss fe等）

    ge_dir_path= r'E:\Pycharm_workspace\alarm_think\data\onenet_warning_data1'
    ge_file_name=r'OSS2.0告警查询导出20200821-1.xlsx'
    ge_file_path = os.path.join(dir_path,file_name)
    ge_equip_col_name='设备类型'
    ge.get_equipment(ge_file_path,0,ge_equip_col_name)


    #2、在设备表的基础上，将设备表根据厂家再细分为华为、中兴、爱立信.
    #   读取源文件，遍历sheet（设备），每个sheet按照厂家再分，处理完
    #   产生新的表，每个表有三个sheet（华为、中兴、爱立信）。
    gef_dir_path = r'E:\Pycharm_workspace\alarm_think\data\onenet_warning_data1'
    gef_file_name = r'OSS2.0告警查询导出20200821-1.xlsx'
    gef_file_path = os.path.join(dir_path, file_name)
    gef.get_equipment_factory(gef_dir_path, gef_file_name)
