"""
    Author: Huang liuchao
    Contact: huanglc50@chinaunicom.cn
    Datetime: 2020/9/15 9:36
    Software: PyCharm
    File description：
        
"""
import os
import Get_equipment as ge
if __name__ == '__main__':
    dir_path= r'E:\Pycharm_workspace\alarm_think\data\onenet_warning_data1'
    file_name=r'OSS2.0告警查询导出20200821-1.xlsx'
    file_path = os.path.join(dir_path,file_name)
    equip_col_name='设备类型'
    ge.get_equipment(file_path,0,equip_col_name)