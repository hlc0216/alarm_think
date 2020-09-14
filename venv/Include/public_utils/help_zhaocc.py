"""
    Author: Huang liuchao
    Contact: huanglc50@chinaunicom.cn
    Datetime: 2020/9/10 20:11
    Software: PyCharm
    File description：
        实现将数值转化为日期格式,并写入文件
"""
from dateutil.parser import parse
import pandas as pd
def num2time(file_path,sheet_name,col_name):
    df = pd.read_excel(file_path,sheet_name=sheet_name)
    num_list = []
    time_list = []
    num_list=df[col_name].tolist()#取出Envent Time那一列数据
    str_temp=''

    # 那一列数据中存在非整数（'Cold Restart'）,parse()方法只认识数字，不认识单词
    for num in num_list:
        if type(num) == int:
            time_list.append(parse(str(num)))
        else:
            time_list.append(num)

    print('Integer value is being converted to time format.....................')
    # 将转换成时间格式的list存入datafram，替换最初的Event Time那一列数据
    df[col_name] = time_list
    # 删除最后一列（pandas在写文件时，会默认在最后添加一列空值）
    df.drop(df.columns[len(df.columns) - 1], axis=1, inplace=True)
    # 写入原文件
    df.to_excel(file_path,index=False)
    print('数值转化为时间格式完成，写入文件成功！！')

if __name__ == '__main__':
    file_path=r'E:\Pycharm_workspace\alarm_think\data\onenet_warning_data1\8月21日爱立信HSS告警信息.xlsx'
    sheet_name=0
    col_name='EventTime'
    num2time(file_path,sheet_name,col_name)



