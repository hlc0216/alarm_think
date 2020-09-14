import hlc_common_utils as hcu
import numpy as np
import pandas as pd
import xlrd
import os
import xlwt
import openpyxl
from openpyxl import load_workbook
import win32com.client as win32
def add_onenet_others_colum_values(source_file_path,to_file_path,sheet_name):
    print('开始抽取其他网元（除了scscf）对应的告警信息.......')
    #

def onenet_warning(file_path,file_name,sheet_name,new_sheet_name):
    print('%s开始处理数据.............'%file_name)
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    # df_value_li = df.values.tolist()
    #制作onenet_li和the_others_li
        #制作onenet_li
    onenet_li = []
    nrows=len(df)
    # print(df.iloc[[0],[1]].values[0][0])#获取第一行第二列单元格的值
    colum_index = df.columns.get_loc("告警标题")#获取列名所在列的索引
    for i in range(nrows):
        if pd.isnull(df.iloc[[i],[colum_index]].values[0][0]) :
            break
        else: #此处注意索引和真实行数存在两行的差距（即行数-索引=2）
            onenet_li.append([i,df.iloc[[i],[colum_index]].values[0][0]])
    # for i in range(10):#验证是否对照
    #     print('索引为%d，行数为%d,对应的告警为：%s'%(i,i+2,df.iloc[[i], [colum_index]].values[0][0]))

    print('单一网元告警list长度为：%d'%len(onenet_li))
    print('验证onenet_li二号元素:%s'%onenet_li[0][1])

        #制作the_other_li
    the_others_li = df['告警标题all'].values.tolist()
    print('除单一网元其余网元告警汇总list长度为：%d'%len(the_others_li))


    temp_li=[]#用于临时存储
    onenet_nrows = []# 将网元的单一警告对应的行抽出，存储到onenet_nrows[]
    onenet_warning_collect=[]
    #开始比对onenet_li和the_others_li，对onenet_li进行一个筛检，只保留其唯一的
    for i in range(len(onenet_li)):
        if onenet_li[i][1] not in the_others_li:
            if onenet_li[i] not in temp_li:#onenet_li[i][0]是行值，onenet_li[i][1]是告警值
                temp_li.append(onenet_li[i])#这个非必须，只用存下面的行即可
                onenet_nrows.append(onenet_li[i][0])#这个非必须，只用根据存的行抽取行数据放入onenet_warning_collect
                onenet_warning_collect.append(df.iloc[onenet_li[i][0]])#根据存取的行数，抽取行信息，存入onenet_warning_collect
    # onenet_warning_collect.append([df.iloc[onenet_nrows]])#这种也能抽出，但是是一个长度为1的一维数组，不方便pandas写excel

    #根据行数抽取源文件，存储到一个新的sheet中
    print(len(onenet_nrows))
    print(len(onenet_warning_collect))

    #开始写数据到新的sheet中
    df1=pd.DataFrame(onenet_warning_collect)#将二维list转换成dataframe
    to_file_path=file_path
    sheet_name=new_sheet_name
    excelWriter = pd.ExcelWriter(to_file_path,engine='openpyxl')
    hcu._excelAddSheet(df1, excelWriter, sheet_name)  # 为excel添加新的sheet，并将dataframe写入新的sheet（不覆盖其他sheet）
    print('%s网元处理数据完成，写入文件成功！'%file_name)