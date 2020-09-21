import pandas as pd
import xlrd
import os
import xlwt
import openpyxl
from openpyxl import load_workbook
# import xlsxwriter
import win32com.client as win32
def Match_column_data(file_path,colum_name,to_file_path):
    """
    有两个sheet中存在相同的列名，列数据是包含关系，根据sheet中的列数据抽取sheet2中
    包含sheet1列数据的行数据
    :param file_path: 文件路径（包含sheet1,sheet2）
    :param colum_name: 两个sheet都包含的列名
    :param to_file_path: 要写入的文件路径
    """
    print('开始')
    file_path = file_path
    df1 = pd.read_excel(file_path, sheet_name=0)
    list1 = df1[colum_name].values.tolist()
    df2 = pd.read_excel(file_path, sheet_name=1)
    list2 = df2.values.tolist()
    list3 = []

    for i in range(len(list1)):
        for j in range(len(list2)):
            if list1[i] == list2[j][1]:
                list3.append(list2[j])

    to_file_path = to_file_path
    df3 = pd.DataFrame(list3)
    excelWriter = pd.ExcelWriter(file_path, engine='openpyxl')
    to_sheet_name = 'result'
    hcu._excelAddSheet(df3, excelWriter, to_sheet_name)
    print('匹配列数据完成，并抽出数据写入新的sheet(result)中！')
def File_content_merge(file_path1,file_path2,to_file_path):
    """
    实现将两个（列名相同的）excel文件的内容合并到一个excel中
    :param file_path1: 文件1路径
    :param file_path2: 文件2路径
    :param to_file_path: 要写入的文件路径
    """
    print('实现将两个（列名相同的）excel文件的内容合并到一个excel中')
    file_path1 = file_path1
    file_path2 = file_path2
    df1 = pd.read_excel(file_path1, sheet_name=0)
    df1_head_list = [df1.columns.tolist()]
    df2 = pd.read_excel(file_path2, sheet_name=0)
    df2_head_list = [df2.columns.tolist()]
    df3 = pd.DataFrame()
    sumlist =  df1_head_list + df1.values.tolist()+df2_head_list + df2.values.tolist()
    to_file_path = to_file_path
    pd.DataFrame(sumlist).to_excel(to_file_path, header=None, index=None)
    print('两个excel文件文本合并完成！')
def append_excel_xlsx(file_path,sheet_name,values):
    """
    采用openpyxl实现在excel文件追加内容
    :param file_path: 文件路径
    :param sheet_name: sheet名（传入必须字符串）
    :param values:  要追加的数据（二维list）
    """
    wb = openpyxl.load_workbook(file_path)
    sheetnames = wb.sheetnames
    if type(sheet_name)==str:
        table = wb[sheet_name]
    else :
        table = wb[sheetnames[0]]
    table = wb.active
    print(table.title)  # 输出表名
    nrows = table.max_row  # 获得行数
    ncolumns = table.max_column  # 获得列数

    # 注意行业列下标是从1开始的
    for i in range(1, len(values) + 1):
        for j in range(1, len(values[i - 1]) + 1):
            table.cell(nrows + i, j).value = values[i - 1][j - 1]

    wb.save(file_path)
    print("追加写入excel数据成功！")

def write_excel_xlsx(file_path, sheet_name, value):
    """
    采用openpyxl将数据写入excel（会覆盖不是追加）
    :param file_path: 文件路径
    :param sheet_name:
    :param value:
    """
    index = len(value)
    workbook = openpyxl.Workbook()  # 新建工作簿（默认有一个sheet？）
    sheet = workbook.active  # 获得当前活跃的工作页，默认为第一个工作页
    sheet.title = sheet_name  # 给sheet页的title赋值
    for i in range(0, index):
        for j in range(0, len(value[i])):
            sheet.cell(row=i + 1, column=j + 1, value=str(value[i][j]))  # 行，列，值 这里是从1开始计数的
    workbook.save(file_path)  # 保存
    print("xlsx格式表格写入数据成功！")
def delete_files(root_file_path):
    for root ,dirs,files in os.walk(root_file_path):
        for name in files:
            if name.endswith('.xls'):
                os.remove(os.path.join(root,name))
                print('删除文件：'+os.path.join(root,name))

def add_sheet(file_path,sheet_name):
    """
    该函数是解决：为excel表格创建sheet
    :param file_path:
    :param sheet_name:
    """
    wb = wb = openpyxl.load_workbook(file_path)
    wb.create_sheet(title=sheet_name, index=None)
    wb.save(file_path)
    print('往 %s 文件添加 %s 成功！'%(file_path ,sheet_name))

def add_col_multi_values():
    """
    该函数是解决，同时向多个sheet写入数据
    """
    file_path = r'E:\Pycharm_workspace\华为HSS U2000历史告警查询结果_20200821.xls'
    to_file_path = r'E:\Pycharm_workspace\lemon_copy.xlsx'
    sheet_name = 'newSheet'
    df1 = pd.read_excel(file_path)
    df2 = pd.read_excel(file_path)

    writer = pd.ExcelWriter(to_file_path)
    df1.to_excel(writer,sheet_name='df_1',index=None)
    df2.to_excel(writer,sheet_name='df_2',index=None)
    writer.save()
    # writer.close()
    print('写入文件成功！')


def _excelAddSheet(dataframe,excelWriter,sheet_name):
    """
该函数主要是为excel添加新的sheet，并将dataframe写入新的sheet（不覆盖其他sheet）
这样就解决了直接使用to_excel覆盖原excel表中的sheet.

    :param dataframe:
    :param excelWriter:
    :param sheet_name:
    """
    book = load_workbook(excelWriter.path)
    excelWriter.book = book
    dataframe.to_excel(excel_writer=excelWriter,sheet_name=sheet_name,index=None,header=None)
    excelWriter.close()

def add_col_values1(file_path,to_file_path,sheet_name,title,value): #在excel最后一列添加列数据
    """
    向excel文件添加一列数据,并将添加数据后的表写入另一个excel文件新的sheet中
    :param file_path: 读入文件路径
    :param to_file_path: 要写入文件的路径
    :param sheet_name: 要新添加的sheet名字
    :param title: 增加列的表头名字
    :param value: 要增加列所对应的数据
    """

    #读入excel文件到dataframe
    df = pd.read_excel(file_path)
    #添加一列数据
    total_row = len(df)
    add_col_values=[]
    print('总共包含%d行', total_row)
    for i in range(total_row):  #制造数据
        add_col_values.append(value)
    df[title]=add_col_values

    #将dataframe写入文件,excel必需已经存在
    excelWriter = pd.ExcelWriter(to_file_path,engine='openpyxl')
    _excelAddSheet(df,excelWriter,sheet_name)#为excel添加新的sheet，并将dataframe写入新的sheet（不覆盖其他sheet）
    print('添加列数据完成，写入文件成功！')

def add_col_values(file_path,sheet_name,title,value): #在excel最后一列添加列数据
    """
    向excel文件添加一列数据,并将添加数据后的表写入原表新的sheet中（函数的重构）
    :param file_path: 要读入和写入的文件
    :param sheet_name: 要新添加的sheet名字
    :param title: 增加列的表头名字
    :param value: 要增加列所对应的数据

    """

    #读入excel文件到dataframe
    df = pd.read_excel(file_path)
    #添加一列数据
    total_row = len(df)
    add_col_values=[]
    print('总共包含%d行', total_row)
    for i in range(total_row):  #制造数据
        add_col_values.append(value)
    df[title]=add_col_values

    #将dataframe写入文件,excel必需已经存在
    excelWriter = pd.ExcelWriter(file_path,engine='openpyxl')
    _excelAddSheet(df,excelWriter,sheet_name)#为excel添加新的sheet，并将dataframe写入新的sheet（不覆盖其他sheet）
    print('添加列数据完成，写入文件成功！')

def extract_column_data(file_path,sheet_name,usecols,new_sheet_name):
    """
抽取表中指定列的数据，并存入新的sheet中
    :param file_path: 待抽取的文件
    :param sheet_name: 要存入的sheet_name
    :param usecols: 要抽取得列名
    """

    df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=usecols)
    df_head_li = [df.columns.tolist()]
    df_value_li = df_head_li + df.values.tolist() #添加列名

    df1 = pd.DataFrame(df_value_li)

    # 开始写入数据（不能直接使用toexcel,因为会覆盖以前的sheet）
    excelWriter = pd.ExcelWriter(file_path, engine='openpyxl')
    _excelAddSheet(df, excelWriter, new_sheet_name)  # 为excel添加新的sheet，并将dataframe写入新的sheet（不覆盖其他sheet）
    print('抽取列数据完成，写入文件成功！')

def xls2xlsx_transformat_file(xls_file_path):
    '''
    针对单个xls文件转化为xlsx
    :param xls_file_path:
    :return:转换成功后xlsx的文件路径
    '''
    file_name=xls_file_path
    if file_name.endswith('.xls'):
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb=excel.Workbooks.Open(file_name)
        wb.SaveAs(file_name + "x", FileFormat=51)  # FileFormat = 51 is for .xlsx extension
        wb.Close()  # FileFormat = 56 is for .xls extension
        excel.Application.Quit()
        print('文件格式转化成功！')
        xlsx_file_path = file_name + 'x'
    else:
        print('文件格式已经是xlsx，不需转化!')
        xlsx_file_path = file_name

    return xlsx_file_path


def xls2xlsx_transformat_dir(dir_path):
    """
    该函数主要解决：将文件夹下的所有xls文件转换为xlsx
    """
    ## 根目录
    # rootdir = u'E:\Pycharm_workspace\data1 - 副本'
    rootdir= dir_path
    # 三个参数：父目录；所有文件夹名（不含路径）；所有文件名
    for parent, dirnames, filenames in os.walk(rootdir):
        for fn in filenames:
            filedir = os.path.join(parent, fn)
            print(filedir)
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(filedir)
            # xlsx: FileFormat=51
            # xls:  FileFormat=56,
            # 后缀名的大小写不通配，需按实际修改：xls，或XLS
            wb.SaveAs(filedir.replace('xls', 'xlsx'), FileFormat=51)
            wb.Close()
            excel.Application.Quit()
    print('格式转换成功')


def num2time(file_path, sheet_name, col_name):
    '''
    将excel文件中某一列数值转换为日期格式，比如
    20200821010304    ------>    2020-08-21 01:03:04

    :param file_path:要处理的源文件路径
    :param sheet_name:源文件中sheet名字
    :param col_name:要转换的列名
    :return:null
    '''
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    num_list = []
    time_list = []
    num_list = df[col_name].tolist()  # 取出Envent Time那一列数据
    str_temp = ''

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
    df.to_excel(file_path, index=False)
    print('数值转化为时间格式完成，写入文件成功！！')


# if __name__ == '__main__':
#     path=xls2xlsx_transformat_file(r'E:\Pycharm_workspace\alarm_think\data\onenet_warning_data1\华为智能网告警8.21.xls')
#     print(path)