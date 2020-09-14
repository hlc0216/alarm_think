import pandas as pd
import xlrd
import xlwt
import xlsxwriter
# workbook = xlsxwriter.Workbook(r'E:\Pycharm_workspace\lemontest.xlsx')
# df=pd.read_excel('E:\Pycharm_workspace\lemon.xlsx')
df=pd.read_excel('E:\Pycharm_workspace\lemon.xlsx')

# data= df.values
# dataframe = pd.DataFrame(data)
# print(dataframe)
# dataframe.to_excel('E:\Pycharm_workspace\lemontest.xlsx')//写入excel文件
# print('写入数据完成')
# print("获取所有的值:\n{0}".format(data))
# print("获取到所有的值L\n{0}".format(data))
# print("输出值\n",df['data'].values)

#读取excel中的数据
    #读取华为表，得到key和value
huawei_key_list=[]
huawei_value_list=[]
df=pd.read_excel('E:\Pycharm_workspace\华为HSS U2000历史告警查询结果_20200821.xls')
huawei_key_list=df['告警源'].values.tolist()  #获得key
# huawei_row_data=df.loc[1].values
# print(huawei_key_list)
result = open('E:\Pycharm_workspace\lemontest.xlsx', 'w', encoding='utf-8')

for i in range(len(huawei_key_list)):
    result.write(str(huawei_key_list[i]))
    result.write('\t')
result.write('\n')
result.close()
# dataframe=pd.DataFrame(df.loc[1].values)
# print(dataframe)
# dataframe.to_excel('E:\Pycharm_workspace\lemontest.xlsx')
# value_list =
# df['告警源'].to_excel('E:\Pycharm_workspace\lemontest.xlsx')

#制作两个字典oss和huawei
#对两个字典求取差值
#将差值输出到excel
