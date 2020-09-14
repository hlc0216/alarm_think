import  pandas  as pd
df=pd.read_excel('E:\Pycharm_workspace\lemon.xlsx')
test_data=[]
for i in df.index.values:#获取行号的索引，并对其进行遍历：
    #根据i来获取每一行指定的数据 并利用to_dict转成字典
    # row_data=df.ix[i,['case_id','module','title','http_method','url','data','expected']].to_dict()
    row_data = df.loc[i, ['title', ]].to_dict()
    test_data.append(row_data)
print("最终获取到的数据是：{0}".format(test_data))
# print("输出值\n",df['data'].values)
testlist=df['data'].values
print(testlist)
# df['data'].to_excel('E:\Pycharm_workspace\lemontest.xlsx')
