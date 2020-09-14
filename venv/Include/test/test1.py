import pandas as pd
import xlrd
import xlwt
import numpy as np
oss_path=r"E:\Pycharm_workspace\OSS2.0告警查询导出.xls"
oss_wangyuan='网元名称'
oss_gaojing='告警标题'

hss_path=r'E:\Pycharm_workspace\华为HSS U2000历史告警查询结果_20200821.xls'
hss_wangyuan='告警源'
hss_gaojing='名称'
def test_del_list():
    list1 = [[1,'tom'],[2,'jerry'],[3,'bob'],[4,'tony'],[6,'mary']]
    list2 = ['tom','bob','libai','harry','dufu','mary']
    list3=[]
    for i in range (len(list1)-1):
        # for j in range(len(list2)-1):
        if list1[i][1] not in list2:
            if list1[i] not in list3:
                list3.append(list1[i])
    print (list3)




def toSumList(file_path,wangyuan,gaojing):
    df=pd.read_excel(file_path)
    print('%s读入文件完毕,包含的列有',file_path)
    print(df.columns)
    wangyuan_list=[]
    gaojing_list=[]
    hebing_list=[]
    one_list=df[wangyuan].values.tolist()
    two_list=df[gaojing].values.tolist()
    sum_list=list(zip(t_list,d_list))
    print('sumlist合并完成')
    return sum_list
def shaijian(sum_list1,sum_list2,shaijian_list):
    shaijian_list=(set(sum_list1))


if __name__ == '__main__':

    # df1=pd.DataFrame(td_list)
    # df1.to_excel(r'E:\Pycharm_workspace\test1.xlsx',index=False)
    # print("写入文件完成")
    # test_del_list()
    shaijian()

