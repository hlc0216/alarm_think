华为漏网元查找方案

    一、需求分析：主要解决专业网管（文件1）漏报到oss2.0系统的网元（文件2）
    文件1：华为HSS U2000历史告警查询结果_20200821.xls
    文件2：OSS2.0告警查询导出20200821-1.xls

    二、方案设计：
    1.选取的两列：    oss--->网元名称       huawei_hss--->告警源
    2.对比选取的两列，查找出漏报的。

    三、具体实现：
        使用python pandas分别从文件中读取两列数据，存入list
        huawei_hss_list=[[]] #二维list，value1=行号 value2 = 列数据
        oss_list[]  #一维list value=列数据
        for i in range(len(huawei_hss_list))
            huawei_hss_list[i][1] is not oss_list
                #记录行数
                #根据行数写到该excel的新sheet
        代码：zte_missing_netcell.py

    四、验证
        使用countif()函数验证 （验证通过，与代码运行结果相同）

    五、注意问题
        1 出来的结果网元有重复的，要不要进行去重


