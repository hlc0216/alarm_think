网元单一告警查找方案

    一、需求分析：
        主要统计oss2.0系统中的各个网元存在的唯一告警

    二、方案设计：
    1.选取的两列：    oss--->网元名称       huawei_hss--->告警源
    2.对比选取的两列，查找出漏报的。

    三、具体实现：
        1.新增一列数据（告警标题all）除了待处理网元之外的网元告警汇总
        2.使用python pandas分别从文件中读取两列数据，存入list
            onenet_li=[[]] #二维list，value1=行号 value2 = 列数据
            the_others_li[]  #一维list value=列数据
            onenet_warning_collect[[]] #二维list 方便后边写入excel
        3.开始比对onenet_li和the_others_li，对onenet_li进行一个筛检，只保留其唯一的
            for i in range(len(onenet_li)):
                if onenet_li[i][1] not in the_others_li:
                    if onenet_li[i] not in temp_li:#onenet_li[i][0]是行值，onenet_li[i][1]是告警值
                        temp_li.append(onenet_li[i])#这个非必须，只用存下面的行即可
                        onenet_nrows.append(onenet_li[i][0])#这个非必须，只用根据存的行抽取行数据放入onenet_warning_collect
                        onenet_warning_collect.append(df.iloc[onenet_li[i][0]])
        4.根据行数抽取源文件，存储到一个新的sheet中
            excelWriter = pd.ExcelWriter(to_file_path,engine='openpyxl')
            rh._excelAddSheet(df1, excelWriter, sheet_name)
        文件列表：
            Gomc网元单一告警  .xlsx
            HSS_BE网元单一告警  .xlsx
            HSS_FE网元单一告警  .xlsx
            MGW网元单一告警 .xlsx
            MMTEL_AS网元单一告警 .xlsx
            MSCSERVER网元单一告警.xlsx
            OSS2.0告警查询导出20200821-1.xlsx
            sbc网元单一告警.xlsx
            scscf网元单一告警.xlsx
        代码列表：
            Gomc_onenet_warning.py
            HSS_BE_onenet_warning.py
            HSS_FE_onenet_warning.py
            MGW_onenet_warning.py
            MMTEL_AS_onenet_warning.py
            MSCSERVER_onenet_warning.py
            onenet_warning_utils.py
            readme.txt
            sbc_onenet_warning.py
            scscf_onenet_warning.py
        5.每个代码对应处理一个网元对应的excel
    四、验证
        使用countif()函数验证 （验证通过，与代码运行结果相同）

    五、注意问题
        1 出来的结果网元有重复的，要不要进行去重


