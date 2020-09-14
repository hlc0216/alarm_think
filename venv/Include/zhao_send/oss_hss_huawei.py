import xlwt, xlrd
from xlutils.copy import copy


def oneWarning_xls(xls_oss_path, new_xls):
    # 读取 xls_oss_path excel
    data_oss_xls = xlrd.open_workbook(xls_oss_path)  # 打开对应地址下的excel文件
    sheet_oss_name = data_oss_xls.sheets()[0]  # 进入第一张表
    print(sheet_oss_name)
    count_nrows_oss = sheet_oss_name.nrows  # 获取总行数
    count_nocls_oss = sheet_oss_name.ncols  # 获得总列数
    line_value_oss = sheet_oss_name.row_values(0)  # 取第一行的值作为字典的key

    # k =1 #漏网元存储在新建excel中的行下标（漏网元的个数
    wb_all = xlrd.open_workbook(new_xls)  # 打开待粘贴的表
    sheets = wb_all.sheet_names()  # 获取工作簿中的所有表格
    sheet2 = wb_all.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个sheet
    new_wb = copy(wb_all)  # 将xlrd对象拷贝转化为xlwt对象，要用到xlutils模块
    new_sheet = new_wb.get_sheet(0)  # 获取转化后工作簿中的第一个sheet

    # 创建一个dict={key=行数：value=[设备厂家，告警标题]}
    dct = dict()  # 创建一个空字典
    for i in range(1, count_nrows_oss):
        dct[i] = [sheet_oss_name.cell(i, 6).value.strip(), sheet_oss_name.cell(i, 9).value.strip()]
    # print(dct)
    # 向新建表中插入属性标题
    for m, content in enumerate(line_value_oss):
        new_sheet.write(0, m, "".join(content))

    k = sheet2.nrows  # 获取表中已经有的数据行数
    # num_miss = 0，写入数据前加标题
    for i in range(1, count_nrows_oss - 1):
        count = 0  # 计算第i行与后面行数中设备类型不同，报警类型不同
        for j in range(i + 1, count_nrows_oss):
            if dct[i][0] == dct[j][0]:  # 首先排除设备类型相同的行数
                continue
            elif dct[i][0] != dct[j][0]:  # 设备类型不同
                if dct[i][1] == dct[j][1]:  # 报警相同时候，停止循环
                    break

                else:  # 排除前面情况下，统计设备不同，报警不同的个数
                    count = count + 1
        if count >= 1:
            row = sheet_oss_name.row_values(i)
            for j, content in enumerate(row):
                new_sheet.write(k, j, content)
            k = k + 1
    new_wb.save(new_xls)
    print(k)

oss_path = r'F:\Warning0828\data\OSS2.0告警查询导出20200821-1.xls'
new_xls = r'F:\Warning0828\03\data\单一网元报警\one_waring.xls'
oneWarning_xls(oss_path, new_xls)









