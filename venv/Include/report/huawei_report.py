import hlc_common_utils as hcu
if __name__ == '__main__':
    root_path= u'E:\Pycharm_workspace\data1' #将此目录下的xls文件修改为xlsx或者删除指定文件
    file_path = r'E:\Pycharm_workspace\data\华为HSS U2000历史告警查询结果_20200821.xlsx'
    to_file_path = r'E:\Pycharm_workspace\data\lemon_copy.xlsx'

    # add_sheet(r'E:\Pycharm_workspace\lemon_copy.xlsx','hlc_sheet')
    # xls2xlsx_transformat()

    #为hss_miss添加厂家,将数据写入另一个sheet
    hss_miss_file_path=r'E:\Pycharm_workspace\alarm_think\data\report_data\hss_miss.xlsx'
    hcu.add_col_values(hss_miss_file_path, 'add_col_values', '厂家', '华为')
    #抽取想要的列数据,并将数据写入另一个sheet
    sheet_name='add_col_values'
    new_sheet_name='extract_column_data'
    use_cols=['厂家', '名称', '告警源', '发生时间(NT)']
    hcu.extract_column_data(hss_miss_file_path,sheet_name,use_cols,new_sheet_name)
