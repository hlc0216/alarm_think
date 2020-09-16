import hlc_common_utils as hcu
def test_append_excel():
    print('追加文本到excel!')

def test_color():
    file_name='hello.xlsx'
    to_sheet_name = 'to_sheet_name'
    sheet_name = 'sheet_name'
    print('\033[30;47;1m白底黑字\033[0m')
    print('\033[30;41;1m红底黑字\033[0m')
    print('\033[33;1m黄色\033[0m')
    print('\033[32;1m绿色\033[0m')
    print('%s网元数据抽出，并写入文件成功!!' % to_sheet_name)


    print('\033[31;1m')

    print('-' * 10)
    print('我是第六代火影')
    print('-' * 10)
    print('\033[0m')

    print('\033[32;1m%s %s厂家的网元数据抽出完毕!!\033[0m' % (file_name[0:-5], sheet_name))

    # 带颜色输出控制台（绿色）