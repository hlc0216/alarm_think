import onenet_warning_utils as owu

if __name__ == '__main__':
    #HSS BE网元单一告警
    file_path = r'E:\Pycharm_workspace\alarm_think\data\testdata\HSS_BE网元单一告警  .xlsx'
    sheet_name = 0
    file_name = u'HSS BE网元单一告警'
    new_sheet_name = 'onenet_warning_collect'
    owu.onenet_warning(file_path, file_name,sheet_name, new_sheet_name)

