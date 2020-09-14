import onenet_warning_utils as owu

if __name__ == '__main__':
    file_path=r'E:\Pycharm_workspace\alarm_think\data\\scscf网元单一告警.xlsx'
    file_name=u'scscf网元单一告警'
    sheet_name=0
    new_sheet_name='onenet_warning_collect'
    owu.onenet_warning(file_path,file_name,sheet_name,new_sheet_name)
