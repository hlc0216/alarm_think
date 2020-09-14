import pandas as pd
class MakePandas():

    def append_excel(self, df, content_list):
        """
        excel文件中追加内容
        :return:
        content_list:待追加的内容列表
        """
        print("进入主任务")
        ds = pd.DataFrame(content_list)
        print(ds)
        df = df.append(ds, ignore_index=True)
        excel_name = r"/test2.xlsx"
        excel_path = r'E:\Pycharm_workspace' + excel_name
        df.to_excel(excel_path, index=False, header=False)

    def remove_row(self, df, row_list):
        """
        excel删除指定列
        :param df:
        :param row_list:
        :return:
        """
        df = df.drop(columns=row_list)
        return df

    def create_excel(self):
        """
        创建excel文件
        :return:
        """

        file_path = os.path.dirname(os.path.abspath(__file__)) + "/demo.xlsx"
        df = pd.DataFrame(columns=["title", "content"])
        df.to_excel(file_path, index=False)


if __name__ == '__main__':
    excel_name = r"\test2.xlsx"
    excel_path = r'E:\Pycharm_workspace' + excel_name

    m = MakePandas()
    df = pd.read_excel(excel_path, header=None)
    b = []
    for i in range(1, 10):
        a = []
        a.append(i)
        a.append(i * 2)
        b.append(a)
    print('a的内容')
    print(a)
    print('b的内容')
    print(b)
    df = m.append_excel(df, b)
    # print(df)
