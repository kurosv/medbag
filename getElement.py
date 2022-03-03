import re


class GetElement:
    def __init__(self, df_lst):
        self.index = "ph_name"
        self.col_name = [self.index, ""]
        # 引数：データフレーム(list)
        self.df_lst = df_lst
        # 集計月取得
        self.get_aggdate()

    # ===============================================
    # 集計月取得
    #  集計月：df(1, 5) -> "-1"指定！
    # ===============================================
    def get_aggdate(self):
        # yyyy年 m月 d日 -> yyyy/mm に変換
        d = self.df_lst.iat[0, 4]
        d = re.findall(r'[0-9]+', d)
        # yyyy + mm : m 0埋め右寄せ指定
        agg_date = d[0] + "/" + f'{d[1]:0>2}'
        self.col_name[1] = agg_date

    # ===============================================
    # データフレーム列抽出
    #  医薬品名：9, 数量：10,
    #    -> "-1"で指定！
    # ===============================================
    def get_aggdf(self):
        # 必要要素抽出
        aggdf = self.df_lst.iloc[:, [8, 9]]
        # 1-3行目はヘッダ行なので削除
        aggdf = aggdf.drop(range(3))
        # 列名追加
        aggdf.columns = self.col_name
        # 薬品名 '' の除去
        aggdf = aggdf.replace(r'^\'(.*)\'$', r'\1', regex=True)
        # 薬品名をIndex指定
        aggdf = aggdf.set_index(self.index)

        return aggdf

# TODO 内服・外用など区分ごとにdf分ける？(保留:要検討)
