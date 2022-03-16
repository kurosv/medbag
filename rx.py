"""
  処方情報処理系

  class ProcRx
    処方情報の初期処理、分割処理

"""


class ProcRx:
    def __init__(self, df):
        self.df = None
        self.rx_list = []
        # 初期処理開始
        self._format_df(df)
        self._get_rx()

    # dfの初期処理
    # 不要情報の削除
    # ヘッダは読み込み時点で削除（header=1 指定）
    def _format_df(self, df):
        d = df
        # 不要列削除
        #  5-12行のみ使用
        d = d.iloc[:, 4:12]

        # データの '' の除去
        # 分2 朝・夕食後 のみなぜか先頭’がない
        #  -> ?処理で対応
        d = d.replace(r'^\'?(.*)\'$', r'\1', regex=True)

        # 【コ】の行削除:comment
        comment_str = "【コ】"
        d = d[~d['6'].isin([comment_str])]

        # # 列名追加
        # aggdf.columns = self.col_name
        # # 薬品名をIndex指定
        # aggdf = aggdf.set_index(self.index)

        # 処理反映
        self.df = d
        self.df.to_csv('C:\\Users\\iruka\\Desktop\\medbag\\rx.csv',
                       header=False, index=False)

    # 処方情報の分割取得
    #   -> return 処方dfリスト
    def _get_rx(self):
        # 処方分割
        df_list = []

        for i in range(1, len(self.df)):
            d = self.df[self.df['4'] == i]
            if d.empty:
                break

            df_list.append(d)

        # list反映
        self.rx_list = df_list

    def get_rxlist(self):
        return self.rx_list
