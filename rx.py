"""
  処方情報処理系

  ProcRx
    処方情報の初期処理、分割処理

"""


class ProcRx:
    def __init__(self, df):
        self.df = None
        self.rx_list = []
        self._format_df(df)

    # dfの初期処理
    # 不要情報の削除
    # ヘッダは読み込み時点で削除（header=1 指定）
    def _format_df(self, df):
        d = df
        # 不要列削除
        #  5-12行のみ使用
        d = d.iloc[:, 4:12]

        # データの '' の除去
        d = d.replace(r'^\'*(.*)\'*', r'\1', regex=True)

        # # 列名追加
        # aggdf.columns = self.col_name
        # # 薬品名をIndex指定
        # aggdf = aggdf.set_index(self.index)

        # 処理反映
        self.df = d

    # 処方情報の分割取得
    def _get_rx(self):
        pass
