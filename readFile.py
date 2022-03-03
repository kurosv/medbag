from tkinter import filedialog
import pandas as pd


# ================================================
# ファイル読み込みクラス
#  Excel/CSVファイルの読み込み
#   -> データフレームとして格納
# ================================================
class ReadFile:
    # ================================================
    # コンストラクタ
    # ================================================
    def __init__(self, path):
        # 初期ディレクトリ指定
        self.dir_name = path
        self.filepath = None
        # 読み込みデータ保持(list)
        self.file_df_lst = []

    # ================================================
    # ファイルを指定：ファイルパス取得
    #  v0.2 複数選択可
    # ================================================
    def _get_filepath(self):
        # CSVファイル
        typ = [('薬剤集計ファイル(CSV)', '*.csv')]
        # self.filepath = filedialog.askopenfilenames(filetypes=typ, initialdir=self.dir_name)
        self.filepath = filedialog.askopenfilename(filetypes=typ, initialdir=self.dir_name)

    # ================================================
    # ファイル読み込み：Excel/CSV判定処理
    # ================================================
    def read_file(self):
        # 読み込みファイル指定
        self._get_filepath()
        # CSV読み込み処理
        # v0.2 複数ファイル対応
        for f in self.filepath:
            if f.endswith('.csv'):
                try:
                    self._read_csv(f)

                except FileNotFoundError as e:
                    # キャンセル/ファイル不明
                    print(e)
                    continue

                finally:
                    pass

        return self.file_df_lst

    # ================================================
    # csv読み込み
    # ================================================
    def _read_csv(self, filepath):
        # 読み込みファイル指定
        input_file = filepath
        # 列数を明示的に指定
        # syukei.csv : 54
        # col_names = [f'{i:02}' for i in range(54)]  # zerofill;2d
        col_names = [f'{i}' for i in range(54)]
        # 指定ファイルをdfとして読み込み
        # 薬剤集計は文字コード:ANSI
        self.file_df_lst.append(pd.read_csv(input_file, header=None, encoding='ANSI', names=col_names))

    # ================================================
    # df取得
    # ================================================
    def get_file_df(self):
        return self.file_df_lst


if __name__ == '__main__':
    rf = ReadFile('C:\\Users\\user\\Desktop\\ph_aggre')
    rf.read_file()
