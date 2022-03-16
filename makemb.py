import openpyxl

"""
makemb
  TODO 薬袋作成

  class MedBagInfo
    薬袋情報設定クラス
    ファイル名：薬袋_yyyymmdd.xlsx
      yyyymmdd = 処方日

"""
FILE_NAME = "薬袋‗"
FILE_EXT = ".xlsx"


class MedBagInfo:
    def __init__(self, path, df_list):
        self.dir_path = path
        self.rx_list = df_list
        self.file_name = FILE_NAME
        self.file_path = self.dir_path
        self._makefile()

    # TODO Excelファイル作成
    def _makefile(self):
        # 日付取得 -> ファイル名設定
        rx = self.rx_list[0]
        rx_date = rx.iat[1, 1]
        self.file_name += rx_date + FILE_EXT
        self.file_path += "\\" + self.file_name
