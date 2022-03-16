"""
makemb
  TODO 薬袋作成
  ・患者ごとにシート分け
  ・シートの縦方向に順次作成
  ・薬袋1枚につき
      - 24行（予定）
      - 3列（幅：2,10,2）
  ・余白：全て0

  class MedBagInfo
    薬袋情報設定クラス
    ファイル名：薬袋_yyyymmdd.xlsx
      yyyymmdd = 処方日

"""
import openpyxl

# 薬袋ファイル設定
FILE_NAME = "薬袋‗"
FILE_EXT = ".xlsx"
# PAPER_SIZE = "A5"
# MARGIN = 0

# 見出し種別
INDEX_ORAL = "[内服薬]"
INDEX_ORAL_ODO = "[内服薬(一包化)]"
INDEX_EXT = "[外用薬]"
INDEX_INJ = "[注射薬]"
INDEX_ASN = "[屯服薬]"

# 薬袋種別
BAG_ORAL = "内　服　薬"
BAG_EXT = "外　用　薬"
BAG_INJ = "注　射　薬"
BAG_ASN = "屯　服　薬"

# 薬袋色設定
COLOR_ORAL = 0x0070C0  # 内服
COLOR_EXT = 0xFF0000  # 外用
COLOR_ASN = 0x00B050  # 頓服
COLOR_ASA = 0xFF0000  # 朝食後
COLOR_HIRU = 0x00B050  # 昼食後
COLOR_YU = 0x0070C0  # 夕食後
COLOR_WAKE = 0xFC6E04  # 起床時
COLOR_SLP = 0x000000  # 寝る前

# 用法
ODO = "１回　１包"

# 薬局情報/スタンプ欄
PH_NAME = "イルカ薬局　錦町店"
PH_ADDRESS = "小樽市錦町５番１０号"
PH_TEL = "TEL　0134-21-5775"
PH_FAX = "FAX　0134-21-5776"
STAMP_INDEX = "薬剤師"

PT_TITLE = "　様"

# 行設定：基本0開始
MAX_ROW = 24  # 1薬袋の行数
ROW_DATE = 0  # 日付行
ROW_INDEX = 1  # 内服・外用・頓服表示
ROW_PT = 3  # 患者名
ROW_DOSAGE = 5  # 用法用量開始位置（⁺2行）
ROW_MED = 7  # 薬品名開始位置（⁺9行）
ROW_PH = 20  # 薬局名挿入位置
ROW_STAMP = 20  # 印鑑枠INDEX挿入位置
# 行高
HTIGHT_INDEX = 2  # 内服・外用・頓服表示
HEIGHT_PT = 1.7  # 患者名、用法メイン
HEIGHT_DOSAGE = 1  # 用法用量
HEIGHT_DATE = 0.75  # 日付、薬局名
HEIGHT_NORMAL = 0.6  # normal

# 列設定：基本0開始
MAX_COL = 3  # 1薬袋の列数
COL_MAIN = 1  # メイン表記位置
COL_STAMP = 2  # 印鑑枠の列
# 列幅
WIDTH_MAIN = 10
WIDTH_ETC = 2


class MedBagInfo:
    def __init__(self, path, df_list):
        self.dir_path = path
        self.rx_list = df_list
        self.file_name = FILE_NAME
        self.file_path = self.dir_path
        self.rx_date = ""
        self._makefile()

    # 処方内容取得
    def _get_rx_info(self):

        for rx in self.rx_list:
            hp = rx.iat[1, 4]  # 病院名
            pt = rx.iat[0, 4]  # 患者名
            # TODO 処方情報のまとめこみ処理
            for row in rx:
                pass

    # TODO Excelファイル作成
    #   同名ファイルは上書き
    def _makefile(self):
        # 日付取得 -> ファイル名設定
        rx = self.rx_list[0]
        self.rx_date = rx.iat[1, 1]
        self.file_name += self.rx_date + FILE_EXT
        self.file_path += "\\" + self.file_name

    # TODO シート設定
    #   -> 患者名シートの作成、印刷設定
    #   印刷設定は行高設定が未着手
    def _sheet_setting(self, wb):
        # TODO シート作成：患者名

        # シート内設定
        for ws in wb.worksheets:
            # 列幅設定
            ws.column_dimensions['A'].width = WIDTH_ETC
            ws.column_dimensions['B'].width = WIDTH_MAIN
            ws.column_dimensions['C'].width = WIDTH_ETC
            # TODO 行高設定
            #   ループ回して設定？

            # 印刷設定
            self._print_setting(ws)

        wb.save(self.file_path)

    # 印刷設定：シートごと
    def _print_setting(self, ws):
        wps = ws.page_setup
        # 印刷の向きを設定：縦
        wps.orientation = ws.ORIENTATION_LANDSCAPE
        # 横を１ページ
        wps.fitToWidth = 1
        # 縦を自動
        wps.fitToHeight = 0
        # fitTo属性を有効にする
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        # 用紙サイズを設定
        wps.paperSize = ws.PAPERSIZE_A5

        # 余白
        ws.page_margins.left = 0
        ws.page_margins.right = 0
        ws.page_margins.top = 0
        ws.page_margins.bottom = 0
        ws.page_margins.header = 0
        ws.page_margins.footer = 0
