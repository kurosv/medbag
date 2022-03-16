import makemb
import readFile
import setDir
import rx

# ===============================================
# 定数定義
# ===============================================
DIR_NAME = "medbag"  # ディレクトリ名


# ================================================
# main
# ================================================
def main():
    # ディレクトリ確認
    # なければ作成
    sd = setDir.SetDir(DIR_NAME)
    path = sd.path

    # ファイル読み込み
    rf = readFile.ReadFile(path)
    df = rf.read_file()
    if df is None:
        exit()

    # 処方情報処理
    r = rx.ProcRx(df)
    rx_list = r.get_rxlist()

    # 薬袋ファイル作成
    mb = makemb.MedBagInfo(path, rx_list)



if __name__ == '__main__':
    main()
