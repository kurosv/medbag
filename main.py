import readFile
# import insertSQL

import readFile
import setDir

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

    print(df)

    # 処方情報分割



if __name__ == '__main__':
    main()