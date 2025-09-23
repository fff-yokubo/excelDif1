import pandas as pd
import sys
import os

# ----------------------------------------------------------
# ExcelファイルをシートごとにCSVファイルへ変換する関数
#   引数:
#       excel_path : 変換対象のExcelファイルパス
# ----------------------------------------------------------
def excel_to_csv(excel_path):
    # Excelファイルを開き、シートの情報を保持するオブジェクトを作成
    xls = pd.ExcelFile(excel_path)

    # Excelファイル内の全てのシート名を順番に処理
    for sheet_name in xls.sheet_names:
        # 現在のシートをDataFrameとして読み込み
        df = pd.read_excel(excel_path, sheet_name=sheet_name)

        # CSVの出力ファイル名を作成
        # 例: "sample.xlsx" の "Sheet1" → "sample_Sheet1.csv"
        # os.path.splitext(excel_path)[0] で拡張子を除いたファイル名を取得
        csv_path = f"xlDif/{sheet_name}.csv"

        # DataFrameをCSVとして保存
        # index=False により、行番号は出力せずデータのみ書き出す
        df.to_csv(csv_path, sep = "\t", index=False)
        print('convert %s.csv done'%sheet_name)

# ----------------------------------------------------------
# スクリプトを直接実行した場合の処理
# ----------------------------------------------------------
if __name__ == "__main__":
    # コマンドライン引数からExcelファイルのパスを取得
    # 実行例: python script.py sample.xlsx
    excel_file = sys.argv[1]

    # 指定されたExcelファイルをCSVへ変換
    excel_to_csv(excel_file)

