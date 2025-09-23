import pandas as pd
import sys
import os

# ----------------------------------------------------------
# ExcelファイルをシートごとにCSVファイルへ変換する関数
#   引数:
#       excel_path : 変換対象のExcelファイルパス
# ----------------------------------------------------------
def excel_to_csv(excel_path):
    # Excelファイル名 (拡張子除去)
    base_name = os.path.splitext(excel_path)[0]
    # 出力ディレクトリ (xlDif/<ファイルパス>)
    outdir = os.path.join("xlDif", base_name)
    os.makedirs(outdir, exist_ok=True)

    # Excelファイルを開き、シートの情報を保持するオブジェクトを作成
    xls = pd.ExcelFile(excel_path)

    # Excelファイル内の全てのシート名を順番に処理
    for sheet_name in xls.sheet_names:
        # 現在のシートをDataFrameとして読み込み
        df = pd.read_excel(excel_path, sheet_name=sheet_name)

        # CSVの出力ファイル名を作成
        # 例: "sample.xlsx" の "Sheet1" → "xlDif/sample/Sheet1.csv"
        csv_path = os.path.join(outdir, f"{sheet_name}.csv")

        # DataFrameをCSVとして保存
        # index=False により、行番号は出力せずデータのみ書き出す
        df.to_csv(csv_path, sep="\t", index=False)
        print(f"convert {sheet_name}.csv done")

# ----------------------------------------------------------
# スクリプトを直接実行した場合の処理
# ----------------------------------------------------------
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python excel_to_csv.py <excel_file>")
        sys.exit(1)

    excel_file = sys.argv[1]
    excel_to_csv(excel_file)
