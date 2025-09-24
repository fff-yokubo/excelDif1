import pandas as pd
import sys
import os

# ----------------------------------------------------------
# ExcelファイルをシートごとにCSVファイルへ変換する関数
#   引数:
#       excel_path : 変換対象のExcelファイルパス
# ----------------------------------------------------------
def excel_to_csv(excel_path):
    # ファイル存在チェック
    if not os.path.exists(excel_path):
        print(f"[WARN] File does not exist: {excel_path}")
        return

    # Excelファイル名 (拡張子除去)
    base_name = os.path.splitext(os.path.basename(excel_path))[0]
    # 出力ディレクトリ (xlDif/<ファイル名>/)
    outdir = os.path.join("xlDif", base_name)
    os.makedirs(outdir, exist_ok=True)

    try:
        # Excelファイルを開き、シートの情報を保持するオブジェクトを作成
        xls = pd.ExcelFile(excel_path)
    except Exception as e:
        print(f"[ERROR] Could not open Excel file: {excel_path}")
        print(f"Details: {e}")
        return

    # Excelファイル内の全てのシート名を順番に処理
    for sheet_name in xls.sheet_names:
        try:
            # 現在のシートをDataFrameとして読み込み
            df = pd.read_excel(excel_path, sheet_name=sheet_name)

            # CSVの出力ファイル名を作成
            csv_path = os.path.join(outdir, f"{sheet_name}.csv")

            # DataFrameをCSVとして保存
            df.to_csv(csv_path, sep="\t", index=False)

            # 出力した行数（ヘッダを含めないデータ行数）
            row_count = len(df)

            print(f"[INFO] {sheet_name}.csv has been exported (rows: {row_count})")
        except Exception as e:
            print(f"[WARN] Error occurred while processing sheet {sheet_name}: {e}")

# ----------------------------------------------------------
# スクリプトを直接実行した場合の処理
# ----------------------------------------------------------
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python excel_to_csv.py <excel_file>")
        sys.exit(0)

    excel_file = sys.argv[1]
    excel_to_csv(excel_file)
