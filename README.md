## 概要

このプロジェクトは **Git で Excel ファイル（.xlsx）の変更履歴を管理しやすくする仕組み** です。
Excel はバイナリ形式のため Git で差分管理が困難ですが、コミット時に Excel をシートごとの **CSV (テキストファイル)** に自動変換し、CSV を差分管理対象にします。

これにより、Excel の更新内容を Git 上で「テキスト差分」として確認可能になります。

---

## 構成

```
.git/hooks/pre-commit     # Git フック（Excel を自動で CSV 化）
scripts/excel_to_csv.py   # Excel → CSV 変換用スクリプト
xlDif/                    # 変換後の CSV ファイル保存先
```

---

## 動作の流れ

1. **Excel を Git に追加 (`git add`)**
2. **コミット直前に pre-commit フックが発動**

   * 変更された `.xlsx` ファイルを検出
   * 出力先フォルダ `xlDif/<Excelファイル名>/` を作成
   * Python スクリプトを呼び出し、CSV に変換
   * 生成された CSV を Git に自動で追加
3. **コミット完了時に Excel と CSV 両方が Git に記録される**
4. **差分確認 (`git diff`) で CSV を比較できる**

---

## pre-commit フック

`.git/hooks/pre-commit` に配置するシェルスクリプトです。

```bash
#!/bin/bash

# pre-commit hook: Excelファイルが変更されていたらCSVに変換してステージに追加
for file in $(git diff --cached --name-only | grep '\.xlsx$'); do
    echo "Converting $file to CSV..."

    # 出力先のディレクトリを作成 (xlDif/<元ファイル名>/)
    outdir="xlDif/${file%.xlsx}"
    mkdir -p "$outdir"

    # Python スクリプトを実行して CSV に変換
    python scripts/excel_to_csv.py "$file"

    # 生成されたCSVを Git に追加
    git add "$outdir"/*.csv
done
```

### 処理詳細

* `git diff --cached --name-only` : コミット予定のファイル一覧を取得
* `grep '\.xlsx$'` : Excel ファイル（拡張子 .xlsx）のみ抽出
* `mkdir -p` : 出力先ディレクトリを作成（存在してもエラーにならない）
* Python スクリプトを呼び出し、対象ファイルを CSV 化
* `git add` : 生成された CSV をコミット対象に追加

---

## Python スクリプト

`scripts/excel_to_csv.py`

```python
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
        print(f"[WARN] 指定されたファイルが存在しません: {excel_path}")
        return  # 処理を中断して終了

    # Excelファイル名 (拡張子除去)
    base_name = os.path.splitext(os.path.basename(excel_path))[0]
    # 出力ディレクトリ (xlDif/<ファイル名>/)
    outdir = os.path.join("xlDif", base_name)
    os.makedirs(outdir, exist_ok=True)

    try:
        # Excelファイルを開き、シートの情報を保持するオブジェクトを作成
        xls = pd.ExcelFile(excel_path)
    except Exception as e:
        print(f"[ERROR] Excelファイルを開けませんでした: {excel_path}")
        print(f"詳細: {e}")
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

            print(f"[INFO] {sheet_name}.csv を出力しました (行数: {row_count})")
        except Exception as e:
            print(f"[WARN] シート {sheet_name} の処理でエラーが発生しました: {e}")

# ----------------------------------------------------------
# スクリプトを直接実行した場合の処理
# ----------------------------------------------------------
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python excel_to_csv.py <excel_file>")
        sys.exit(0)  # 引数なしでも異常終了しないように変更

    excel_file = sys.argv[1]
    excel_to_csv(excel_file)

```

### 処理詳細

1. **Excel のパスを受け取り** → `xlDif/<Excelファイル名>/` に出力フォルダを作成
2. **全シートを走査**
3. **シートごとに DataFrame として読み込み**
4. **シート名.csv** として保存

   * 区切り文字: タブ (`\t`)
   * 行番号は出力せず（`index=False`）
5. 処理結果を標準出力に表示

---

## セットアップ

### 1. 依存ライブラリのインストール

```bash
pip install pandas openpyxl
```

### 2. pre-commit フックの設置

```bash
cp pre-commit .git/hooks/pre-commit
chmod +x .git/hooks/pre-commit
```

---

## 利用方法

1. Excel ファイルを編集し、Git に追加

```bash
git add sample.xlsx
git commit -m "update excel"
```

2. コミット時に自動で CSV が生成される

```
xlDif/sample/Sheet1.csv
xlDif/sample/Sheet2.csv
```

3. 差分確認

```bash
git diff HEAD^ HEAD xlDif/sample/Sheet1.csv
```

→ Excel の変更内容をテキストとして確認可能

---

## メリット

* **Excel の差分を Git 上で確認可能**
* **複数人での共同開発・管理が容易**
* **バイナリ比較を避け、レビュー効率が向上**

