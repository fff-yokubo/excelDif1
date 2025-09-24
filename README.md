
# Excel x Git 差分管理

## 目的
Excelファイルの差分管理ができる仕組みを Git 上に構築すること。

このリポジトリは、Git に Excel ファイル（.xlsx）をコミットする際に自動で CSV に変換し、差分管理を容易にする仕組みを提供します。

主な構成:
- `pre-commit`: Git フック。ステージされた `.xlsx` ファイルを検出し、自動で CSV に変換してステージに追加します。
- `scripts/excel_to_csv.py`: Excel ファイルをシートごとに TSV (タブ区切りの CSV) へ変換する Python スクリプトです。

---

## 機能概要

### pre-commit フック
- ステージに追加された `.xlsx` ファイルを検出。
- Excel → CSV 変換を実行。
- 生成された CSV (`xlDif/<元ファイル名>/<シート名>.csv`) を Git に自動追加。
- 直前のコミットとの差分を確認し、変更がある場合のみログに通知。
- CSV が生成されなかった場合は空ディレクトリを削除。

### Python スクリプト (`excel_to_csv.py`)
- 引数で指定した Excel ファイルをシート単位で読み込み、CSV ファイルとして出力。
- 出力先は `xlDif/<Excelファイル名>/<シート名>.csv`。
- 区切り文字は **タブ (`\t`)**。
- 各シートのデータ行数をログに表示。
- ファイルやシートが読み込めない場合は警告・エラーを表示。

---

## セットアップ

1. `.git/hooks/pre-commit` に `pre-commit` をコピーし、実行権限を付与します。
   ```bash
   cp pre-commit .git/hooks/pre-commit
   chmod +x .git/hooks/pre-commit
````

2. Python 依存ライブラリをインストールします。

   ```bash
   pip install pandas openpyxl
   ```

---

## 使い方

### Git コミット時に自動変換

1. Excel ファイルを Git に追加:

   ```bash
   git add sample.xlsx
   ```

2. コミットすると、自動的に CSV が生成され、ステージに追加されます。

出力例:

```
[INFO] CSV files added for sample.xlsx
[INFO] Sheet1.csv has been exported (rows: 120)
[INFO] Changes detected in Sheet1.csv compared to previous commit: Update sales data (a1b2c3d)
```

### スクリプトを直接実行

```bash
$ python scripts/excel_to_csv.py sample.xlsx
```

出力例:

```
[INFO] Sheet1.csv has been exported (rows: 120)
[INFO] Sheet2.csv has been exported (rows: 45)
```

生成物:

```
xlDif/
  └─ sample/
       ├─ Sheet1.csv
       └─ Sheet2.csv
```

---


### フローの説明
1. **Excel ファイル**
   - 開発者が作業した Excel ファイル。
   - pre-commit フックで CSV に変換される。

2. **ローカル Git コミット**
   - Excel と生成された CSV が同じコミットに内包される。
   - CSV は Excel の内容を反映した自動生成物。
   - 差分管理は CSV を基準に行う。

3. **GitHub リポジトリ**
   - push によりリモートに送信。
   - 他ブランチとの並行作業も可能。

4. **Squash マージコミット**
   - 複数のコミットが 1 つにまとめられる。
   - CSV は最終状態だけがコミットに残る。
   - 以前の差分履歴は squash 後は保持されない。

---

## 注意点
- CSV は自動生成物であり、Squash Merge 後は最終状態のみ履歴に残ります。
- 複数ブランチで同じ CSV を変更すると merge 時にコンフリクトする可能性があります。
- 差分確認は生成された CSV ファイルを基準に行うことを推奨します。





---

## pre-commit フックのコード

```bash
#!/bin/bash

# pre-commit フック: ステージされた Excel ファイルを検出し、CSV に変換してステージに追加する
for file in $(git diff --cached --name-only | grep '\.xlsx$'); do
    outdir="xlDif/${file%.xlsx}"
    mkdir -p "$outdir"

    python scripts/excel_to_csv.py "$file"

    if compgen -G "$outdir/*.csv" > /dev/null; then
        git add "$outdir"/*.csv
        echo "[INFO] CSV files added for $file"

        last_commit_info=$(git log -1 --pretty=format:"%s (%h)" HEAD~1 2>/dev/null || echo "N/A")

        for csv in "$outdir"/*.csv; do
            if ! git diff --quiet HEAD~1 -- "$csv"; then
                echo "[INFO] Changes detected in $(basename "$csv") compared to previous commit: $last_commit_info"
            fi
        done
    else
        echo "[WARN] No CSV files generated for $file"
        echo "[INFO] Removing empty directory: $outdir"
        rm -rf "$outdir"
    fi
done
```

---

## Python スクリプト (`scripts/excel_to_csv.py`)

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
```

---

## 注意点

* 出力ファイルはタブ区切り (`.csv`) です。
* Excel 内の全シートが変換対象。
* 差分管理は生成された CSV を基準に行います。
* 大きな Excel ファイルでは変換に時間がかかる場合があります。

