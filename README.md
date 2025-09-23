

# Excel Diff with Git

Excelファイル（`.xlsx`）をGitで効率的に管理するための仕組みです。
通常、Excelはバイナリ形式のため差分が確認できませんが、本プロジェクトでは **コミット時に自動でCSVへ変換** し、
Gitのテキスト差分として変更点を可視化できるようにしています。

---

## 機能概要

- Gitの **pre-commitフック** により、Excelファイルが変更された場合に自動的にCSVへ変換
- 変換処理は `scripts/excel_to_csv.py` によって実行
- 生成されたCSVファイルはExcelと同時にGitへコミットされる
- CSV形式なので `git diff` でデータの変更点を直接確認可能

---

## ディレクトリ構成

```

.
├── .git/hooks/pre-commit        # コミット前に実行されるフックスクリプト
├── scripts/
│   └── excel\_to\_csv.py          # ExcelをシートごとにCSVへ変換するスクリプト
└── README.md                    # 本ドキュメント

````

---

## pre-commit フックスクリプト

`.git/hooks/pre-commit` に以下を設定してください。
Excelファイルがステージングされていた場合、自動でCSVへ変換されます。

```bash
#!/bin/bash

# Excelファイルが変更されていたらCSVに変換
for file in $(git diff --cached --name-only | grep '.xlsx$'); do
    echo "Converting $file to CSV..."
    python scripts/excel_to_csv.py "$file"
    git add "${file%.xlsx}"*.csv
done
````

### 実行権限の付与

```bash
chmod +x .git/hooks/pre-commit
```

---

## Excel → CSV 変換スクリプト

`scripts/excel_to_csv.py` では、ExcelファイルをシートごとにCSVファイルへ変換します。

```python
import pandas as pd
import sys
import os

def excel_to_csv(excel_path):
    xls = pd.ExcelFile(excel_path)
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        csv_path = f"{os.path.splitext(excel_path)[0]}_{sheet_name}.csv"
        df.to_csv(csv_path, index=False)

if __name__ == "__main__":
    excel_file = sys.argv[1]
    excel_to_csv(excel_file)
```

### 処理の流れ

1. Excelファイルを読み込み、含まれるすべてのシートを取得
2. 各シートをDataFrameに変換（pandasを利用）
3. シートごとに `元ファイル名_シート名.csv` の形式で保存

   * 例: `report.xlsx` の `Sheet1` → `report_Sheet1.csv`
4. pre-commitフック経由で呼び出され、CSVがコミット対象に追加される

---

## 使い方

1. Excelファイルをリポジトリに追加または変更する

   ```bash
   git add sample.xlsx
   ```

2. コミットを実行すると、フックが動作してCSVファイルが生成される

   ```bash
   git commit -m "Update Excel data"
   ```

3. Gitには以下が登録される

   * 元のExcelファイル
   * シートごとに出力されたCSVファイル

---

## メリット

* **Excelのまま残す** → ユーザーは通常通りExcelで編集可能
* **CSVで差分確認** → Git上ではテキスト形式で変更点を追跡可能
* **レビューが容易** → Pull Requestなどで変更内容を確認しやすい

---

## 必要な環境

* Python 3.x
* ライブラリ: pandas, openpyxl

インストール例:

```bash
pip install pandas openpyxl
```
