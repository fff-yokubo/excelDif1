了解です。先ほどの **完全版 Python スクリプト** の概要と使い方をまとめた **Markdown ドキュメント**を作成しました。

---

# 架空店舗売上データ生成スクリプト — ドキュメント

## 概要

このスクリプトは、**架空の店舗売上データ**を含む複雑で巨大な Excel ファイルを自動生成します。
データはすべてランダムに生成され、テスト、パフォーマンス検証、レポートテンプレートの確認などに利用できます。

出力される Excel ファイルには以下のシートが含まれます：

* **Products**
  商品マスタ（200 SKU）。カテゴリ、原価、希望小売価格（MSRP）を含む。
* **Stores**
  店舗マスタ（50 店舗）。地域、開店日、マネージャー名を含む。
* **Employees**
  従業員データ（500 名）。所属店舗、役職、採用日などを含む。
* **Sales**
  売上トランザクション（デフォルト 20,000 行）。

  * 各取引には日付、店舗、商品、数量、単価、割引が含まれる
  * **NetSale / TotalCost / Profit は Excel 数式で埋め込み**（Excel 上で自動計算される）
* **MonthlySummary**
  年月ごとの NetSale / TotalCost / Profit 集計（値として書き出し）
* **PivotLike**
  カテゴリ × 年月 の売上を **SUMIFS 関数**で集計する擬似ピボット表
* **MonthlyChart**
  月別売上の推移をプロットしたグラフを貼り付け
* **README**
  生成日時・行数・範囲などのメタ情報

---

## 必要な環境

Python ライブラリ：

```bash
pip install pandas numpy openpyxl matplotlib
```

---

## 使い方

### 基本実行

```bash
python generate_complex_excel.py
```

デフォルトで `complex_huge_sample.xlsx` が生成されます（Sales シートは 20,000 行）。

### オプション

* `--rows` : Sales の行数を指定（例：50,000 行）
* `--out` : 出力ファイル名を指定
* `--seed` : 乱数シードを指定（同じシードなら再現可能）

例：

```bash
# 50,000 行の Sales を含むファイルを生成
python generate_complex_excel.py --rows 50000 --out store_sales_50k.xlsx --seed 42
```

---

## 出力ファイルの特徴

* **Sales シート**

  * 列 L: `NetSale = Quantity * UnitPrice * (1 - Discount)`
  * 列 M: `TotalCost = VLOOKUP(ProductID, Products!A:D, 4, FALSE) * Quantity`
  * 列 N: `Profit = NetSale - TotalCost`
  * → Excel を開くと自動計算される

* **MonthlySummary**

  * 年月単位で売上を集計（NetSale / TotalCost / Profit）

* **PivotLike**

  * 各カテゴリごとに年月別売上を **SUMIFS** で集計
  * 例:
    `=SUMIFS(Sales!L:L, Sales!H:H,"Beverages", Sales!B:B,">=2023-01-01", Sales!B:B,"<=2023-01-31")`

* **MonthlyChart**

  * `MonthlySummary` を基に描画したグラフを画像として貼付

---

## 応用

* `--rows` を大きくすればパフォーマンステスト用の巨大ファイルを生成可能
* カラムやデータスキーマを拡張して、顧客属性・決済方法・返品情報などを追加可能
* 本物のピボットテーブルを作成したい場合は、VBA や Excel COM を組み合わせる必要あり

---

👉 このスクリプトは **テスト用データ生成**に最適です。本番利用する場合は実データと混同しないよう注意してください。

---

この説明文を README.md にしてプロジェクトに同梱すると便利です。

---

ご希望なら、この **Markdown ドキュメントを README.md として出力する Python コード** も追加で書きますか？
