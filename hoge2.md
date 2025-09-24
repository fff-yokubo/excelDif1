
# Excel → CSV 自動変換 (Git pre-commit フック付き)

## 目的
Excelファイルの差分管理ができる仕組みを Git 上に構築すること。

---

## Excel → CSV → Git → Squash Merge フロー

```

[Local]
Excelファイル更新
│
▼
Git add
│
▼
(Local) Git commit
   └─ CSV 生成物(当該Excelファイルに差分が発生していた場合のみ)
      (xlDif/\<Excelファイル名>/<シート名>.csv)
│
▼
git push
│
▼
[Remote]
merge --squash
│
▼
Squash マージコミット
(Excel + CSV の最終状態)

```

### フローの説明
1. **Excel ファイル**
   - 開発者が作業した Excel ファイル。
   - pre-commit フックで CSV に変換される。

2. **ローカル Git コミット**
   - Excel と生成された CSV が同じコミットに含まれる。
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
```

---

この図では：

* **CSV が Local Git Commit 内に内包**されていることが枝分かれ形式で明示
* 横向きフローでスクロールせず全体の流れを把握可能

必要であれば、次のステップとして **「Squash 前後の CSV 差分イメージ」付きフロー** も作って、さらに視覚的に分かりやすくできます。作りますか？
