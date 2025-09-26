import openpyxl
import argparse

def colnum_string(n):
    """列番号をA1形式の列名に変換する (例: 1 → A, 27 → AA)"""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result

def load_excel_as_dict(path):
    """
    Excelを辞書形式に変換する
    {
        "Sheet1": {(row, col): value, ...},
        ...
    }
    """
    print(f"[INFO] Excelファイルを読み込み中: {path}")
    wb = openpyxl.load_workbook(path, data_only=True)
    sheets_dict = {}

    for sheet in wb.sheetnames:
        print(f"[INFO] シート読込中: {sheet}")
        ws = wb[sheet]
        values = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:  # 値があるセルだけを保持
                    values[(cell.row, cell.column)] = str(cell.value)
        sheets_dict[sheet] = values
    return sheets_dict

def compare_excels(file1, file2, output_md="diff_report.md"):
    """
    2つのExcelファイルを比較して差分をMarkdown形式で出力する
    """
    print(f"[INFO] ファイル比較開始: {file1} vs {file2}")

    # Excelファイルを辞書形式に変換
    data1 = load_excel_as_dict(file1)
    data2 = load_excel_as_dict(file2)

    report_lines = []
    report_lines.append(f"# Excel差分レポート\n")
    report_lines.append(f"- 比較元: `{file1}`\n- 比較先: `{file2}`\n")

    # シート名の集合を取得
    sheets1 = set(data1.keys())
    sheets2 = set(data2.keys())

    # 差分分類
    added_sheets = sheets2 - sheets1
    removed_sheets = sheets1 - sheets2
    common_sheets = sheets1 & sheets2

    # 追加されたシートを出力
    if added_sheets:
        report_lines.append("## 追加されたシート\n")
        for s in added_sheets:
            report_lines.append(f"- {s}")
            print(f"[INFO] 追加されたシート検出: {s}")
        report_lines.append("")

    # 削除されたシートを出力
    if removed_sheets:
        report_lines.append("## 削除されたシート\n")
        for s in removed_sheets:
            report_lines.append(f"- {s}")
            print(f"[INFO] 削除されたシート検出: {s}")
        report_lines.append("")

    # 共通シートを比較
    for sheet in common_sheets:
        print(f"[INFO] シート比較中: {sheet}")
        report_lines.append(f"## シート: {sheet}\n")

        vals1 = data1[sheet]
        vals2 = data2[sheet]

        changed = []
        all_cells = set(vals1.keys()) | set(vals2.keys())
        for pos in sorted(all_cells):
            v1 = vals1.get(pos, None)
            v2 = vals2.get(pos, None)
            if v1 != v2:  # 値に差がある場合のみ記録
                col = colnum_string(pos[1])
                cell = f"{col}{pos[0]}"
                changed.append((cell, v1, v2))

        # 差分がある場合は表形式で出力
        if changed:
            report_lines.append("| セル | 旧値 | 新値 |")
            report_lines.append("|------|------|------|")
            for cell, v1, v2 in changed:
                report_lines.append(f"| {cell} | {v1 if v1 else ''} | {v2 if v2 else ''} |")
                print(f"[DIFF] {sheet} {cell}: '{v1}' → '{v2}'")
            report_lines.append("")
        else:
            report_lines.append("変更なし\n")
            print(f"[INFO] 差分なし: {sheet}")

    # Markdownファイルに書き出し
    with open(output_md, "w", encoding="utf-8") as f:
        f.write("\n".join(report_lines))

    print(f"[INFO] 差分レポートを出力しました: {output_md}")


if __name__ == "__main__":
    # コマンドライン引数からファイル名を取得
    parser = argparse.ArgumentParser(description="2つのExcelファイルを比較してMarkdown形式で差分を出力します。")
    parser.add_argument("file1", help="比較元のExcelファイル")
    parser.add_argument("file2", help="比較先のExcelファイル")
    parser.add_argument("-o", "--output", default="diff_report.md", help="出力Markdownファイル名 (デフォルト: diff_report.md)")
    args = parser.parse_args()

    # 差分比較を実行
    compare_excels(args.file1, args.file2, args.output)
