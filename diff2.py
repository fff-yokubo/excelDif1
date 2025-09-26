import openpyxl
import argparse

def colnum_string(n):
    """列番号をA1形式に変換 (例: 1 -> A, 27 -> AA)"""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result

def load_excel_as_dict(path):
    """Excelを辞書形式に変換"""
    wb = openpyxl.load_workbook(path, data_only=True)
    sheets_dict = {}
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        values = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    values[(cell.row, cell.column)] = str(cell.value)
        sheets_dict[sheet] = values
    return sheets_dict

def compare_excels(file1, file2, output_md="diff_report.md", threshold=50):
    """2つのExcelファイルを比較してMarkdown差分を出力"""
    data1 = load_excel_as_dict(file1)
    data2 = load_excel_as_dict(file2)

    report_lines = []
    report_lines.append(f"# Excel差分レポート\n")
    report_lines.append(f"- 比較元: `{file1}`\n- 比較先: `{file2}`\n")

    sheets1 = set(data1.keys())
    sheets2 = set(data2.keys())

    added_sheets = sheets2 - sheets1
    removed_sheets = sheets1 - sheets2
    common_sheets = sheets1 & sheets2

    if added_sheets:
        report_lines.append("## 追加されたシート\n")
        for s in added_sheets:
            report_lines.append(f"- {s}")
        report_lines.append("")

    if removed_sheets:
        report_lines.append("## 削除されたシート\n")
        for s in removed_sheets:
            report_lines.append(f"- {s}")
        report_lines.append("")

    for sheet in common_sheets:
        report_lines.append(f"## シート: {sheet}\n")

        vals1 = data1[sheet]
        vals2 = data2[sheet]

        changed = []
        long_text_blocks = []

        all_cells = set(vals1.keys()) | set(vals2.keys())
        for pos in sorted(all_cells):
            v1 = vals1.get(pos, None)
            v2 = vals2.get(pos, None)
            if v1 != v2:
                col = colnum_string(pos[1])
                cell = f"{col}{pos[0]}"

                need_long_block = (v1 and len(v1) > threshold) or (v2 and len(v2) > threshold)
                anchor = f"{sheet}_{cell}".replace(" ", "_")

                if need_long_block:
                    v1_display = f"[旧値はこちら](#{anchor}_old)"
                    v2_display = f"[新値はこちら](#{anchor}_new)"
                    long_text_blocks.append((sheet, cell, v1 or "", v2 or "", anchor))
                else:
                    v1_display = v1 or ""
                    v2_display = v2 or ""

                changed.append((cell, v1_display, v2_display))

        if changed:
            report_lines.append("| セル | 旧値 | 新値 |")
            report_lines.append("|------|------|------|")
            for cell, v1, v2 in changed:
                report_lines.append(f"| {cell} | {v1} | {v2} |")
            report_lines.append("")
        else:
            report_lines.append("変更なし\n")

        # 長文セルを縦に別枠表示（旧値・新値それぞれ見出しレベル4）
        for sheet_name, cell, v1, v2, anchor in long_text_blocks:
            report_lines.append(f"### {sheet_name} {cell}\n")
            report_lines.append(f"#### <a name=\"{anchor}_old\"></a>旧値\n")
            report_lines.append("```")
            report_lines.append(v1)
            report_lines.append("```")
            report_lines.append(f"#### <a name=\"{anchor}_new\"></a>新値\n")
            report_lines.append("```")
            report_lines.append(v2)
            report_lines.append("```")
            report_lines.append("")

    with open(output_md, "w", encoding="utf-8") as f:
        f.write("\n".join(report_lines))

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Excel差分比較ツール")
    parser.add_argument("file1", help="比較元のExcelファイル")
    parser.add_argument("file2", help="比較先のExcelファイル")
    parser.add_argument("-o", "--output", default="diff_report.md", help="出力Markdownファイル")
    parser.add_argument("-t", "--threshold", type=int, default=50, help="長文閾値")
    args = parser.parse_args()

    compare_excels(args.file1, args.file2, args.output, args.threshold)
