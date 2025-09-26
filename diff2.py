import openpyxl
import sys
import os
from openpyxl import load_workbook

def excel_diff_report(old_file, new_file, output_md="diff_report.md"):
    print(f"[INFO] Start comparison: {old_file} vs {new_file}")

    # ファイル存在チェック
    if not os.path.exists(old_file):
        print(f"[ERROR] Old file not found: {old_file}")
        return
    if not os.path.exists(new_file):
        print(f"[ERROR] New file not found: {new_file}")
        return

    # Excel読み込み
    print("[INFO] Loading workbooks...")
    old_wb = load_workbook(old_file, data_only=True)
    new_wb = load_workbook(new_file, data_only=True)

    # シート一覧取得
    old_sheets = set(old_wb.sheetnames)
    new_sheets = set(new_wb.sheetnames)

    added_sheets = new_sheets - old_sheets
    removed_sheets = old_sheets - new_sheets
    common_sheets = old_sheets & new_sheets

    print(f"[INFO] Sheets found - Added: {added_sheets}, Removed: {removed_sheets}, Common: {common_sheets}")

    with open(output_md, "w", encoding="utf-8") as f:
        f.write(f"# Excel差分レポート\n\n")
        f.write(f"- 比較元: `{old_file}`\n")
        f.write(f"- 比較先: `{new_file}`\n\n")

        if added_sheets:
            f.write("## 追加されたシート\n\n")
            for s in added_sheets:
                f.write(f"- {s}\n")
            f.write("\n")

        if removed_sheets:
            f.write("## 削除されたシート\n\n")
            for s in removed_sheets:
                f.write(f"- {s}\n")
            f.write("\n")

        for sheet_name in common_sheets:
            print(f"[INFO] Comparing sheet: {sheet_name}")
            f.write(f"## シート: {sheet_name}\n\n")

            old_ws = old_wb[sheet_name]
            new_ws = new_wb[sheet_name]

            max_row = max(old_ws.max_row, new_ws.max_row)
            max_col = max(old_ws.max_column, new_ws.max_column)

            diff_table = []
            long_texts = []

            for r in range(1, max_row + 1):
                for c in range(1, max_col + 1):
                    cell = old_ws.cell(r, c).value
                    new_cell = new_ws.cell(r, c).value
                    if cell != new_cell:
                        coord = f"{old_ws.cell(r, c).coordinate}"
                        print(f"[DEBUG] Difference found at {sheet_name} {coord}")

                        # 長文の場合は別枠に出力
                        if (cell and len(str(cell)) > 50) or (new_cell and len(str(new_cell)) > 50):
                            diff_table.append([coord,
                                f"[旧値はこちら](#{sheet_name}_{coord}_old)",
                                f"[新値はこちら](#{sheet_name}_{coord}_new)"
                            ])
                            long_texts.append((sheet_name, coord, cell, new_cell))
                        else:
                            diff_table.append([coord, str(cell), str(new_cell)])

            if diff_table:
                f.write("| セル | 旧値 | 新値 |\n")
                f.write("|------|------|------|\n")
                for row in diff_table:
                    f.write("| " + " | ".join(row) + " |\n")
                f.write("\n")
            else:
                f.write("変更なし\n\n")

            # 長文別枠出力
            for (sheet, coord, old_val, new_val) in long_texts:
                f.write(f"### {sheet} {coord}\n")
                if old_val is not None:
                    f.write(f"#### <a name=\"{sheet}_{coord}_old\"></a>旧値\n")
                    f.write("```\n" + str(old_val) + "\n```\n\n")
                if new_val is not None:
                    f.write(f"#### <a name=\"{sheet}_{coord}_new\"></a>新値\n")
                    f.write("```\n" + str(new_val) + "\n```\n\n")

    print(f"[INFO] Diff report generated: {output_md}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python excel_diff.py old.xlsx new.xlsx [output.md]")
    else:
        old_file = sys.argv[1]
        new_file = sys.argv[2]
        output_md = sys.argv[3] if len(sys.argv) > 3 else "diff_report.md"
        excel_diff_report(old_file, new_file, output_md)
