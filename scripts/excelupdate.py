import xlwings as xw
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from pathlib import Path
import sys
from utils import get_exam_path

# --- 設定：装飾スタイル (RGB) ---
COMMENT_FILL_RGB = (255, 230, 153)  # FFE699
COMMENT_FONT_BOLD = True

def output_summary_v2(stats, comment_cnt, sheet_name, log_dir=None):
    """結果をポップアップ表示し、ログファイルに記録する"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    msg = (
        f"【処理結果レポート】 ({now})\n"
        f"対象シート: {sheet_name}\n"
        f"{'-'*30}\n"
        f"・問題数: {stats['total_questions']}\n"
        f"・小問題数: {stats['total_subquestions']}\n"
        f"・得点合計: {stats['total_points']} 点\n"
        f"・改ページ数: {stats['pagebreaks']}\n"
        f"・コメント挿入/更新: {comment_cnt} 箇所\n"
        f"・QID更新: {stats['qid_updates']} 箇所\n"
        f"{'-'*30}\n"
        "更新が完了しました（保存は手動で行ってください）。"
    )

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showinfo("実行完了", msg)
    root.destroy()

    if log_dir is None:
        log_dir = Path(__file__).parent
    else:
        log_dir = Path(log_dir)

    log_dir.mkdir(parents=True, exist_ok=True)

    log_file = log_dir / "process_history.log"

    with open(log_file, "a", encoding="utf-8") as f:
        f.write(msg + "\n\n")

    print(f"✅ 画面に結果を表示しました。履歴: {log_file}")

def _output_summary_v2(stats, comment_cnt, sheet_name):
    """結果をポップアップ表示し、ログファイルに記録する"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    msg = (
        f"【処理結果レポート】 ({now})\n"
        f"対象シート: {sheet_name}\n"
        f"{'-'*30}\n"
        f"・問題数: {stats['total_questions']}\n"
        f"・小問題数: {stats['total_subquestions']}\n"
        f"・得点合計: {stats['total_points']} 点\n"
        f"・改ページ数: {stats['pagebreaks']}\n"
        f"・コメント挿入/更新: {comment_cnt} 箇所\n"
        f"・QID更新: {stats['qid_updates']} 箇所\n"
        f"{'-'*30}\n"
        "更新が完了しました（保存は手動で行ってください）。"
    )

    # 1. ポップアップ表示
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)  # 最前面に表示
    messagebox.showinfo("実行完了", msg)
    root.destroy()

    # 2. ログファイル出力
    log_file = Path(__file__).parent / "process_history.log"
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(msg + "\n\n")
    
    print(f"✅ 画面に結果を表示しました。履歴: {log_file}")

#def update_active_sheet():
def update_active_sheet(log_dir=None):
    """現在アクティブなExcelシートを直接更新する"""
    try:
        # 現在アクティブなブックとシートを取得
        wb = xw.books.active
        ws = wb.sheets.active
        print(f"🚀 接続中: [{wb.name}] {ws.name}")
    except Exception as e:
        print(f"❌ Excelが見つかりません。ファイルを開いてから実行してください。: {e}")
        return

    # ---- 1パス目：データの走査と集計 ----
    # パフォーマンスのため、A列からC列のデータを一括取得
    last_row = ws.range('A' + str(ws.cells.last_cell.row)).end('up').row
    # values は (A, B, C) のタプルのリストになる
    all_values = ws.range(f'A1:C{last_row}').value

    actions = []  # (row_index, label_text)
    qnum = 0
    sub_index = 0
    stats = {
        "total_questions": 0,
        "total_subquestions": 0,
        "total_points": 0,
        "pagebreaks": 0,
        "qid_updates": 0
    }

    for i, row_data in enumerate(all_values):
        row_idx = i + 1
        tag = str(row_data[0] or "").strip() # A列
        
        if tag == "e_exam":
            print(f"🔚 e_exam 到達 (Row {row_idx})")
            break

        if tag == "b_question":
            qnum += 1
            stats["total_questions"] += 1
            sub_index = 0
            
            # QID (C列) の更新
            qid = f"Q{qnum:03d}"
            current_qid = str(row_data[2] or "").strip()
            if current_qid != qid:
                ws.range(f'C{row_idx}').number_format = "@"
                ws.range(f'C{row_idx}').value = qid
                stats["qid_updates"] += 1

            actions.append((row_idx, f"# 【問題{qnum}】"))

        elif tag == "b_subquest":
            sub_index += 1
            stats["total_subquestions"] += 1
            actions.append((row_idx, f"# 【問題{qnum}-{sub_index}】"))

        elif tag in ["answer", "subanswer"]:
            # C列の得点を加算
            val = row_data[2]
            if isinstance(val, (int, float)):
                stats["total_points"] += int(val)

        elif tag == "PAGEBREAK":
            stats["pagebreaks"] += 1

#     # ---- 2パス目：行挿入とコメント更新（下から適用） ----
#     line_str = "-" * 100
#     comment_change_cnt = 0

#     for row_idx, text1 in reversed(actions):
#         # 直前行に既にコメントがあるか確認
#         prev_row = row_idx - 1
#         if prev_row >= 1:
#             prev_val = str(ws.range(f'A{prev_row}').value or "").strip()
#             if prev_val.startswith("# 【問題"):
#                 # 既存コメントの更新
#                 if prev_val != text1:
#                     ws.range(f'A{prev_row}').value = text1
#                     comment_change_cnt += 1
#                 continue

#         # 新規挿入
# ###        ws.api.Rows(row_idx).Insert()
#         ws.range(f"{row_idx}:{row_idx}").insert(shift='down')
#         target_range = ws.range(f'A{row_idx}:B{row_idx}')
#         target_range.value = [text1, line_str]
        
#         # 装飾
#         target_range.color = COMMENT_FILL_RGB
#         target_range.api.Font.Bold = COMMENT_FONT_BOLD
#         ws.range(f'B{row_idx}').number_format = "@"
#         comment_change_cnt += 1


# ---- 2パス目：行挿入とコメント更新（下から適用） ----
    line_str = "-" * 200
    comment_change_cnt = 0

    for row_idx, text1 in reversed(actions):
        # 直前行に既にコメントがあるか確認
        prev_row = row_idx - 1
        if prev_row >= 1:
            prev_val = str(ws.range(f'A{prev_row}').value or "").strip()
            if prev_val.startswith("# 【問題"):
                # 既存コメントの更新
                if prev_val != text1:
                    ws.range(f'A{prev_row}').value = text1
                    comment_change_cnt += 1
                continue

        # --- 【修正】行挿入（Mac/Win共通） ---
        ws.range(f"{row_idx}:{row_idx}").insert(shift='down')
        
        target_range = ws.range(f'A{row_idx}:B{row_idx}')
        target_range.value = [text1, line_str]
        
        # --- 【修正】装飾（.api を使わず標準プロパティを使用） ---
        target_range.color = COMMENT_FILL_RGB
        target_range.font.bold = COMMENT_FONT_BOLD # api.Font.Bold ではなく font.bold
        
        # --- 【修正】表示形式 ---
        ws.range(f'B{row_idx}').number_format = "@"
        comment_change_cnt += 1




    # 結果の表示と保存
    #output_summary_v2(stats, comment_change_cnt, ws.name)
    output_summary_v2(stats, comment_change_cnt, ws.name, log_dir=log_dir)

if __name__ == "__main__":
    subject_no = sys.argv[1] if len(sys.argv) >= 2 else None
    target_year = sys.argv[2] if len(sys.argv) >= 3 else "2026"

    log_dir = None

    if subject_no:
        excel_path, work_dir, exam_koma_no, sub_folder = get_exam_path(subject_no, target_year)
        log_dir = work_dir

        print(f"科目番号: {subject_no}")
        print(f"年度: {target_year}")
        print(f"試験コマ番号: {exam_koma_no}")
        print(f"想定Excel: {excel_path}")
        print(f"ログ出力先: {log_dir}")

    update_active_sheet(log_dir=log_dir)