import pandas as pd
import random
import string

def shuffle_choices(df):
    new_rows = []
    i = 0
    n = len(df)

    while i < n:
        code = df.iloc[i, 0]

        # --- b_select ブロック開始 ---
        if code == "b_select":
            new_rows.append(df.iloc[i])  # b_select
            i += 1

            # --- 元の選択肢を読み取る ---
            original_selects = []
            while df.iloc[i, 0] == "select":
                original_selects.append(df.iloc[i].copy())
                i += 1

            # e_select 行
            e_select_row = df.iloc[i].copy()
            i += 1

            # --- b_answer 部分 ---
            new_rows.append(df.iloc[i])  # b_answer
            i += 1

            ans_row = df.iloc[i].copy()
            original_answer = ans_row[1]  # 例: "C"
            i += 1

            e_answer_row = df.iloc[i].copy()
            i += 1

            # --- 元の正解のインデックスを求める ---
            old_labels = list(string.ascii_uppercase)
            correct_index_original = old_labels.index(original_answer)

            # --- ランダムシャッフルで正解が同じ位置にならないようにする ---
            while True:
                shuffled = original_selects.copy()
                random.shuffle(shuffled)

                # 新しい正解のインデックスを探す
                correct_text = original_selects[correct_index_original][1]
                new_correct_index = next(
                    idx for idx, row in enumerate(shuffled) if row[1] == correct_text
                )

                # 🔥 位置が同じなら再シャッフル、違えばOK 🔥
                if new_correct_index != correct_index_original:
                    break

            # --- 新しいラベルを付ける ---
            labels = list(string.ascii_uppercase)
            for idx, row in enumerate(shuffled):
                row["label"] = labels[idx]
                new_rows.append(row)

            new_rows.append(e_select_row)

            # --- answer を新ラベルへ置き換え ---
            new_answer_label = labels[new_correct_index]
            ans_row[1] = new_answer_label

            new_rows.append(ans_row)
            new_rows.append(e_answer_row)

        else:
            # その他の行はそのまま
            new_rows.append(df.iloc[i])
            i += 1

    return pd.DataFrame(new_rows)

# ====== 実行例 ======
df = pd.read_excel("exam.xlsx", header=None)
new_df = shuffle_choices(df)
new_df.to_excel("exam_shuffled.xlsx", index=False)
