import json
import openpyxl
from pathlib import Path
import sys
import re

# 既存プロジェクトの共通ユーティリティ（v1と同じ）
from utils import setspace, parse_with_number
from utils import calc_excel_hash
from versioncontrol import ensure_version_entry

# v2: qpattern は b_exam ブロックの qpattern 行から取得（v1同様）
qpattern = None


# -------------------------
# v2: Backslash display conversion
# -------------------------
# Excelセル中に「\0」など「バックスラッシュ + 数字」が含まれる場合、
# LaTeX表示でバックスラッシュ記号として出すために
# \textbackslash 0 のように変換する（JSON段階で整形）
_BACKSLASH_DIGIT_RE = re.compile(r"\\(?=\d)")

def conv_text(v) -> str:
    """
    Convert textual cells for LaTeX-friendly backslash display.
      - '\0', '\12' ... -> '\textbackslash 0', '\textbackslash 12'
      - If a real NUL character (0x00) is present, treat it as '\0'
    Notes:
      - This is applied ONLY to plain-text fields (question text, choices text, notes, etc.)
      - Do NOT apply to code blocks or file paths.
    """
    if v is None:
        return ""
    s = str(v).strip()
    if s == "":
        return ""
    # safety: real NUL (rare in Excel, but possible after processing)
    s = s.replace("\x00", r"\textbackslash 0")
    # backslash before digit(s)
    s = _BACKSLASH_DIGIT_RE.sub(r"\\textbackslash ", s)
    return s


def _parse_select_style(style_raw: str):
    """
    v2:
      - "normal" -> ("normal", None)
      - "inline" -> ("inline", None)          # sep unspecified => TeX default
      - "inline(8)" -> ("inline", 8.0)       # sep specified (number)
    """
    s = (style_raw or "").strip()
    if s == "":
        return ("inline", None)

    if s == "normal":
        return ("normal", None)

    if s == "inline":
        return ("inline", None)

    if s.startswith("inline(") and s.endswith(")"):
        inner = s[len("inline("):-1].strip()
        # 数値として解釈（intでもfloatでも可）
        try:
            return ("inline", float(inner))
        except ValueError:
            # ここは運用上エラーにしてよい（曖昧にしない）
            raise ValueError(f"b_select style の inline(...) が数値ではありません: {style_raw}")

    # 想定外は一旦エラー（仕様の揺れを防ぐ）
    raise ValueError(f"b_select style が不正です: {style_raw}")


def _to_int_or_none(v):
    if v is None:
        return None
    s = str(v).strip()
    if s == "":
        return None
    try:
        return int(s)
    except:
        return None


def _to_float_or_none(v):
    if v is None:
        return None
    s = str(v).strip()
    if s == "":
        return None
    try:
        return float(s)
    except:
        return None


def _is_one(v) -> bool:
    """v2: PB は 1 だけを ON 扱い（ユーザー指定）"""
    if v is None:
        return False
    try:
        return int(str(v).strip()) == 1
    except:
        return False


def _sort_questions_for_B(questions):
    """
    version=B: b_question の D列 orderB（int）で並び替える
    - 未入力があればエラー（必須運用）
    - 重複もエラー
    """
    orders = []
    for q in questions:
        ob = q.get("_orderB")
        if ob is None:
            raise ValueError(f"orderB 未入力の大問があります（qid={q.get('qid')}, 元番号={q.get('number')}）")
        orders.append(ob)

    if len(set(orders)) != len(orders):
        # どれが重複しているか簡易で出す
        dup = sorted([x for x in set(orders) if orders.count(x) > 1])
        raise ValueError(f"orderB が重複しています: {dup}")

    return sorted(questions, key=lambda q: q["_orderB"])


def excel_to_json_v2(ws, version="A"):
    """
    v2:
      - A版: Excel出現順（=orderA）で出力
      - B版: b_question D列(orderB)で並べ替え、問番号は振り直す
      - PB_*_after / LS_*_after は「その大問の直後」に root 配列へ pagebreak / vspace を挿入
      - LINESPACE タグは subgroup 内のみ（validateで保証）→ vspace に変換
      - PAGEBREAK タグは v2 では使わない想定（validateで禁止）
    """
    global qpattern

    cover = None
    current_q = None
    current_sub = None
    current_select = None
    current_code = None
    current_exam = None
    current_multiline = None
    current_subgroup = None

    q_number = 0
    sub_number = 0

    questions_raw = []  # 大問だけを集める（並べ替え対象）
    # cover は最後に先頭へ

    for row in ws.iter_rows(values_only=True):
        if not row or not row[0]:
            continue

        tag = str(row[0]).strip()
        # v1同様: B列以降は params として利用（必要箇所は row で参照）
        raw_params = [("" if x is None else str(x)) for x in row[1:]]     # ★生のまま
        params = [str(x).strip() if x is not None else "" for x in row[1:]]

        # -------------------------
        # exam ヘッダ
        # -------------------------
        if tag == "b_exam":
            current_exam = {
                "type": "cover",
                "title": None,
                "notes": [],
                "subject": None,
                "fsyear": None,
                "ansnote": None,
                "anssize": None,
            }

        elif tag == "examtitle" and current_exam is not None:
            current_exam["title"] = conv_text(params[0])

        elif tag == "b_examnote" and current_exam is not None:
            # note 行は examnote で積む
            pass

        elif tag == "examnote" and current_exam is not None:
            if params and params[0]:
                current_exam["notes"].append(conv_text(params[0]))

        elif tag == "e_examnote" and current_exam is not None:
            pass

        elif tag == "subject" and current_exam is not None:
            current_exam["subject"] = params[0]

        elif tag == "fsyear" and current_exam is not None:
            current_exam["fsyear"] = params[0]

        elif tag == "ansnote" and current_exam is not None:
            current_exam["ansnote"] = conv_text(params[0])

        elif tag == "qpattern":
            # 例: "A,B"
            qpattern = [x.strip() for x in params[0].split(",") if x.strip()]

        elif tag == "anssize" and current_exam is not None:
            # 既存 setspace を流用（"(50,60)" → [50.0,60.0]）
            from utils import setspace as _setspace
            current_exam["anssize"] = list(_setspace(params[0], "ANSSIZE"))

        elif tag == "e_exam" and current_exam is not None:
            cover = current_exam
            current_exam = None
            # ここで break しない（v1では break していたが、念のため続きも読めるようにする）
            continue

        # -------------------------
        # 大問
        # -------------------------
        elif tag == "b_question":
            q_number += 1
            sub_number = 0

            # v2列: C=qid, D=orderB, E=PB_A_after, F=LS_A_after, G=PB_B_after, H=LS_B_after
            # v2: qid は必須（make_numberv2.py で付与する運用）
            if len(row) >= 3 and row[2] is not None and str(row[2]).strip():
                qid = str(row[2]).strip()
            else:
                raise ValueError(f"qid 未入力の大問があります（b_question 行: 元番号={q_number}）")

            orderB = _to_int_or_none(row[3] if len(row) >= 4 else None)
            pb_a = row[4] if len(row) >= 5 else None
            ls_a = row[5] if len(row) >= 6 else None
            pb_b = row[6] if len(row) >= 7 else None
            ls_b = row[7] if len(row) >= 8 else None

            # versionに応じて after を選択
            pb_after = _is_one(pb_a) if version == "A" else _is_one(pb_b)
            ls_after = _to_float_or_none(ls_a if version == "A" else ls_b)

            current_q = {
                "number": str(q_number),  # B版は後で振り直す
                "qid": qid,               # v2追加（ユーザー指定）
                "question": None,
                "content": [],
                "subquestions": [],
                "_orderB": orderB,
                "_pb_after": pb_after,
                "_ls_after": ls_after,
            }

        elif tag == "question" and current_q is not None:
            current_q["question"] = conv_text(params[0])

        # 単独行（大問内）
        elif tag == "sline" and current_q is not None:
            wspace = setspace(row[2], "SPACEB_A")
            sline = {"type": "text", "values": [conv_text(params[0])], "space_before": wspace[0], "space_after": wspace[1]}
            current_q["content"].append(sline)

        # 複数行（大問内）
        elif tag == "b_multiline" and current_q is not None:
            wspace = setspace(row[2], "SPACEB_A")
            current_multiline = {"type": "multiline", "values": [], "space_before": wspace[0], "space_after": wspace[1]}

        elif tag == "text" and current_multiline is not None:
            current_multiline["values"].append(conv_text(params[0]))

        elif tag == "e_multiline" and current_multiline is not None:
            current_q["content"].append(current_multiline)
            current_multiline = None

        # 画像（大問内）
        elif tag == "image" and current_q is not None:
            wimg = parse_with_number(params[0],0.85)
            img = {"type": "image", "path": wimg[0], "width": wimg[1]}
            current_q["content"].append(img)

        # 大問コード
        elif tag == "b_code":
            flg = (params[0] == "linenumber") if params else False
            current_code = {"type": "code", "linenumber": flg, "lines": []}

        elif tag == "code" and current_code is not None:
            current_code["lines"].append(raw_params[0])

        elif tag == "e_code" and current_code is not None and current_q is not None:
            current_q["content"].append(current_code)
            current_code = None

        # 選択肢ブロック（大問）
        elif tag == "b_select" and current_q is not None:
            style_raw = params[0] if params else "inline"
            sty, sep = _parse_select_style(style_raw)

            select_block = {"type": "choices", "style": sty, "values": []}
            # 例外のときだけ sep を入れる（=TeX側のデフォルトを活かす）
            if sep is not None:
                select_block["sep"] = sep

            current_select = select_block

        # elif tag == "b_select" and current_q is not None:
        #     style = params[0] if params else "inline"
        #     select_block = {"type": "choices", "style": "inline", "values": []}

        #     if style == "normal":
        #         select_block["style"] = "normal"
        #     elif style.startswith("inline(") and style.endswith(")"):
        #         try:
        #             space = int(style[len("inline("):-1])
        #         except ValueError:
        #             space = 8
        #         select_block["space"] = space
        #     else:
        #         select_block["space"] = 8

        #     current_select = select_block

        elif tag == "select" and current_select is not None:
            order_num = None
            if len(row) > 6 and row[6]:
                try:
                    order_num = int(row[6])
                except:
                    order_num = None

            current_select["values"].append({
                "label": None if len(params) < 2 else params[1],
                "text": conv_text(params[0]),
                "order": order_num
            })

        elif tag == "e_select" and current_select is not None:
            if version != "A":
                order_list = [v.get("order") for v in current_select["values"]]
                if any(order_list):
                    current_select["values"].sort(key=lambda v: (v["order"] if v["order"] is not None else 999))

            for i, v in enumerate(current_select["values"]):
                if not v.get("label"):
                    v["label"] = chr(ord("A") + i)
                v.pop("order", None)

            current_q["content"].append(current_select)
            current_select = None

        # -------------------------
        # 小問グループ
        # -------------------------
        elif tag == "b_subgroup":
            current_subgroup = {"subquestions": []}

        elif tag == "b_subquest":
            sub_number += 1
            current_sub = {"number": str(sub_number), "question": None, "content": []}

        elif tag == "subquest" and current_sub is not None:
            current_sub["question"] = conv_text(params[0])

        elif tag == "subsline" and current_sub is not None:
            wspace = setspace(row[2], "SPACEB_A")
            subsline = {"type": "text", "values": [conv_text(params[0])], "space_before": wspace[0], "space_after": wspace[1]}
            current_sub["content"].append(subsline)

        elif tag == "b_submultiline" and current_sub is not None:
            wspace = setspace(row[2], "SPACEB_A")
            current_multiline = {"type": "multiline", "values": [], "space_before": wspace[0], "space_after": wspace[1]}

        elif tag == "subtext" and current_multiline is not None:
            current_multiline["values"].append(conv_text(params[0]))

        elif tag == "e_submultiline" and current_multiline is not None:
            current_sub["content"].append(current_multiline)
            current_multiline = None

        elif tag == "subimage" and current_sub is not None:
            wimg = parse_with_number(params[0],0.85)
            img = {"type": "image", "path": wimg[0], "width": wimg[1]}
            current_sub["content"].append(img)

        # sub code
        elif tag == "b_subcode" and current_sub is not None:
            flg = (params[0] == "linenumber") if params else False
            current_code = {"type": "code", "linenumber": flg, "lines": []}

        elif tag == "subcode" and current_code is not None:
            current_code["lines"].append(raw_params[0])

        elif tag == "e_subcode" and current_code is not None:
            current_sub["content"].append(current_code)
            current_code = None

        # subselect
        elif tag == "b_subselect" and current_sub is not None:
            style_raw = params[0] if params else "inline"
            sty, sep = _parse_select_style(style_raw)

            select_block = {"type": "choices", "style": sty, "values": []}
            if sep is not None:
                select_block["sep"] = sep

            current_select = select_block

        # elif tag == "b_subselect" and current_sub is not None:
        #     style = params[0] if params else "inline"
        #     select_block = {"type": "choices", "style": "inline", "values": []}

        #     if style == "normal":
        #         select_block["style"] = "normal"
        #     elif style.startswith("inline(") and style.endswith(")"):
        #         try:
        #             space = int(style[len("inline("):-1])
        #         except ValueError:
        #             space = 8
        #         select_block["space"] = space
        #     else:
        #         select_block["space"] = 8

        #     current_select = select_block

        elif tag == "subselect" and current_select is not None:
            order_num = None
            if len(row) > 6 and row[6]:
                try:
                    order_num = int(row[6])
                except:
                    order_num = None

            current_select["values"].append({
                "label": None if len(params) < 2 else params[1],
                "text": conv_text(params[0]),
                "order": order_num
            })

        elif tag == "e_subselect" and current_select is not None:
            if version != "A":
                order_list = [v.get("order") for v in current_select["values"]]
                if any(order_list):
                    current_select["values"].sort(key=lambda v: (v["order"] if v["order"] is not None else 999))

            for i, v in enumerate(current_select["values"]):
                if not v.get("label"):
                    v["label"] = chr(ord("A") + i)
                v.pop("order", None)

            current_sub["content"].append(current_select)
            current_select = None

        # 小問終了
        elif tag == "e_subquest":
            current_subgroup["subquestions"].append(current_sub)
            current_sub = None

        # subgroup 終了 → 大問へ合流
        elif tag == "e_subgroup":
            if current_q is not None:
                current_q["subquestions"].extend(current_subgroup["subquestions"])
            current_subgroup = None

        # v2: LINESPACE は subgroup 内のみ → vspace に変換
        elif tag == "LINESPACE":
            # v2: {"type":"vspace","value": float}（負数・小数可）
            v = _to_float_or_none(params[0] if params else None)
            if v is None:
                # 空ならデフォルト 1.0 とせず、明示しない（validateで弾く想定）
                raise ValueError("LINESPACE の値が空です（v2では数値必須）")
            vspace = {"type": "vspace", "value": v}
            if current_subgroup is not None:
                current_subgroup["subquestions"].append(vspace)
            else:
                # v2運用外（validateで弾く想定）
                raise ValueError("LINESPACE は subgroup 外では使用できません（v2）")

        # v2: PAGEBREAK 行タグは使わない
        elif tag == "PAGEBREAK":
            raise ValueError("PAGEBREAK 行タグは v2 では使用しません（PB_*_after を使用してください）")

        # 大問終了
        elif tag == "e_question":
            if current_q is not None:
                if not current_q["subquestions"]:
                    current_q.pop("subquestions")
                questions_raw.append(current_q)
            current_q = None

        # その他タグは v1同様に無視/拡張余地
        else:
            pass

    if cover is None and current_exam is not None:
        cover = current_exam

    # -------------------------
    # version別の並び替えと、PB/LS after の挿入
    # -------------------------
    # v2: qid の重複チェック（重複は不可）
    qids = [q.get("qid") for q in questions_raw]
    if len(set(qids)) != len(qids):
        dup = sorted([x for x in set(qids) if qids.count(x) > 1])
        raise ValueError(f"qid が重複しています: {dup}")

    if version == "B":
        questions = _sort_questions_for_B(questions_raw)
    else:
        questions = questions_raw

    # v2: 問番号を振り直す（B版は必須。A版も再付番して整合させる）
    for i, q in enumerate(questions, start=1):
        q["number"] = str(i)

    results = []
    if cover is not None:
        results.append(cover)

    for q in questions:
        pb_after = q.pop("_pb_after", False)
        ls_after = q.pop("_ls_after", None)
        q.pop("_orderB", None)  # JSONに残さない

        results.append(q)

        # v2: after 挿入（root配列）
        if ls_after is not None:
            results.append({"type": "vspace", "value": ls_after})
        if pb_after:
            results.append({"type": "pagebreak"})

    return results


if __name__ == "__main__":
    """
    使い方:
      python makedocjsonv2.py <sheetname> [excel_filename]
    例:
      python makedocjsonv2.py 2022001 試験問題.xlsx
    """
    if len(sys.argv) < 2:
        print("Usage: python makedocjsonv2.py <sheetname> [excel_filename]")
        sys.exit(1)

    curdir = Path(__file__).parent.parent
    sheetname = sys.argv[1]
    examdata = "試験問題.xlsx" if len(sys.argv) < 3 else sys.argv[2]
    excel_path = curdir / "input" / examdata

    wb = openpyxl.load_workbook(excel_path)
    ws = wb[sheetname]

    # hash / version 管理（v1と同じ）
    ehash = calc_excel_hash(ws)
    wver = ensure_version_entry(ehash, str(excel_path), sheetname)

    # qpattern を読む（v1と同じ）
    _ = excel_to_json_v2(ws, version="A")
    qp = qpattern if qpattern else ["A"]

    versionmode = "single" if (not qp or qp == ["A"]) else "multi"
    outjson = {"versionmode": versionmode, "versions": []}

    for v in qp:
        questions = excel_to_json_v2(ws, version=v)

        block = {
            "version": v,
            "questions": questions,
            "metainfo": {
                "type": "metainfo",
                "hash": ehash,
                "createdatetime": str(wver["createdatetime"]) if isinstance(wver, dict) and "createdatetime" in wver else "",
                "verno": wver.get("verno") if isinstance(wver, dict) else None,
                "inputpath": str(excel_path),
                "sheetname": sheetname,
            },
        }
        outjson["versions"].append(block)

    out = curdir / "work" / f"{sheetname}.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        json.dump(outjson, f, ensure_ascii=False, indent=2)

    print(f"✅ jsonファイルを作成しました: {out}")