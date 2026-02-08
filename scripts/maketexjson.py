import json
import openpyxl
from pathlib import Path
import sys
import re
import subprocess
from datetime import datetime
from typing import Optional

# 既存プロジェクトの共通ユーティリティ（v1と同じ）
from utils import setspace, parse_with_number
from utils import calc_excel_hash
from versioncontrol import ensure_version_entry

from contract import normalize_document, validate_document, ContractError

# v2: qpattern は b_exam ブロックの qpattern 行から取得（v1同様）
qpattern = None

def escape_tex_outside_inline_math(s: str) -> str:
    """
    \(...\) / \[...\] の中はそのまま、
    それ以外の部分だけ { } & をエスケープする。
    """
    out = []
    i = 0
    in_math = False
    end_token = None  # r"\)" or r"\]"

    while i < len(s):
        if not in_math:
            # math start: \(  or \[
            if s.startswith(r"\(", i):
                in_math = True
                end_token = r"\)"
                out.append(r"\(")
                i += 2
                continue
            if s.startswith(r"\[", i):
                in_math = True
                end_token = r"\]"
                out.append(r"\[")
                i += 2
                continue

            ch = s[i]
            if ch == "{":
                out.append(r"\{")
            elif ch == "}":
                out.append(r"\}")
            elif ch == "&":
                out.append(r"\&")
            else:
                out.append(ch)
            i += 1
        else:
            # math end
            if end_token and s.startswith(end_token, i):
                in_math = False
                out.append(end_token)
                i += 2
                end_token = None
                continue

            # 数式中は一切触らない
            out.append(s[i])
            i += 1

    return "".join(out)

# -------------------------
# v2: Backslash display conversion
# -------------------------
# Excelセル中に「\0」など「バックスラッシュ + 数字」が含まれる場合、
# LaTeX表示でバックスラッシュ記号として出すために
# \textbackslash 0 のように変換する（JSON段階で整形）
_BACKSLASH_DIGIT_RE = re.compile(r"\\(?=\d)")

# 追加：クォート内の \n \t \r \0 など（\ + 英数字）を対象にする
_BACKSLASH_IN_QUOTES_RE = re.compile(r"(?<=[\'\"])\\(?=[A-Za-z0-9])")
_RAW_TEX_RE = re.compile(r"\[\[(.+?)\]\]", re.DOTALL)

def _conv_plain_segment(seg: str) -> str:
    # safety: real NUL
    seg = seg.replace("\x00", r"\textbackslash 0")
    # '\n' や "\n" のようなクォート内の \n \t \r \0 など
    seg = _BACKSLASH_IN_QUOTES_RE.sub(r"\\textbackslash ", seg)
    # \0, \12 ...（バックスラッシュ + 数字）
    seg = _BACKSLASH_DIGIT_RE.sub(r"\\textbackslash ", seg)
    # 数式 \( ... \) の外側だけ { } & をエスケープ
    seg = escape_tex_outside_inline_math(seg)
    return seg

def conv_text(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if s == "":
        return ""

    out = []
    last = 0
    for m in _RAW_TEX_RE.finditer(s):
        # 通常部分は変換
        out.append(_conv_plain_segment(s[last:m.start()]))

        # [[...]] の中は「そのまま」通す（[[ ]] 自体は剥がす）
        out.append(f"[[{m.group(1)}]]")

        last = m.end()

    # 残り
    out.append(_conv_plain_segment(s[last:]))

    return "".join(out)


def _parse_select_style(style_raw: str):
    """
    v2 canonical:
      - "normal" -> ("normal", None)
      - "inline" -> ("inline", None)        # sep unspecified => TeX default
      - "inline(8)" -> ("inline", 8)        # sep is int (mm)
    """
    s = (style_raw or "").strip()
    if s == "":
        return ("inline", None)

    if s == "normal":
        return ("normal", None)

    if s == "inline":
        return ("inline", None)

    # 空白も許容: inline( 8 )
    m = re.match(r"^inline\s*\(\s*(\d+)\s*\)\s*$", s)
    if m:
        return ("inline", int(m.group(1)))

    raise ValueError(f"b_select style が不正です: {style_raw}")

# --- label scheme helpers (for choices shuffle) ---
# Excel C column can be used for custom label style like "①". When choices are shuffled
# (using Excel G column "order"), we want labels to be re-numbered in display order so
# they don't look like "②,④,①,③".

def _detect_label_scheme(labels):
    # labels: list of str/None
    first = None
    for x in labels:
        if x is None:
            continue
        s = str(x).strip()
        if s:
            first = s
            break
    if not first:
        return 'alpha'

    # circled ①(U+2460) .. ⑳(U+2473)
    o = ord(first[0]) if len(first) == 1 else -1
    if 0x2460 <= o <= 0x2473:
        return 'circled'

    if len(first) == 1 and 'A' <= first <= 'Z':
        return 'alpha'

    if first.isdigit():
        return 'digit'

    return 'custom'


def _make_labels(scheme: str, n: int):
    if n <= 0:
        return []

    if scheme == 'circled':
        out = []
        for i in range(n):
            if i < 20:
                out.append(chr(0x2460 + i))
            else:
                out.append(str(i + 1))
        return out

    if scheme == 'digit':
        return [str(i + 1) for i in range(n)]

    # default alpha
    labels = []
    for i in range(n):
        if i < 26:
            labels.append(chr(ord('A') + i))
        else:
            labels.append(str(i + 1))
    return labels

import re

def parse_before_after(cell_value, *, default=(0, 0)):
    """
    "(5,0)" "( 5 , 0 )" "" None を許容
    戻り値は (before:int, after:int)
    """
    if cell_value is None:
        return default
    s = str(cell_value).strip()
    if s == "":
        return default

    m = re.match(r"^\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)$", s)
    if not m:
        raise ValueError(f"(before,after) の形式が不正です: {cell_value}")
    return int(m.group(1)), int(m.group(2))

def make_vspace(mm, *, tag, src):
    return {"type": "vspace", "value_mm": int(mm), "tag": tag, "src": src}


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

    追加（デバッグ用）:
      - JSON要素に tag / src を付与する
        - tag: Excel上のタグ名
        - src: {sheet,row[,row_end]}
      - ブロック（b_...〜e_...）は row_end を付与する
    """
    global qpattern

    sheetname = getattr(ws, "title", "") or ""

#    def make_src(row: int, row_end: int): # | None = None):
#    def make_src(row: int, row_end=None):
    def make_src(row: int, row_end: Optional[int] = None):

        src = {"sheet": sheetname, "row": row}
        if row_end is not None:
            src["row_end"] = row_end
        return src

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

    for row_i, row_cells in enumerate(ws.iter_rows(), start=1):
        row = [c.value for c in row_cells]
        if not row or not row[0]:
            continue

        tag = str(row[0]).strip()

        # --- 解答用紙専用タグは、問題用紙JSONでは無視 ---
        if tag in ("#ansbreak"):
            continue

        # v1同様: B列以降は params として利用（必要箇所は row で参照）
        raw_params = [("" if x is None else str(x)) for x in row[1:]]     # ★生のまま（code用）
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
                "tag": tag,
                "src": make_src(row_i),
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
            # cover の src をブロック範囲にする
            current_exam["src"]["row_end"] = row_i
            cover = current_exam
            current_exam = None
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
                # デバッグ用
                "tag": tag,
                "src": make_src(row_i),
            }

        elif tag == "question" and current_q is not None:
            current_q["question"] = conv_text(params[0])

        # 単独行（大問内）
        elif tag == "sline" and current_q is not None:
            before, after = parse_before_after(row[2])   # C列の (before,after)
            src = make_src(row_i)

            if before != 0:
                current_q["content"].append({
                    "type": "vspace",
                    "value_mm": int(before),
                    "tag": "space_before",
                    "src": src,
                })

            current_q["content"].append({
                "type": "text",
                "value": conv_text(params[0]),
                "tag": "sline",
                "src": src,
            })

            if after != 0:
                current_q["content"].append({
                    "type": "vspace",
                    "value_mm": int(after),
                    "tag": "space_after",
                    "src": src,
                })

        # # 複数行（大問内）
        # elif tag == "b_multiline" and current_q is not None:
        #     wspace = setspace(row[2], "SPACEB_A")
        #     current_multiline = {
        #         "type": "multiline",
        #         "values": [],
        #         "space_before": wspace[0],
        #         "space_after": wspace[1],
        #         "tag": tag,
        #         "src": make_src(row_i),
        #     }

        # elif tag == "text" and current_multiline is not None:
        #     current_multiline["values"].append(conv_text(params[0]))

        # elif tag == "e_multiline" and current_multiline is not None:
        #     current_multiline["src"]["row_end"] = row_i
        #     current_q["content"].append(current_multiline)
        #     current_multiline = None

        # 画像（大問内）
        elif tag == "image" and current_q is not None:
            wimg = parse_with_number(params[0], 0.85)
            img = {
                "type": "image",
                "path": wimg[0],
                "width": wimg[1],
                "tag": tag,
                "src": make_src(row_i),
            }
            current_q["content"].append(img)

        # 複数行（大問内）
        elif tag == "b_multiline" and current_q is not None:
            # ★ b_multiline の (before,after) は B列（row[1]）という仕様
            wspace = setspace(row[1], "SPACEB_A")  # (before, after)
            before = int(wspace[0])
            after  = int(wspace[1])

            # バッファ：space_* は保持しない（確定時にvspaceへ変換）
            current_multiline = {
                "values": [],
                "before": before,
                "after": after,
                "tag": tag,             # "b_multiline"
                "src": make_src(row_i), # {"sheet":..., "row":...}
            }

        elif tag == "text" and current_multiline is not None:
            current_multiline["values"].append(conv_text(params[0]))

        elif tag == "e_multiline" and current_multiline is not None:
            # ブロック範囲を確定
            current_multiline["src"]["row_end"] = row_i
            src_block = current_multiline["src"]

            # values が空ならエラーにする（Excel不整合）
            values = [v for v in current_multiline["values"] if str(v).strip() != ""]
            if not values:
                raise ValueError(f"{src_block['sheet']}!R{src_block['row']}-R{row_i}: multiline が空です")

            # before vspace
            if current_multiline["before"] != 0:
                current_q["content"].append({
                    "type": "vspace",
                    "value_mm": int(current_multiline["before"]),
                    "tag": "space_before",
                    "src": src_block,
                })

            # multiline 本体（space_* を持たせない）
            current_q["content"].append({
                "type": "multiline",
                "values": values,
                "tag": "b_multiline",
                "src": src_block,
            })

            # after vspace
            if current_multiline["after"] != 0:
                current_q["content"].append({
                    "type": "vspace",
                    "value_mm": int(current_multiline["after"]),
                    "tag": "space_after",
                    "src": src_block,
                })

            current_multiline = None


        # 大問コード
        elif tag == "b_code":
            flg = (params[0] == "linenumber") if params else False
            current_code = {
                "type": "code",
                "linenumber": flg,
                "lines": [],
                "tag": tag,
                "src": make_src(row_i),
            }

        elif tag == "code" and current_code is not None:
            current_code["lines"].append(raw_params[0])

        elif tag == "e_code" and current_code is not None and current_q is not None:
            current_code["src"]["row_end"] = row_i
            current_q["content"].append(current_code)
            current_code = None

        # 選択肢ブロック（大問）
        elif tag == "b_select" and current_q is not None:
            style_raw = params[0] if params else "inline"
            sty, sep = _parse_select_style(style_raw)

            select_block = {
                "type": "choices",
                "style": sty,
                "values": [],
                "tag": tag,
                "src": make_src(row_i),
            }
            # 例外のときだけ sep を入れる（=TeX側のデフォルトを活かす）
            if sep is not None:
                select_block["sep"] = sep

            current_select = select_block

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
                "order": order_num,
                "tag": tag,
                "src": make_src(row_i),
            })

        elif tag == "e_select" and current_select is not None:
            # Decide label scheme *before* shuffle
            values = current_select.get("values") or []
            label_scheme = _detect_label_scheme([
                v.get("label") for v in values if isinstance(v, dict)
            ])

            # Shuffle for versions other than A, using Excel G column (order)
            if version != "A":
                order_list = [v.get("order") for v in values if isinstance(v, dict)]
                if any(x is not None for x in order_list):
                    values.sort(key=lambda v: v.get("order", 999) if isinstance(v, dict) else 999)

            # Re-number labels in *display order* for known schemes (A,B,C... / ①②③... / 1,2,3...)
            labels_new = _make_labels(label_scheme, len(values))
            for i, v in enumerate(values):
                if not isinstance(v, dict):
                    continue
                if label_scheme in ("alpha", "circled", "digit"):
                    v["label"] = labels_new[i]
                else:
                    if not v.get("label"):
                        v["label"] = labels_new[i]
                v.pop("order", None)

#            out["content"].append(current_select)
            # b_select / e_select は「大問」側にぶら下げる
            if current_q is None:
                raise ValueError("e_select outside of question")
            current_q["content"].append(current_select)

            current_select = None

        # -------------------------
        # 小問グループ
        # -------------------------
        elif tag == "b_subgroup":
            current_subgroup = {"subquestions": [], "_src": make_src(row_i), "tag": tag}

        elif tag == "b_subquest":
            sub_number += 1

            # v2列（b_question と同じ配置を小問でも使用）
            # C=qid, D=orderB(未使用), E=PB_A_after, F=LS_A_after, G=PB_B_after, H=LS_B_after
            pb_a = row[4] if len(row) >= 5 else None
            ls_a = row[5] if len(row) >= 6 else None
            pb_b = row[6] if len(row) >= 7 else None
            ls_b = row[7] if len(row) >= 8 else None

            # versionに応じて after を選択
            pb_after = _is_one(pb_a) if version == "A" else _is_one(pb_b)
            ls_after = _to_float_or_none(ls_a if version == "A" else ls_b)

            current_sub = {
                "number": str(sub_number),
                "question": None,
                "content": [],
                "_pb_after": pb_after,
                "_ls_after": ls_after,
                "tag": tag,
                "src": make_src(row_i),
            }


        elif tag == "subquest" and current_sub is not None:
            current_sub["question"] = conv_text(params[0])

        # elif tag == "subsline" and current_sub is not None:
        #     wspace = setspace(row[2], "SPACEB_A")
        #     subsline = {
        #         "type": "text",
        #         "values": [conv_text(params[0])],
        #         "space_before": wspace[0],
        #         "space_after": wspace[1],
        #         "tag": tag,
        #         "src": make_src(row_i),
        #     }
        #     current_sub["content"].append(subsline)

        elif tag == "subsline" and current_sub is not None:
            before, after = parse_before_after(row[2])   # subslineの(before,after)列に合わせる
            src = make_src(row_i)

            if before != 0:
                current_sub["content"].append({
                    "type": "vspace",
                    "value_mm": int(before),
                    "tag": "space_before",
                    "src": src,
                })

            current_sub["content"].append({
                "type": "text",
                "value": conv_text(params[0]),
                "tag": "subsline",
                "src": src,
            })

            if after != 0:
                current_sub["content"].append({
                    "type": "vspace",
                    "value_mm": int(after),
                    "tag": "space_after",
                    "src": src,
                })


        # elif tag == "b_submultiline" and current_sub is not None:
        #     wspace = setspace(row[2], "SPACEB_A")
        #     current_multiline = {
        #         "type": "multiline",
        #         "values": [],
        #         "space_before": wspace[0],
        #         "space_after": wspace[1],
        #         "tag": tag,
        #         "src": make_src(row_i),
        #     }


        # elif tag == "subtext" and current_multiline is not None:
        #     current_multiline["values"].append(conv_text(params[0]))

        # elif tag == "e_submultiline" and current_multiline is not None:
        #     current_multiline["src"]["row_end"] = row_i
        #     current_sub["content"].append(current_multiline)
        #     current_multiline = None



        elif tag == "b_submultiline" and current_sub is not None:
            # ★ b_submultiline の (before,after) は B列（row[1]）想定
            wspace = setspace(row[1], "SPACEB_A")  # (before, after)
            before = int(wspace[0])
            after  = int(wspace[1])

            # バッファ：space_* は保持しない（確定時にvspaceへ変換）
            current_submultiline = {
                "values": [],
                "before": before,
                "after": after,
                "tag": "b_submultiline",
                "src": make_src(row_i),   # {"sheet":..., "row":...}
            }

        elif tag == "subtext" and current_submultiline is not None:
            current_submultiline["values"].append(conv_text(params[0]))

        elif tag == "e_submultiline" and current_submultiline is not None:
            # ブロック範囲を確定
            current_submultiline["src"]["row_end"] = row_i
            src_block = current_submultiline["src"]

            # values が空ならエラー（Excel不整合）
            values = [v for v in current_submultiline["values"] if str(v).strip() != ""]
            if not values:
                raise ValueError(f"{src_block['sheet']}!R{src_block['row']}-R{row_i}: submultiline が空です")

            # before vspace
            if current_submultiline["before"] != 0:
                current_sub["content"].append({
                    "type": "vspace",
                    "value_mm": int(current_submultiline["before"]),
                    "tag": "space_before",
                    "src": src_block,
                })

            # multiline 本体（space_* を持たせない）
            current_sub["content"].append({
                "type": "multiline",
                "values": values,
                "tag": "b_submultiline",
                "src": src_block,
            })

            # after vspace
            if current_submultiline["after"] != 0:
                current_sub["content"].append({
                    "type": "vspace",
                    "value_mm": int(current_submultiline["after"]),
                    "tag": "space_after",
                    "src": src_block,
                })

            current_submultiline = None


        elif tag == "subimage" and current_sub is not None:
            wimg = parse_with_number(params[0], 0.85)
            img = {
                "type": "image",
                "path": wimg[0],
                "width": wimg[1],
                "tag": tag,
                "src": make_src(row_i),
            }
            current_sub["content"].append(img)

        # sub code
        elif tag == "b_subcode" and current_sub is not None:
            flg = (params[0] == "linenumber") if params else False
            current_code = {
                "type": "code",
                "linenumber": flg,
                "lines": [],
                "tag": tag,
                "src": make_src(row_i),
            }

        elif tag == "subcode" and current_code is not None:
            current_code["lines"].append(raw_params[0])

        elif tag == "e_subcode" and current_code is not None:
            current_code["src"]["row_end"] = row_i
            current_sub["content"].append(current_code)
            current_code = None

        # subselect
        elif tag == "b_subselect" and current_sub is not None:
            style_raw = params[0] if params else "inline"
            sty, sep = _parse_select_style(style_raw)

            select_block = {
                "type": "choices",
                "style": sty,
                "values": [],
                "tag": tag,
                "src": make_src(row_i),
            }
            if sep is not None:
                select_block["sep"] = sep

            current_select = select_block

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
                "order": order_num,
                "tag": tag,
                "src": make_src(row_i),
            })

        elif tag == "e_subselect" and current_select is not None:
            # Decide label scheme *before* shuffle
            values = current_select.get("values") or []
            label_scheme = _detect_label_scheme([
                v.get("label") for v in values if isinstance(v, dict)
            ])

            # Shuffle for versions other than A, using Excel G column (order)
            if version != "A":
                order_list = [v.get("order") for v in values if isinstance(v, dict)]
                if any(x is not None for x in order_list):
                    values.sort(key=lambda v: v.get("order", 999) if isinstance(v, dict) else 999)

            # Re-number labels in *display order*
            labels_new = _make_labels(label_scheme, len(values))
            for i, v in enumerate(values):
                if not isinstance(v, dict):
                    continue
                if label_scheme in ("alpha", "circled", "digit"):
                    v["label"] = labels_new[i]
                else:
                    if not v.get("label"):
                        v["label"] = labels_new[i]
                v.pop("order", None)

            current_sub["content"].append(current_select)
            current_select = None

        # 小問終了
        elif tag == "e_subquest":
            # 小問を閉じる
            current_sub["src"]["row_end"] = row_i

            pb_after = current_sub.pop("_pb_after", False)
            ls_after = current_sub.pop("_ls_after", None)
            sub_src = current_sub.get("src")

            current_subgroup["subquestions"].append(current_sub)
            current_sub = None

            # v2: 小問 after（subquestions 配列に挿入）
            if ls_after is not None:
                current_subgroup["subquestions"].append({
                    "type": "vspace",
                    "value_mm": int(ls_after) if isinstance(ls_after, int) else int(round(ls_after)),
                    "tag": "LS_sub_after",
                    "src": sub_src if isinstance(sub_src, dict) else {"sheet": sheetname, "row": "?"},
                })
            if pb_after:
                current_subgroup["subquestions"].append({
                    "type": "pagebreak",
                    "tag": "PB_sub_after",
                    "src": sub_src if isinstance(sub_src, dict) else {"sheet": sheetname, "row": "?"},
                })


        # subgroup 終了 → 大問へ合流
        elif tag == "e_subgroup":
            # subgroup範囲を閉じる（JSONに残すわけではないが、内部保持）
            if current_subgroup is not None:
                current_subgroup["_src"]["row_end"] = row_i
            if current_q is not None:
                current_q["subquestions"].extend(current_subgroup["subquestions"])
            current_subgroup = None

        # v2: LINESPACE は subgroup 内のみ → vspace に変換
        elif tag == "LINESPACE":
            v = _to_float_or_none(params[0] if params else None)
            if v is None:
                raise ValueError("LINESPACE の値が空です（v2では数値必須）")
            vspace = {
                "type": "vspace",
                "value_mm": int(v) if isinstance(v, int) else int(round(v)),
                "tag": tag,
                "src": make_src(row_i),
            }
            if current_subgroup is not None:
                current_subgroup["subquestions"].append(vspace)
            else:
                raise ValueError("LINESPACE は subgroup 外では使用できません（v2）")

        # v2: PAGEBREAK 行タグは使わない
        elif tag == "PAGEBREAK":
            raise ValueError("PAGEBREAK 行タグは v2 では使用しません（PB_*_after を使用してください）")

        # 大問終了
        elif tag == "e_question":
            if current_q is not None:
                # ブロック終端
                current_q["src"]["row_end"] = row_i

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

        # after要素の src は b_question 行（q["src"]["row"]）に寄せる
        q_src = q.get("src")

        results.append(q)

        # v2: after 挿入（root配列）
        if ls_after is not None:
            results.append({
                "type": "vspace",
                "value_mm": int(ls_after) if isinstance(ls_after, int) else int(round(ls_after)),
                "tag": "LS_after",
                "src": q_src if isinstance(q_src, dict) else {"sheet": sheetname, "row": "?"},
            })
        if pb_after:
            results.append({
                "type": "pagebreak",
                "tag": "PB_after",
                "src": q_src if isinstance(q_src, dict) else {"sheet": sheetname, "row": "?"},
            })

    return results


def run_json_validator(json_path: Path, strict: bool = False) -> None:
    validator = Path(__file__).resolve().parent / "validate_json.py"
    cmd = [sys.executable, str(validator), str(json_path), "--warn-unknown-keys"]
    if strict:
        cmd.append("--strict")

    r = subprocess.run(cmd)
    if r.returncode != 0:
        # そのまま上流の処理を止める（次の makelatex に進ませない）
        raise RuntimeError(f"JSON validation failed: {json_path}")
 

def apply_version_suffix_to_cover_title(questions: list, version: str) -> None:
    """cover 要素の title に " (A)" / " (B)" などのバージョン表記を付ける。

    - questions 内の最初の type=="cover" のみ対象
    - すでに末尾が "(A)" 等なら二重に付けない
    - title が無い場合は subject から title を補って付ける
    """
    suffix = f"({version})"
    for q in questions:
        if not isinstance(q, dict):
            continue
        if q.get("type") != "cover":
            continue

        base = q.get("title")
        if base is None or not isinstance(base, str) or base.strip() == "":
            # title が空なら subject をベースにする（運用上の保険）
            subj = q.get("subject")
            base = str(subj) if subj is not None else ""

        # すでに付いているなら何もしない
        if isinstance(base, str) and base.rstrip().endswith(f"({version})"):
            q["title"] = base
        else:
            q["title"] = f"{base}{suffix}" if base else suffix.strip()
        break   

_MATH_TOKEN_RE = re.compile(r"\\\(|\\\)|\\\[|\\\]")

def _protect_math_segments_in_str(s: str, mapping: dict, counter: list) -> str:
    """
    文字列中の \(...\) と \[...\] をトークンに置換して退避する。
    normalize_document 等の一律エスケープから数式を守る。
    """
    if not s:
        return s

    out = []
    i = 0
    while i < len(s):
        # start inline \( ... \)
        if s.startswith(r"\(", i):
            j = s.find(r"\)", i + 2)
            if j == -1:
                # 閉じがない → そのまま出す（入力ミス）
                out.append(s[i:])
                break
            seg = s[i:j+2]
            token = f"ZZMATH{counter[0]:06d}ZZ"
            counter[0] += 1
            mapping[token] = seg
            out.append(token)
            i = j + 2
            continue

        # start display \[ ... \]
        if s.startswith(r"\[", i):
            j = s.find(r"\]", i + 2)
            if j == -1:
                out.append(s[i:])
                break
            seg = s[i:j+2]
            token = f"ZZMATH{counter[0]:06d}ZZ"
            counter[0] += 1
            mapping[token] = seg
            out.append(token)
            i = j + 2
            continue

        out.append(s[i])
        i += 1

    return "".join(out)

def _restore_math_segments_in_str(s: str, mapping: dict) -> str:
    if not s or not mapping:
        return s
    # token は英数字のみなので、置換は安全
    for token, seg in mapping.items():
        s = s.replace(token, seg)
    return s

def protect_math_segments(obj, mapping: dict, counter: list):
    """
    outjson 全体を再帰で走査し、文字列中の数式 \(...\), \[...\] を退避する
    """
    if isinstance(obj, dict):
        for k, v in obj.items():
            if isinstance(v, str):
                obj[k] = _protect_math_segments_in_str(v, mapping, counter)
            else:
                protect_math_segments(v, mapping, counter)
    elif isinstance(obj, list):
        for idx, v in enumerate(obj):
            if isinstance(v, str):
                obj[idx] = _protect_math_segments_in_str(v, mapping, counter)
            else:
                protect_math_segments(v, mapping, counter)

def restore_math_segments(obj, mapping: dict):
    if isinstance(obj, dict):
        for k, v in obj.items():
            if isinstance(v, str):
                obj[k] = _restore_math_segments_in_str(v, mapping)
            else:
                restore_math_segments(v, mapping)
    elif isinstance(obj, list):
        for idx, v in enumerate(obj):
            if isinstance(v, str):
                obj[idx] = _restore_math_segments_in_str(v, mapping)
            else:
                restore_math_segments(v, mapping)


if __name__ == "__main__":
    """
    使い方:
      python makedocjsonv2.py <sheetname> [excel_filename]
    例:
      python makedocjsonv2.py 2022001 試験問題.xlsm
    """
    if len(sys.argv) < 2:
        print("Usage: python makedocjsonv2.py <sheetname> [excel_filename]")
        sys.exit(1)

    curdir = Path(__file__).parent.parent
    sheetname = sys.argv[1]
    examdata = "試験問題.xlsm" if len(sys.argv) < 3 else sys.argv[2]
    excel_path = curdir / "input" / examdata

    wb = openpyxl.load_workbook(excel_path)
    ws = wb[sheetname]

    # hash / version 管理（v1と同じ）
    ehash = calc_excel_hash(ws)
    wver = ensure_version_entry(ehash, str(excel_path), sheetname)
    print("wver",wver)

    # qpattern を読む（v1と同じ）
    _ = excel_to_json_v2(ws, version="A")
    qp = qpattern if qpattern else ["A"]

    # single: バージョンが1つだけ（Aのみ／Bのみ等）
    # multi : 複数バージョン（例: A,B）
    versionmode = "single" if (not qp or len(qp) == 1) else "multi"

#    versionmode = "single" if (not qp or qp == ["A"]) else "multi"
    outjson = {"versionmode": versionmode, "versions": []}

    for v in qp:
        questions = excel_to_json_v2(ws, version=v)

        # multi のときだけ、cover.title に " (A)" / " (B)" を付与
        if versionmode == "multi":
            apply_version_suffix_to_cover_title(questions, v)
        dt = datetime.now()
        block = {
            "version": v,
            "questions": questions,
            "metainfo": {
                "type": "metainfo",
                "hash": ehash,
#                "createdatetime": str(wver["createdatetime"]) if isinstance(wver, dict) and "createdatetime" in wver else "",
#                "verno": wver.get("verno") if isinstance(wver, dict) else None,
                "createdatetime": dt.strftime('%Y-%m-%d %H:%M:%S'),
                "verno": wver,
                "inputpath": str(excel_path),
                "sheetname": sheetname,
            },
        }
        outjson["versions"].append(block)

    _math_map = {}
    _counter = [0]
    # 数式を一時退避（\(...\), \[...\]）
    protect_math_segments(outjson, _math_map, _counter)
    # ... outjson を作り終えた直後（dumpの直前）...
    normalize_document(outjson)
    # 数式を復元
    restore_math_segments(outjson, _math_map)

    try:
        validate_document(outjson, strict=True, warn_unknown_keys=True)
    except ContractError as e:
        print(e)          # src付きでエラーを出す
        raise SystemExit(2)

    out = curdir / "work" / f"{sheetname}.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        json.dump(outjson, f, ensure_ascii=False, indent=2)

    print(f"✅ jsonファイルを作成しました: {out}")

    run_json_validator(out, strict=False)  # まずは strict=False 推奨
    print("✅ JSON contract validation OK")