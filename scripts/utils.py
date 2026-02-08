# examtools/utils.py
from pathlib import Path
import hashlib
import re
import json

def calc_excel_hash(sheet):
    """シート内容からハッシュ値を計算"""
    content = []
    for row in sheet.iter_rows(values_only=True):
        content.append(",".join("" if v is None else str(v) for v in row))
    return hashlib.md5("\n".join(content).encode("utf-8")).hexdigest()

def jsonmetainfo(workdir, jsonfile):
    path = Path(workdir) / jsonfile
    if not path.exists():
        return (None, None, None)
    with open(path, "r", encoding="utf-8") as f:
        jdata = json.load(f)

    # ------------------------------
    # 1. dict 形式（versionmode  singleでもmultiでも１つ目の配列要素をチェック）
    # ------------------------------
    if isinstance(jdata, dict):

        versions = jdata.get("versions", [])
        if not versions:
            return (None, None, None)

        v0 = versions[0]['questions'][0]

        # questions.json 形式:  entry["metainfo"] に直接入っている
        if "metainfo" in v0:
            meta = v0["metainfo"]
        else:
            # answers.json 形式:  entry["header"]["metainfo"] に入っている
            header = v0.get("header", {})
            meta = header.get("metainfo", {})

        return (
            meta.get("hash"),
            meta.get("inputpath"),
            meta.get("sheetname")
        )

    # ------------------------------
    # 2. list 形式（旧仕様）
    # ------------------------------
    if isinstance(jdata, list) and jdata:
        first = jdata[0]

        # 2-1. ヘッダに metainfo があるパターン
        if isinstance(first, dict) and "metainfo" in first:
            meta = first["metainfo"]
            return (
                meta.get("hash"),
                meta.get("inputpath"),
                meta.get("sheetname")
            )

        # 2-2. 配列内に type == "metainfo" の要素があるパターン
        item = next(
            (d for d in jdata if isinstance(d, dict) and d.get("type") == "metainfo"),
            None
        )
        if item:
            return (
                item.get("hash"),
                item.get("inputpath"),
                item.get("sheetname")
            )

    # どのパターンにも当てはまらない場合
    return (None, None, None)

def setspace(text, p1):
    defaults = {"ANSSIZE": (50.0, 60.0), "SPACEB_A": (0.0, 0.0), "ANSWH": (1.0, 1.0)}
    DEFAULTb, DEFAULTa = defaults.get(p1, (0.0, 0.0))

    def parse_num(val, default):
        try:
            return float(val) if val.strip() else default
        except ValueError:
            return default

    # 入力が文字列でなければデフォルト
    if not isinstance(text, str):
        return DEFAULTb, DEFAULTa

    # ( ) で囲まれていなければデフォルト
    text = text.strip()
    if not (text.startswith("(") and text.endswith(")")):
        return DEFAULTb, DEFAULTa

    # () 中身を分割
    parts = text[1:-1].split(",")

    # --- ANSWH専用ルール ---
    if p1 == "ANSWH":
        if len(parts) == 1:   # 例: (3)
            return parse_num(parts[0], 1.0), 1.0
        elif len(parts) == 2: # 例: (,3)
            return parse_num(parts[0], 1.0), parse_num(parts[1], 1.0)
        return 1.0, 1.0

    # --- 共通ルール ---
    if len(parts) != 2:
        return DEFAULTb, DEFAULTa
    return parse_num(parts[0], DEFAULTb), parse_num(parts[1], DEFAULTa)

def parse_with_number(s: str, default: float = 0.5) -> tuple[str, float]:
    """
    文字列から abc[数値] の形式を解析する。
    数値が不正または存在しない場合はデフォルト値を返す。
    """
    # 正規表現で「名前」「[中身]」を分解
    match = re.match(r"^([^\[]+)(?:\[(.*)\])?$", s.strip())
    if not match:
        return s, default

    name, num_part = match.groups()

    if not num_part:  # [] がない、または空
        return name, default

    try:
        value = float(num_part)
    except ValueError:
        value = default

    return name, value

if __name__ == "__main__":
    from pathlib import Path

    # 現在の場所
    curdir = Path(__file__).parent.parent
    # --- JSON書き出し ---
    sheetname="2021901"
    
    print(jsonmetainfo(curdir / "work",f"answers_{sheetname}.json"))

