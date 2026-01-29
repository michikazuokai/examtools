#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""makelatexv2.py (Step3+Step4)

JSON(v2, canonical) -> LaTeX.

Changes in this fixed version:
- Aligns with Step2 "single environments":
  * choices only: \begin{choices}[type=normal|inline,sep=int] ... \citem{A}{text}
  * code only:    \begin{code}[linenumber=true|false] ... \end{code}
- Emits TeX trace comments to map output back to Excel rows:
  * %% QBEGIN / %% QEND
  * %% type=... tag=... src=Sheet!R..-R..
- Supports canonical JSON keys:
  * text: {type:"text", value:"..."}
  * vspace: {type:"vspace", value_mm:int}
  * multiline: {type:"multiline", values:[...]}
  * choices: {type:"choices", style:"normal|inline", sep?:int, values:[{label,text,src,tag},...]}

CLI (kept compatible):
  python makelatex.py <sheetname> --version A

"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Dict, List, Optional


import re
from datetime import datetime

# =========================
# Helpers
# =========================

def latex_escape(s: Any) -> str:
    """Escape LaTeX-special characters for plain text,
    BUT do NOT escape inside inline/display math: \( ... \), \[ ... \].

    Also avoid double-escaping already-escaped sequences like \{, \}, \&, \_ ...
    """
    if s is None:
        return ""

    text = str(s)

    # outside-math replacements
    rep_basic = {
        "&": r"\&",
        "%": r"\%",
        "#": r"\#",
        "_": r"\_",
        "{": r"\{",
        "}": r"\}",
    }

    out: list[str] = []
    i = 0
    in_math = False
    end_token: str | None = None

    while i < len(text):
        if not in_math:
            # enter math
            if text.startswith(r"\(", i):
                in_math = True
                end_token = r"\)"
                out.append(r"\(")
                i += 2
                continue
            if text.startswith(r"\[", i):
                in_math = True
                end_token = r"\]"
                out.append(r"\[")
                i += 2
                continue

            ch = text[i]
            prev = text[i - 1] if i > 0 else ""

            # already escaped like \{ \} \& \_ ... -> keep as-is
            if ch in rep_basic:
                if prev == "\\":
                    out.append(ch)
                else:
                    out.append(rep_basic[ch])
                i += 1
                continue

            if ch == "~":
                if prev == "\\":
                    out.append(ch)
                else:
                    out.append(r"\textasciitilde{}")
                i += 1
                continue

            if ch == "^":
                if prev == "\\":
                    out.append(ch)
                else:
                    out.append(r"\textasciicircum{}")
                i += 1
                continue

            out.append(ch)
            i += 1

        else:
            # exit math
            if end_token and text.startswith(end_token, i):
                in_math = False
                out.append(end_token)
                i += 2
                end_token = None
                continue

            # inside math: do nothing (preserve \frac{...}{...}, _, ^, etc.)
            out.append(text[i])
            i += 1

    return "".join(out)


def latex_escape_multiline(s: Any) -> str:
    """
    LaTeX の引数に入れる複数行テキスト用。
    - 各行を latex_escape して
    - 改行は LaTeX の \\ に変換
    """
    if s is None:
        return ""
    text = str(s)
    lines = text.splitlines()
    if not lines:
        return ""
    return r"\\ ".join(latex_escape(line) for line in lines)

def _src_loc(src: Any) -> str:
    if not isinstance(src, dict):
        return "?"
    sheet = src.get("sheet", "?")
    row = src.get("row", "?")
    row_end = src.get("row_end")
    if row_end is not None:
        return f"{sheet}!R{row}-R{row_end}"
    return f"{sheet}!R{row}"


def _trace_line(obj: Dict[str, Any], prefix: str = "") -> str:
    """Return a TeX comment line mapping an element back to Excel."""
    t = obj.get("type", "?")
    tag = obj.get("tag", "?")
    loc = _src_loc(obj.get("src"))
    extra = (prefix + " ") if prefix else ""
    return f"%% {extra}type={t} tag={tag} src={loc}"


def _get_mm(item: Dict[str, Any], key_primary: str, key_fallback: Optional[str] = None, default: int = 0) -> int:
    """Read mm value as int. Accepts legacy float/string but coerces to int."""
    v = item.get(key_primary)
    if v is None and key_fallback:
        v = item.get(key_fallback)
    if v is None:
        return default
    try:
        # allow "8", 8.0 etc.
        return int(round(float(v)))
    except Exception:
        return default


def _boolish(x: Any) -> bool:
    if isinstance(x, bool):
        return x
    if isinstance(x, (int, float)):
        return x != 0
    if isinstance(x, str):
        return x.strip().lower() in ("1", "true", "yes", "y", "on")
    return False


# =========================
# Renderers (elements)
# =========================

def render_text(item: Dict[str, Any], with_trace: bool = True) -> str:
    # canonical: value
    # legacy: values[list[str]]
    parts: List[str] = []
    if with_trace:
        parts.append(_trace_line(item))

    if "value" in item and isinstance(item.get("value"), str):
        parts.append(rf"\sline{{{latex_escape(item.get('value'))}}}")
        return "\n".join(parts)

    vals = item.get("values") or []
    if isinstance(vals, list):
        for v in vals:
            parts.append(rf"\sline{{{latex_escape(v)}}}")
    return "\n".join(parts)


def render_vspace(item: Dict[str, Any], with_trace: bool = True) -> str:
    mm = _get_mm(item, "value_mm", key_fallback="value", default=0)
    lines: List[str] = []
    if with_trace:
        lines.append(_trace_line(item))
    lines.append(rf"\vspace*{{{mm}mm}}")
    return "\n".join(lines)


def render_pagebreak(item: Dict[str, Any], with_trace: bool = True) -> str:
    lines: List[str] = []
    if with_trace:
        lines.append(_trace_line(item))
    lines.append(r"\newpage")
    return "\n".join(lines)


def render_image(item: Dict[str, Any], with_trace: bool = True) -> str:
    # canonical: path, width (ratio)
    path = item.get("path", "")
    width_val = item.get("width", 0.85)
    try:
        width_val = float(width_val)
    except Exception:
        width_val = 0.85

    lines: List[str] = []
    if with_trace:
        lines.append(_trace_line(item))
    lines.append(rf"\image{{{latex_escape(path)}}}{{{width_val}}}")
    return "\n".join(lines)


def render_code(item: Dict[str, Any], with_trace: bool = True) -> str:
    lines = item.get("lines") or []
    if not isinstance(lines, list):
        lines = []

    linenumber = _boolish(item.get("linenumber"))
    opt = "[linenumber=true]" if linenumber else "[linenumber=false]"

    out: List[str] = []
    if with_trace:
        out.append(_trace_line(item))
    out.append(rf"\begin{{code}}{opt}")
    # verbatim-like (listings). Do NOT escape.
    out.extend([str(x) for x in lines])
    out.append(r"\end{code}")
    return "\n".join(out)


def render_multiline(item: Dict[str, Any], with_trace: bool = True) -> str:
    vals = item.get("values") or []
    if not isinstance(vals, list):
        vals = []

    out: List[str] = []
    if with_trace:
        out.append(_trace_line(item))
    out.append(r"\begin{multiline}")
    out += [rf"\mline{{{latex_escape(v)}}}" for v in vals]
    out.append(r"\end{multiline}")
    return "\n".join(out)


def render_choices(item: Dict[str, Any], with_trace: bool = True) -> str:
    style = item.get("style", "normal")
    values = item.get("values") or []
    if not isinstance(values, list):
        values = []

    # legacy alias: space -> sep
    sep = item.get("sep", None)
    if sep is None:
        sep = item.get("space", None)

    opts: List[str] = []
    if style == "inline":
        opts.append("type=inline")
        if sep is not None:
            try:
                opts.append(f"sep={int(round(float(sep)))}")
            except Exception:
                # leave it off; validator should have caught it
                pass
    else:
        opts.append("type=normal")

    opt = "[" + ",".join(opts) + "]" if opts else ""

    out: List[str] = []
    if with_trace:
        out.append(_trace_line(item))

    out.append(rf"\begin{{choices}}{opt}")

    # Each option: emit trace + \citem{label}{text}
    for v in values:
        if isinstance(v, str):
            # simplest representation
            if with_trace:
                out.append("%% type=choiceitem tag=? src=?")
            out.append(rf"\citem{{}}{{{latex_escape(v)}}}")
            continue

        if not isinstance(v, dict):
            if with_trace:
                out.append("%% type=choiceitem tag=? src=?")
            out.append(rf"\citem{{}}{{{latex_escape(str(v))}}}")
            continue

        if with_trace:
            out.append(_trace_line({"type": "choiceitem", "tag": v.get("tag", "select"), "src": v.get("src", {})}))

        label = v.get("label", "")
        text = v.get("text", "")
        out.append(rf"\citem{{{latex_escape(label)}}}{{{latex_escape(text)}}}")

    out.append(r"\end{choices}")
    return "\n".join(out)


def render_content_list(content: Any, with_trace: bool = True) -> str:
    parts: List[str] = []
    if not isinstance(content, list):
        content = []

    for item in content:
        if not isinstance(item, dict):
            parts.append(rf"% [WARN] non-dict content item: {latex_escape(item)}")
            continue

        t = item.get("type")
        if t == "choices":
            parts.append(render_choices(item, with_trace=with_trace))
        elif t == "text":
            parts.append(render_text(item, with_trace=with_trace))
        elif t == "multiline":
            parts.append(render_multiline(item, with_trace=with_trace))
        elif t in ("image", "subimage"):
            parts.append(render_image(item, with_trace=with_trace))
        elif t == "code":
            parts.append(render_code(item, with_trace=with_trace))
        elif t == "vspace":
            parts.append(render_vspace(item, with_trace=with_trace))
        elif t == "pagebreak":
            parts.append(render_pagebreak(item, with_trace=with_trace))
        else:
            # Keep a trace even for unknown types.
            if with_trace:
                parts.append(_trace_line(item))
            parts.append(rf"% [WARN] unknown content type: {latex_escape(t)}")

    return "\n\n".join(parts)


def build_cover_footer(metainfo: dict) -> str:
    # hash 先頭7
    h = (metainfo or {}).get("hash", "")
    h7 = str(h)[:7] if h else "0000000"

    # 処理年（createdatetime優先、無ければ今年）
    cd = (metainfo or {}).get("createdatetime", "")
    yy = None
    if isinstance(cd, str) and cd:
        m = re.match(r"^\s*(\d{4})", cd)
        if m:
            yy = int(m.group(1)) % 100
    if yy is None:
        yy = datetime.now().year % 100

    # version（2桁）
    ver = metainfo.get("verno")
    v = ver if isinstance(ver, int) else 0

    return f"{h7}{yy:02d}-{v:02d}"


# =========================
# Document generation
# =========================

def _q_loc(q: Dict[str, Any]) -> str:
    return _src_loc(q.get("src"))


def generate_version_tex(data: Dict[str, Any], version: str = "A", include_cover: bool = False, with_trace: bool = True) -> str:
    versions = data.get("versions", []) or []
    v = next((x for x in versions if x.get("version") == version), None)
    if not v:
        raise ValueError(f"version {version} not found")

    out: List[str] = []

    for q in v.get("questions", []) or []:
        if not isinstance(q, dict):
            continue

        qtype = q.get("type")

        # cover
        if qtype == "cover":
            if include_cover:
                # JSON cover -> LaTeX title page
                title = q.get("title") or q.get("subject") or ""

                notes = q.get("notes")
                if notes is None:
                    notes = q.get("instruction")  # legacy alias
                if notes is None:
                    notes = ""

                items = []
                if isinstance(notes, list):
                    items = [str(x).strip() for x in notes if str(x).strip()]
                else:
                    # 改行で分割して \item 化
                    items = [line.strip() for line in str(notes).splitlines() if line.strip()]

                notes_text = "\n".join([r"\item " + latex_escape(x) for x in items])

                # if isinstance(notes, list):
                #     notes_text = "\n".join(str(x) for x in notes)
                # else:
                #     notes_text = str(notes)

                # ★ここ：footer_text を自動生成（JSON側に無ければ）
                metainfo = v.get("metainfo") or {}
                # version_no は JSON上のどこにあるかで選ぶ：
                # 例：metainfo.verno を使う / もしくは versions の version を使う
                footer_text = build_cover_footer(metainfo)

                out.append(r"% --- COVER PAGE ---")
                out.append(
                    rf"\ExamCover{{{latex_escape(title)}}}{{{latex_escape_multiline(notes_text)}}}[{latex_escape(str(footer_text))}]"
                )
                out.append("")
            continue


        # root controls (top-level)
        if qtype == "pagebreak":
            out.append(render_pagebreak(q, with_trace=with_trace))
            out.append("")
            continue
        if qtype == "vspace":
            out.append(render_vspace(q, with_trace=with_trace))
            out.append("")
            continue

        question = q.get("question", "")
        if not question:
            continue

        # QBEGIN trace
        if with_trace:
            qid = q.get("qid", "")
            no = q.get("number", "")
            tag = q.get("tag", "?")
            loc = _q_loc(q)
            out.append(f"%% QBEGIN qid={qid} no={no} tag={tag} src={loc}")

        out.append(rf"\begin{{question}}{{{latex_escape(question)}}}")
        out.append(render_content_list(q.get("content", []), with_trace=with_trace))

        # subquestions container (subquestion only switches mode + indent)
        subs = q.get("subquestions", []) or []
        if isinstance(subs, list) and subs:
            if with_trace:
                out.append(f"%% SUBBEGIN parent_qid={q.get('qid','')} src={_q_loc(q)}")
            out.append(r"\begin{subquestion}")
            for sq in subs:
                if not isinstance(sq, dict):
                    continue

                # allow controls in subquestions list
                if sq.get("type") == "vspace":
                    out.append(render_vspace(sq, with_trace=with_trace))
                    continue
                if sq.get("type") == "pagebreak":
                    out.append(render_pagebreak(sq, with_trace=with_trace))
                    continue

                sq_q = sq.get("question", "")
                if not sq_q:
                    continue

                if with_trace:
                    out.append(f"%% SQBEGIN tag={sq.get('tag','?')} src={_src_loc(sq.get('src'))}")

                out.append(rf"\begin{{question}}{{{latex_escape(sq_q)}}}")
                out.append(render_content_list(sq.get("content", []), with_trace=with_trace))
                out.append(r"\end{question}")

                if with_trace:
                    out.append("%% SQEND")

            out.append(r"\end{subquestion}")
            if with_trace:
                out.append("%% SUBEND")

        out.append(r"\end{question}")

        # QEND trace
        if with_trace:
            out.append(f"%% QEND qid={q.get('qid','')} src={_q_loc(q)}")

        out.append("")

    return "\n\n".join(out).strip() + "\n"

def project_root() -> Path:
    # scripts/ の1つ上を root とみなす
    return Path(__file__).resolve().parent.parent

def load_versions_from_json(sheet: str) -> List[str]:
    """
    work/<sheet>.json を読み、versions[].version から A/B を取得する。
    ない場合は ["A"] とみなす。
    """
    root = project_root()
    json_path = root / "work" / f"{sheet}.json"
    if not json_path.exists():
        # ③だけ担当なので、jsonが無ければ版が分からない
        raise FileNotFoundError(f"work json not found: {json_path}")

    data = json.loads(json_path.read_text(encoding="utf-8"))
    vers = []
    for v in (data.get("versions") or []):
        vv = v.get("version")
        if vv:
            vers.append(str(vv))
    return vers if vers else ["A"]


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("sheetname", help="Excel sheet name (also json name, e.g. 1020201)")
#    ap.add_argument("--version", default="A", help="A/B/... (default A)")
    ap.add_argument("--version", help="A/B/... (default A)")
    ap.add_argument("--in", dest="inpath", default=None, help="input json path (default: ../work/<sheet>.json)")
    ap.add_argument("--out", dest="outpath", default=None, help="output tex path (default: ../sandbox/<sheet>.tex)")
    ap.add_argument("--cover", action="store_true", help="include cover comments")
    ap.add_argument("--no-trace", action="store_true", help="disable trace comments")
    args = ap.parse_args()

#    rootdir = Path(__file__).parent.parent
    rootdir = project_root()
    inpath = Path(args.inpath) if args.inpath else (rootdir / "work" / f"{args.sheetname}.json")
    #outpath = Path(args.outpath) if args.outpath else (rootdir / "sandbox" / f"{args.sheetname}.tex")

    data = json.loads(inpath.read_text(encoding="utf-8"))

    # バージョンの取得
    if args.version is None:
        vers =  load_versions_from_json(args.sheetname)
    else:
        vers =  [args.version]
    print(vers)
    for ver in vers:
        outpath = rootdir / "output" /   args.sheetname / ver / f"{args.sheetname}_{ver}_body.tex"

#        tex = generate_version_tex(data, version=args.version, include_cover=args.cover, with_trace=(not args.no_trace))
        tex = generate_version_tex(data, ver, include_cover=args.cover, with_trace=(not args.no_trace))

        outpath.parent.mkdir(parents=True, exist_ok=True)
        outpath.write_text(tex, encoding="utf-8")
        print(f"✅ wrote: {outpath}")


if __name__ == "__main__":
    main()
