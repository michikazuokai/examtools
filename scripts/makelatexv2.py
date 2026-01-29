import json
from pathlib import Path
import argparse

# =========================
# Helpers
# =========================
def latex_escape(s: str) -> str:
    # Escape LaTeX-special characters for plain text.
    # NOTE:
    #  - Backslash is kept as-is (makedocjsonv2 already converts \0 -> \textbackslash 0)
    #  - $ is kept as-is so you can embed math if needed
    if s is None:
        return ""
    rep = {
        "&": r"\&",
        "%": r"\%",
        "#": r"\#",
        "_": r"\_",
        "{": r"\{",
        "}": r"\}",
        "~": r"\textasciitilde{}",
        "^": r"\textasciicircum{}",
    }
    return "".join(rep.get(ch, ch) for ch in str(s))


def _opt_kv(before=None, after=None) -> str:
    kv = []
    if isinstance(before, (int, float)) and abs(before) > 1e-12:
        kv.append(f"before={before}")
    if isinstance(after, (int, float)) and abs(after) > 1e-12:
        kv.append(f"after={after}")
    return "[" + ",".join(kv) + "]" if kv else ""


def render_sline(item: dict) -> str:
    vals = item.get("values", []) or []
    sb = item.get("space_before", 0.0)
    sa = item.get("space_after", 0.0)
    opt = _opt_kv(sb, sa)
    return "\n".join(rf"\sline{opt}{{{latex_escape(v)}}}" for v in vals)


def render_multiline(item: dict) -> str:
    vals = item.get("values", []) or []
    sb = item.get("space_before", 0.0)
    sa = item.get("space_after", 0.0)
    opt = _opt_kv(sb, sa)

    out = [rf"\begin{{multiline}}{opt}"]
    out += [rf"\mline{{{latex_escape(v)}}}" for v in vals]
    out.append(r"\end{multiline}")
    return "\n".join(out)


def render_choices(item: dict) -> str:
    style = item.get("style", "normal")
    values = item.get("values", []) or []

    # v2: legacy "space" (from inline(8)) is treated as sep
    sep = item.get("sep", None)
    if sep is None:
        sep = item.get("space", None)

    opts = []
    if style == "inline":
        opts.append("type=inline")
        if sep is not None:
            opts.append(f"sep={sep}")
    else:
        opts.append("type=normal")

    opt = "[" + ",".join(opts) + "]" if opts else ""
    out = [rf"\begin{{choices}}{opt}"]
    for v in values:
        label = v.get("label", "")
        text = latex_escape(v.get("text", ""))
        if label:
            out.append(rf"  \item[\ChoiceLabel{{{latex_escape(label)}}}] {text}")
        else:
            out.append(rf"  \item {text}")
    out.append(r"\end{choices}")
    return "\n".join(out)


def render_image(item: dict) -> str:
    path = item.get("path", "")
    width_val = item.get("width", 0.85)
    try:
        width_val = float(width_val)
    except Exception:
        width_val = 0.85
    return rf"\image{{{latex_escape(path)}}}{{{width_val}}}"


def render_code(item: dict) -> str:
    lines = item.get("lines", []) or []
    opt = "[linenumber=1]" if item.get("linenumber") else ""
    out = [rf"\begin{{code}}{opt}"]
    out.extend(lines)  # verbatim
    out.append(r"\end{code}")
    return "\n".join(out)


def render_vspace(item: dict) -> str:
    v = item.get("value", 0.0)
    try:
        v = float(v)
    except Exception:
        v = 0.0
    return rf"\vspace*{{{v}mm}}"


def render_pagebreak(_: dict) -> str:
    return r"\newpage"


def render_content_list(content: list) -> str:
    parts = []
    for item in content or []:
        t = item.get("type")
        if t == "choices":
            parts.append(render_choices(item))
        elif t == "text":
            parts.append(render_sline(item))
        elif t == "multiline":
            parts.append(render_multiline(item))
        elif t == "image":
            parts.append(render_image(item))
        elif t == "code":
            parts.append(render_code(item))
        elif t == "vspace":
            parts.append(render_vspace(item))
        elif t == "pagebreak":
            parts.append(render_pagebreak(item))
        else:
            parts.append(rf"% [WARN] unknown content type: {t}")
    return "\n\n".join(parts)


def generate_version_tex(data: dict, version: str = "A", include_cover: bool = False) -> str:
    versions = data.get("versions", []) or []
    v = next((x for x in versions if x.get("version") == version), None)
    if not v:
        raise ValueError(f"version {version} not found")

    out = []
    for q in v.get("questions", []) or []:
        qtype = q.get("type")

        # cover
        if qtype == "cover":
            if include_cover:
                title = latex_escape(q.get("title", ""))
                out.append(rf"% --- COVER: {title} ---")
                for n in (q.get("notes", []) or []):
                    out.append(rf"% {latex_escape(n)}")
                out.append("")
            continue

        # root controls
        if qtype == "pagebreak":
            out.append(r"\newpage")
            continue
        if qtype == "vspace":
            out.append(render_vspace(q))
            continue

        question = q.get("question", "")
        if not question:
            continue

        out.append(rf"\begin{{question}}{{{latex_escape(question)}}}")
        out.append(render_content_list(q.get("content", [])))

        # subquestions container (subquestion only switches mode + indent)
        subs = q.get("subquestions", []) or []
        if subs:
            out.append(r"\begin{subquestion}")
            for sq in subs:
                if isinstance(sq, dict) and sq.get("type") in ("vspace", "pagebreak"):
                    out.append(render_vspace(sq) if sq.get("type") == "vspace" else render_pagebreak(sq))
                    continue

                sq_q = sq.get("question", "") if isinstance(sq, dict) else ""
                if not sq_q:
                    continue
                out.append(rf"\begin{{question}}{{{latex_escape(sq_q)}}}")
                out.append(render_content_list(sq.get("content", [])))
                out.append(r"\end{question}")
            out.append(r"\end{subquestion}")

        out.append(r"\end{question}")
        out.append("")

    return "\n\n".join(out).strip() + "\n"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("sheetname", help="Excel sheet name (also json name, e.g. 1020201)")
    ap.add_argument("--version", default="A", help="A/B/... (default A)")
    ap.add_argument("--in", dest="inpath", default=None, help="input json path (default: ../work/<sheet>.json)")
    ap.add_argument("--out", dest="outpath", default=None, help="output tex path (default: ../sandbox/<sheet>.tex)")
    ap.add_argument("--cover", action="store_true", help="include cover comments")
    args = ap.parse_args()

    rootdir = Path(__file__).parent.parent
    inpath = Path(args.inpath) if args.inpath else (rootdir / "work" / f"{args.sheetname}.json")
    outpath = Path(args.outpath) if args.outpath else (rootdir / "sandbox" / f"{args.sheetname}.tex")

    data = json.loads(inpath.read_text(encoding="utf-8"))
    tex = generate_version_tex(data, version=args.version, include_cover=args.cover)

    outpath.parent.mkdir(parents=True, exist_ok=True)
    outpath.write_text(tex, encoding="utf-8")
    print(f"✅ wrote: {outpath}")


if __name__ == "__main__":
    main()
