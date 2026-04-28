#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
texmerge_compile.py  (V2)
  body.tex + templates -> full .tex -> lualatex compile -> pdf

This script does NOT:
  - generate json
  - generate body.tex from json

It only:
  - finds body.tex for each version (A/B...)
  - copies templates into output/<sheet>/<ver>/
  - injects graphicspath for images/<sheet>/
  - creates full tex (compile-ready)
  - runs lualatex

Usage:
  python scripts/texmerge_compile.py 1020201
  python scripts/texmerge_compile.py 1020201 --version A
  python scripts/texmerge_compile.py 1020201 --runs 2
"""

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path
from typing import List, Optional
from utils import get_exam_path

def project_root() -> Path:
    # scripts/ の1つ上を root とみなす
    return Path(__file__).resolve().parent.parent


def read_text(p: Path) -> str:
    return p.read_text(encoding="utf-8")


def write_text(p: Path, s: str) -> None:
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(s, encoding="utf-8")


def run_cmd(cmd: List[str], cwd: Optional[Path] = None) -> None:
    print("▶", " ".join(cmd))
    r = subprocess.run(cmd, cwd=str(cwd) if cwd else None)
    if r.returncode != 0:
        raise RuntimeError(f"Command failed (code={r.returncode}): {' '.join(cmd)}")

def load_versions_from_json_path(json_path: Path) -> List[str]:
    if not json_path.exists():
        raise FileNotFoundError(f"work json not found: {json_path}")

    data = json.loads(json_path.read_text(encoding="utf-8"))
    vers = []
    for v in (data.get("versions") or []):
        vv = v.get("version")
        if vv:
            vers.append(str(vv))
    return vers if vers else ["A"]

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

def find_body_tex(work_dir: Path, sheet: str, version: str) -> Path:
    body_path = work_dir / "latex" / version / f"{sheet}_{version}_body.tex"

    if body_path.exists():
        return body_path

    raise FileNotFoundError(f"body.tex not found: {body_path}")

def _find_body_tex(sheet: str, version: str) -> Path:
    """
    body.tex の探索ルール（②側の出力場所が揺れても拾えるようにする）
    """
    root = project_root()
    candidates = [
        root / "output" / sheet / version / f"{sheet}_{version}_body.tex",
        root / "work" / f"{sheet}_{version}_body.tex",
        root / "output" / sheet / version / "body.tex",
    ]
    for p in candidates:
        if p.exists():
            return p
    raise FileNotFoundError(
        "body.tex not found. Tried:\n" + "\n".join(str(c) for c in candidates)
    )


def copy_templates_to(build_dir: Path) -> None:
    """
    templates/latex を build_dir にコピー。
    \input{preamble.tex} 等が同一ディレクトリを参照できるようにするため。
    """
    root = project_root()
    tdir = root / "templates" / "latex"
    if not tdir.exists():
        raise FileNotFoundError(f"templates/latex not found: {tdir}")

    for name in ["main.tpl.tex", "preamble.tex", "styles.tex", "macros.tex"]:
        src = tdir / name
        if not src.exists():
            raise FileNotFoundError(f"template missing: {src}")
        shutil.copy2(src, build_dir / name)


def inject_graphicspath(build_dir: Path, sheet: str) -> None:
    """
    preamble.tex に \graphicspath を注入する。
    - preamble.tex に @@GRAPHICSPATH@@ があれば置換
    - 無ければ末尾に追記
    画像は images/<sheet>/ を見る。
    """
    root = project_root()
    preamble_path = build_dir / "preamble.tex"
    if not preamble_path.exists():
        raise FileNotFoundError(f"preamble.tex not found in build_dir: {preamble_path}")

    img_dir = (root / "images" / sheet).resolve()
    img_dir_posix = img_dir.as_posix()

    # 画像参照はファイル名だけでいけるようにする（\includegraphics{5.png} 等）
    gsp = (
        r"\newcommand{\assetpath}{" + img_dir_posix + r"}" + "\n"
        r"\graphicspath{{\assetpath/}}" + "\n"
    )

    preamble = read_text(preamble_path)
    if "@@GRAPHICSPATH@@" in preamble:
        preamble = preamble.replace("@@GRAPHICSPATH@@", gsp)
    else:
        preamble += "\n% --- auto inserted graphicspath ---\n" + gsp

    write_text(preamble_path, preamble)


def build_full_tex(build_dir: Path, body_tex_filename: str, out_tex_filename: str) -> Path:
    """
    main.tpl.tex の @@BODY@@ を \input{body_tex_filename} に差し替えて
    compile-ready の .tex を作る。
    """
    main_tpl_path = build_dir / "main.tpl.tex"
    if not main_tpl_path.exists():
        raise FileNotFoundError(f"main.tpl.tex not found in build_dir: {main_tpl_path}")

    main_tpl = read_text(main_tpl_path)
    insert = rf"\input{{{body_tex_filename}}}"

    if "@@BODY@@" in main_tpl:
        full = main_tpl.replace("@@BODY@@", insert)
    else:
        # 置換口が無い場合は事故回避で document 内先頭に挿入
        full = main_tpl.replace(r"\begin{document}", r"\begin{document}" + "\n" + insert + "\n")

    out_tex = build_dir / out_tex_filename
    write_text(out_tex, full)
    return out_tex


def compile_lualatex(tex_path: Path, runs: int = 1) -> Path:
    """
    build_dir で lualatex を回す。
    """
    build_dir = tex_path.parent
    for _ in range(max(1, runs)):
        run_cmd(
#            ["lualatex", "-interaction=nonstopmode", "-halt-on-error", "-file-line-error", tex_path.name],
            ["lualatex", "-file-line-error", "-interaction=nonstopmode", "-halt-on-error", tex_path.name],
            cwd=build_dir
        )

    pdf_path = tex_path.with_suffix(".pdf")
    if not pdf_path.exists():
        raise RuntimeError(f"PDF not generated: {pdf_path}")
    return pdf_path


def _main():
    ap = argparse.ArgumentParser()
    ap.add_argument("sheet", help="sheet name (subject id), e.g. 1020201")
    ap.add_argument("--version", default=None, help="compile only one version, e.g. A or B")
    ap.add_argument("--runs", type=int, default=2, help="lualatex runs (default=1)")
    args = ap.parse_args()

    root = project_root()
    sheet = args.sheet

    versions = load_versions_from_json(sheet)
    if args.version:
        versions = [args.version]

    for ver in versions:
        # ②の出力(body)を探す
        body_path = find_body_tex(sheet, ver)


        # 出力先（版ごと）
        out_dir = root / "output" / sheet / ver
        out_dir.mkdir(parents=True, exist_ok=True)

        # テンプレコピー
        copy_templates_to(out_dir)

        # graphicspath 注入
        inject_graphicspath(out_dir, sheet)

        # body を out_dir にコピー（\input 参照を安定させる）
        body_name = f"{sheet}_{ver}_body.tex"
        dst_body = out_dir / body_name

        # すでに同じ場所ならコピー不要（SameFileError回避）
        try:
            if body_path.resolve() != dst_body.resolve():
                shutil.copy2(body_path, dst_body)
        except FileNotFoundError:
            # resolve() が失敗するケース（ネットワーク/特殊パス等）対策
            if str(body_path) != str(dst_body):
                shutil.copy2(body_path, dst_body)

        # compile-ready tex を生成
        full_name = f"{sheet}_{ver}.tex"
        full_tex_path = build_full_tex(out_dir, body_name, full_name)
        print(f"✅ TeX merged: {full_tex_path}")

        # compile
        pdf_path = compile_lualatex(full_tex_path, runs=args.runs)
        print(f"✅ PDF compiled: {pdf_path}")

    print("🎯 Done.")

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("subject_no", help="科目番号。通常はシート名にも使う")
    ap.add_argument("year", nargs="?", default="2026", help="年度。省略時は2026")
    ap.add_argument("sheetname", nargs="?", default=None, help="シート名。省略時は科目番号")
    ap.add_argument("--version", default=None, help="compile only one version, e.g. A or B")
    ap.add_argument("--runs", type=int, default=2, help="lualatex runs")
    args = ap.parse_args()

    sheet = args.sheetname if args.sheetname else args.subject_no

    excel_path, work_dir, exam_koma_no, sub_folder = get_exam_path(args.subject_no, args.year)

    json_path = work_dir / f"{sheet}.json"

    versions = load_versions_from_json_path(json_path)
    if args.version:
        versions = [args.version]

    for ver in versions:
        body_path = find_body_tex(work_dir, sheet, ver)

        # コンパイル先
        exam_dir = excel_path.parent
        out_dir = exam_dir / "pdf" / ver
        # out_dir = work_dir / "pdf" / ver
        out_dir.mkdir(parents=True, exist_ok=True)

        copy_templates_to(out_dir)

        inject_graphicspath(out_dir, sheet)

        body_name = f"{sheet}_{ver}_body.tex"
        dst_body = out_dir / body_name

        if body_path.resolve() != dst_body.resolve():
            shutil.copy2(body_path, dst_body)

        full_name = f"{sheet}_{ver}.tex"
        full_tex_path = build_full_tex(out_dir, body_name, full_name)
        print(f"✅ TeX merged: {full_tex_path}")

        pdf_path = compile_lualatex(full_tex_path, runs=args.runs)
        print(f"✅ PDF compiled: {pdf_path}")

    print("🎯 Done.")

if __name__ == "__main__":
    main()
