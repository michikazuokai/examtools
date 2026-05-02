#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
make_pdf.py  (V2)
  body.tex + templates -> full .tex -> lualatex compile -> pdf

This script does NOT:
  - generate json
  - generate body.tex from json

It only:
  - reads work/{subject}.json
  - finds work/latex/{version}/{subject}_{version}_body.tex
  - checks source_excel_hash between JSON and body.tex
  - copies templates into exam_dir/pdf/{version}/
  - injects graphicspath for exam_dir/images/
  - creates full tex
  - runs lualatex

Usage:
  python scripts/make_pdf.py 1020201
  python scripts/make_pdf.py 1020201 --runs 2
"""

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path
from typing import List, Optional
from utils import add_subject_arg, load_exam_context

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


def find_body_tex(work_dir: Path, sheet: str, version: str) -> Path:
    body_path = work_dir / "latex" / version / f"{sheet}_{version}_body.tex"

    if body_path.exists():
        return body_path

    raise FileNotFoundError(f"body.tex not found: {body_path}")

def load_json_data(json_path: Path) -> dict:
    """
    JSONファイルを読み込む。
    """
    if not json_path.exists():
        raise FileNotFoundError(
            f"JSONファイルが見つかりません。\n"
            f"先に make_json.py を実行してください。\n"
            f"JSON path: {json_path}"
        )

    return json.loads(json_path.read_text(encoding="utf-8"))


def get_json_source_hash_by_version(data: dict, version: str) -> str:
    """
    JSON内の指定versionから source_excel_hash を取得する。
    """
    for block in data.get("versions") or []:
        if str(block.get("version")) != str(version):
            continue

        metainfo = block.get("metainfo", {}) or {}
        source_hash = metainfo.get("source_excel_hash") or metainfo.get("hash")

        if not source_hash:
            raise RuntimeError(
                f"JSONの version={version} に source_excel_hash がありません。\n"
                f"先に make_json.py を再実行してください。"
            )

        return str(source_hash)

    raise RuntimeError(f"JSON内に version={version} が見つかりません。")


def get_body_tex_source_hash(body_path: Path) -> str:
    """
    body.tex 冒頭コメントから source_excel_hash を取得する。

    make_latex.py 側で次のようなコメントを出している前提:
      % source_excel_hash: xxxxx
    """
    if not body_path.exists():
        raise FileNotFoundError(f"body.tex が見つかりません: {body_path}")

    for line in body_path.read_text(encoding="utf-8").splitlines()[:30]:
        line = line.strip()
        if line.startswith("% source_excel_hash:"):
            value = line.split(":", 1)[1].strip()
            if value:
                return value

    raise RuntimeError(
        "body.tex に source_excel_hash コメントがありません。\n"
        "先に make_latex.py を再実行してください。\n"
        f"body.tex: {body_path}"
    )


def require_body_matches_json(body_path: Path, json_data: dict, version: str) -> str:
    """
    body.tex の source_excel_hash と JSON metainfo の source_excel_hash が一致するか確認する。
    """
    json_hash = get_json_source_hash_by_version(json_data, version)
    body_hash = get_body_tex_source_hash(body_path)

    if json_hash != body_hash:
        raise RuntimeError(
            "body.tex が現在のJSONと一致しません。\n"
            "先に make_latex.py を再実行してください。\n"
            f"version   : {version}\n"
            f"JSON hash : {json_hash}\n"
            f"TeX hash  : {body_hash}\n"
            f"body.tex  : {body_path}"
        )

    return body_hash


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
    preamble.tex に \\graphicspath を注入する。
    今回の出力構造:
      exam_dir/
        images/
        pdf/
          A/
    なので、pdf/A から images へは ../../images/
    """
    preamble_path = build_dir / "preamble.tex"
    if not preamble_path.exists():
        raise FileNotFoundError(f"preamble.tex not found in build_dir: {preamble_path}")

    gsp = r"\graphicspath{{../../images/}}" + "\n"

    preamble = read_text(preamble_path)

    if "@@GRAPHICSPATH@@" in preamble:
        preamble = preamble.replace("@@GRAPHICSPATH@@", gsp)
    else:
        # すでに graphicspath がある場合は二重追加しない
        if r"\graphicspath" not in preamble:
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


def get_versions_from_json_data(data: dict) -> list[str]:
    versions = []
    for block in data.get("versions", []):
        ver = block.get("version")
        if ver:
            versions.append(str(ver))

    if not versions:
        raise RuntimeError("JSON内に versions が見つかりません。")

    return versions

def main() -> None:
    ap = argparse.ArgumentParser(description="LaTeX本文からPDFを作成します。")
    add_subject_arg(ap)
    ap.add_argument("--runs", type=int, default=2, help="lualatex runs")
    args = ap.parse_args()

    exam_context = load_exam_context(args.subject, load_workbook=False)

    subject = exam_context.subject
    sheet = exam_context.sheetname
    work_dir = exam_context.work_dir
    exam_dir = exam_context.exam_dir

    json_path = work_dir / f"{subject}.json"

    print(f"科目番号: {exam_context.subject}")
    print(f"年度: {exam_context.fsyear}")
    print(f"シート名: {exam_context.sheetname}")
    print(f"試験コマ番号: {exam_context.exam_koma_no}")
    print(f"入力JSON: {json_path}")
    print(f"入力TeX: {work_dir / 'latex'}")
    print(f"出力PDF: {exam_dir / 'pdf'}")


    json_data = load_json_data(json_path)
    versions = get_versions_from_json_data(json_data)

    print(f"出力版: {','.join(versions)}")

    for ver in versions:
        body_path = find_body_tex(work_dir, sheet, ver)

        # hashチェック：body.tex が JSON と同じExcel由来か確認する
        source_hash = require_body_matches_json(body_path, json_data, ver)
        print(f"source_excel_hash({ver}): {source_hash}")

        # コンパイル先
        out_dir = exam_dir / "pdf" / ver
        out_dir.mkdir(parents=True, exist_ok=True)

        copy_templates_to(out_dir)

        # 画像パス注入
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
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        if "--debug" in sys.argv:
            import traceback
            traceback.print_exc()
        else:
            print()
            print("🙅🏻‍♂️ エラー:")
            print(e)
        raise SystemExit(1)
