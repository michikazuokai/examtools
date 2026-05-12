#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
試験問題生成フローを一括実行するスクリプト。

実行例:
    python scripts/make_all.py 2031002

実行順:
    1. validate_excel.py
    2. make_json.py
    3. make_latex.py
    4. make_pdf.py
    5. make_anspdf.py

途中でエラーが発生した場合は、その段階で停止する。
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def run_step(step_no: int, title: str, script_path: Path, subject: str) -> None:
    """
    1つのスクリプトを実行する。
    エラーが出た場合は例外を送出する。
    """
    cmd = [sys.executable, str(script_path), subject]

    print()
    print("=" * 60)
    print(f"Step {step_no}: {title}")
    print("=" * 60)
    print("実行:", " ".join(cmd))

    result = subprocess.run(cmd)

    if result.returncode != 0:
        raise RuntimeError(
            f"{title} でエラーが発生しました。\n"
            f"停止したStep : {step_no}\n"
            f"停止した処理 : {script_path.name}\n"
            f"終了コード   : {result.returncode}"
        )

    print(f"✅ Step {step_no} 完了: {title}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="試験問題生成フローを一括実行します。"
    )
    parser.add_argument("subject", help="科目番号。例: 2031002")
    args = parser.parse_args()

    subject = str(args.subject)

    # このスクリプト自身が scripts フォルダにある前提
    scripts_dir = Path(__file__).resolve().parent

    steps = [
        (1, "Excelチェック・補正", scripts_dir / "validate_excel.py"),
        (2, "JSON作成", scripts_dir / "make_json.py"),
        (3, "LaTeX本文作成", scripts_dir / "make_latex.py"),
        (4, "PDF作成", scripts_dir / "make_pdf.py"),
        (5, "解答用紙PDF作成", scripts_dir / "make_anspdf.py"),
    ]

    print(f"科目番号: {subject}")
    print("試験問題生成フローを開始します。")

    for step_no, title, script_path in steps:
        if not script_path.exists():
            raise FileNotFoundError(
                f"実行対象のスクリプトが見つかりません。\n"
                f"Step {step_no}: {title}\n"
                f"Path: {script_path}"
            )

        run_step(step_no, title, script_path, subject)

    print()
    print("=" * 60)
    print("🎯 全工程が正常終了しました。")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print()
        print("🔥 一括実行を停止しました")
        print("-" * 60)
        print(e)
        print("-" * 60)
        raise SystemExit(1)