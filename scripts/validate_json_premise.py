#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""validate_json_v2.py

JSON(v2) contract validator for the pipeline:
  Excel -> JSON(v2) -> LaTeX(v2) -> PDF

Even if you already validate Excel, you still want this because:
- Excel validation ensures *input* is sane (tags/columns/order).
- This validator ensures the *output JSON* matches what json->latex expects.
  (i.e., it catches converter bugs/spec drift *before* LaTeX compile.)

Key features
- Validates element objects by `type`.
- Enforces presence of `tag` and `src{sheet,row[,row_end]}` on elements.
- Emits errors with Excel location when possible.
- Traverses unknown top-level layouts by finding dicts that look like elements.

Usage:
  python validate_json_v2.py path/to/doc.json
  python validate_json_v2.py --strict path/to/doc.json

Exit code:
  0 if no errors (warnings may exist)
  2 if errors exist
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple


# -----------------------------
# Contract (schema-ish) section
# -----------------------------

# NOTE:
# - We deliberately keep the contract minimal to start.
# - Add types/keys as your JSON grows.
# - The validator supports WARNING mode for unknown keys.

COMMON_REQUIRED_KEYS = {"type", "tag", "src"}

# Some element-like objects in your JSON may not originate from a single Excel row
# (e.g., cover metadata). For these types, we don't require tag/src.
NO_SRC_REQUIRED_TYPES = {"cover", "metainfo"}

# Per-element-type required / optional keys.
# "src" should be a dict: {sheet,row[,row_end]}
CONTRACT: Dict[str, Dict[str, Any]] = {
    # Cover / header block (often a container-like element)
    # Keep this permissive so it doesn't get in the way of validating real elements.
    "cover": {
        "required": set(),
        "optional": {
            "title",
            "notes",
            "instruction",  # legacy alias
            "footer_text",
            "subtitle",
            "kaito_message",
            "subject",
            "fsyear",
            "term",
            "qversion",
            "width",
            "height",
            "metainfo",
        },
        "enums": {},
    },
    # Plain text line / paragraph
    "text": {
        "required": {"value"},
        "optional": {"role"},  # e.g., teacher-only, etc. (if you use it)
        "enums": {},
    },
    # Alternative naming sometimes used
    "sline": {
        "required": {"value"},
        "optional": {"role"},
        "enums": {},
    },
    # Multiple lines
    "multiline": {
        "required": {"values"},
        "optional": {"align", "indent"},
        "enums": {
            # Optional if you use alignment
            "align": {"left", "center", "right"},
        },
    },
    # Code block
    "code": {
        "required": {"lines"},
        "optional": {"language", "linenumber", "caption"},
        "enums": {
            "linenumber": {True, False},
        },
    },
    # Choices / multiple choice
    # Canonical form in your current JSON uses: style + values[]
    # (values items are usually dicts with {label,text,...}).
    "choices": {
        "required": {"values", "style"},
        "optional": {"answer", "shuffle", "sep"},
        "enums": {
            "style": {"normal", "inline"},
        },
    },
    # Images
    "image": {
        "required": {"path"},
        "optional": {"width", "caption"},
        "enums": {},
    },
    "subimage": {
        "required": {"path"},
        "optional": {"width", "caption"},
        "enums": {},
    },
    # Spacing
    "vspace": {
        "required": {"value_mm"},
        "optional": {},
        "enums": {},
    },
    # Pagebreak (if represented in JSON)
    "pagebreak": {
        "required": set(),
        "optional": {},
        "enums": {},
    },
    # Meta info / header elements (if you keep them)
    "metainfo": {
        "required": {"hash"},
        "optional": {"createdatetime", "verno", "inputpath", "sheetname"},
        "enums": {},
    },
}

# If you use nested question objects as elements, you can add a contract here.
# But usually questions are containers with an "elements" list and are not
# treated as leaf elements.


# -----------------------------
# Validation infrastructure
# -----------------------------

@dataclass
class Issue:
    level: str  # 'ERROR' or 'WARN'
    message: str
    loc: str  # e.g., Sheet!R123 or JSON path


def format_loc(elem: Optional[Dict[str, Any]], fallback: str = "?") -> str:
    if not isinstance(elem, dict):
        return fallback
    src = elem.get("src")
    if not isinstance(src, dict):
        return fallback
    sheet = src.get("sheet", "?")
    row = src.get("row", "?")
    row_end = src.get("row_end")
    if row_end is not None:
        return f"{sheet}!R{row}-R{row_end}"
    return f"{sheet}!R{row}"


def is_element_dict(d: Dict[str, Any]) -> bool:
    # Heuristic: an element has a string 'type'.
    t = d.get("type")
    return isinstance(t, str) and len(t) > 0


def iter_elements(obj: Any, path: str = "$") -> Iterable[Tuple[Dict[str, Any], str]]:
    """Traverse JSON and yield (element_dict, json_path) for dicts that look like elements."""
    if isinstance(obj, dict):
        if is_element_dict(obj):
            yield obj, path
        for k, v in obj.items():
            yield from iter_elements(v, f"{path}.{k}")
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            yield from iter_elements(v, f"{path}[{i}]")


def validate_src(elem: Dict[str, Any], strict: bool, issues: List[Issue], path: str) -> None:
    # For certain element types (e.g., cover/metainfo), src/tag may be omitted.
    t = elem.get("type")
    if isinstance(t, str) and t in NO_SRC_REQUIRED_TYPES:
        return
    loc = format_loc(elem, fallback=path)
    if "tag" not in elem:
        issues.append(Issue("ERROR" if strict else "WARN", "missing key: tag", loc))
    if "src" not in elem:
        issues.append(Issue("ERROR" if strict else "WARN", "missing key: src", loc))
        return

    src = elem.get("src")
    if not isinstance(src, dict):
        issues.append(Issue("ERROR", "src must be an object/dict", loc))
        return

    if "sheet" not in src or not isinstance(src.get("sheet"), str) or not src.get("sheet"):
        issues.append(Issue("ERROR", "src.sheet must be a non-empty string", loc))
    if "row" not in src or not isinstance(src.get("row"), int):
        # allow numeric-but-not-int? keep strict.
        issues.append(Issue("ERROR", "src.row must be an integer", loc))
    if "row_end" in src and not isinstance(src.get("row_end"), int):
        issues.append(Issue("ERROR", "src.row_end must be an integer when present", loc))


def validate_type_contract(elem: Dict[str, Any], strict: bool, warn_unknown_keys: bool, issues: List[Issue], path: str) -> None:
    t = elem.get("type")
    loc = format_loc(elem, fallback=path)
    if not isinstance(t, str):
        issues.append(Issue("ERROR", "type must be a string", loc))
        return

    # --- light normalization (backward compatibility) ---
    # Your current canonical choices form is: style + values[].
    # If older JSON uses 'items', normalize to 'values' with a warning.
    if t == "choices":
        if "values" not in elem and isinstance(elem.get("items"), list):
            issues.append(Issue("WARN" if not strict else "ERROR", "choices: 'items' is deprecated; use 'values'", loc))
            elem["values"] = elem.get("items")
        # If older JSON used opts{type,sep}, migrate to style/sep.
        if "style" not in elem and isinstance(elem.get("opts"), dict):
            o = elem.get("opts")
            if isinstance(o.get("type"), str):
                issues.append(Issue("WARN" if not strict else "ERROR", "choices: 'opts.type' is deprecated; use 'style'", loc))
                elem["style"] = o.get("type")
            if "sep" not in elem and isinstance(o.get("sep"), (int, float)):
                issues.append(Issue("WARN" if not strict else "ERROR", "choices: 'opts.sep' is deprecated; use top-level 'sep'", loc))
                elem["sep"] = o.get("sep")
    if t in ("image", "subimage"):
        if "path" not in elem and isinstance(elem.get("file"), str):
            issues.append(Issue("WARN" if not strict else "ERROR", f"{t}: 'file' is deprecated; use 'path'", loc))
            elem["path"] = elem.get("file")

    if t not in CONTRACT:
        # Unknown element type. Decide policy:
        msg = f"unknown element type: {t!r}"
        issues.append(Issue("ERROR" if strict else "WARN", msg, loc))
        return

    # 'type' is always required. 'tag'/'src' are required for most elements,
    # but some container/meta objects (e.g., cover/metainfo) may omit them.
    required = set(CONTRACT[t].get("required", set())) | {"type"}
    if t not in NO_SRC_REQUIRED_TYPES:
        required |= {"tag", "src"}
    optional = set(CONTRACT[t].get("optional", set()))
    enums = dict(CONTRACT[t].get("enums", {}))

    # Required keys
    for k in required:
        if k not in elem:
            issues.append(Issue("ERROR", f"{t}: missing required key: {k}", loc))

    # Key sets
    allowed = required | optional
    unknown = set(elem.keys()) - allowed
    if unknown and warn_unknown_keys:
        issues.append(Issue("WARN" if not strict else "ERROR", f"{t}: unknown keys: {sorted(unknown)}", loc))

    # Type-specific structural checks
    if t in ("text", "sline"):
        if "value" in elem and not isinstance(elem.get("value"), str):
            issues.append(Issue("ERROR", f"{t}: value must be string", loc))

    if t == "multiline":
        v = elem.get("values")
        if not isinstance(v, list) or not all(isinstance(x, str) for x in v):
            issues.append(Issue("ERROR", "multiline: values must be list[str]", loc))

    if t == "code":
        lines = elem.get("lines")
        if not isinstance(lines, list) or not all(isinstance(x, str) for x in lines):
            issues.append(Issue("ERROR", "code: lines must be list[str]", loc))
        ln = elem.get("linenumber")
        if ln is not None and not isinstance(ln, bool):
            issues.append(Issue("ERROR", "code: linenumber must be bool when present", loc))

    if t == "choices":
        values = elem.get("values")
        if not isinstance(values, list) or len(values) == 0:
            issues.append(Issue("ERROR", "choices: values must be non-empty list", loc))
        else:
            for idx, it in enumerate(values):
                if isinstance(it, str):
                    # allow simplest representation: list[str]
                    continue
                if not isinstance(it, dict):
                    issues.append(Issue("ERROR", f"choices: values[{idx}] must be dict or str", loc))
                    continue
                # Canonical item shape: {label,text,...}
                if "text" not in it:
                    issues.append(Issue("ERROR", f"choices: values[{idx}] missing text", loc))
                lbl = it.get("label")
                if lbl is not None and not isinstance(lbl, str):
                    issues.append(Issue("ERROR", f"choices: values[{idx}].label must be string when present", loc))
                if "src" in it and not isinstance(it.get("src"), dict):
                    issues.append(Issue("ERROR", f"choices: values[{idx}].src must be dict when present", loc))

        # sep (mm) is optional for inline layouts
        sep = elem.get("sep")
        if sep is not None and not isinstance(sep, int):
            issues.append(Issue("ERROR", "choices.sep must be int (mm) when present", loc))

    if t in ("image", "subimage"):
        p = elem.get("path")
        if p is not None and not isinstance(p, str):
            issues.append(Issue("ERROR", f"{t}: path must be string", loc))
        w = elem.get("width")
        if w is not None and not isinstance(w, (int, float)):
            issues.append(Issue("ERROR", f"{t}: width must be number (ratio)", loc))

    if t == "vspace":
        mm = elem.get("value_mm")
        if not isinstance(mm, int):
            issues.append(Issue("ERROR", "vspace: value_mm must be int (mm)", loc))

    if t == "metainfo":
        h = elem.get("hash")
        if h is not None and not isinstance(h, str):
            issues.append(Issue("ERROR", "metainfo: hash must be string", loc))

    # Enums checks (generic)
    for k, allowed_vals in enums.items():
        if k in elem and elem[k] is not None and elem[k] not in allowed_vals:
            issues.append(Issue("ERROR", f"{t}: {k} must be one of {sorted(allowed_vals)}", loc))


def validate_document(doc: Any, strict: bool, warn_unknown_keys: bool) -> List[Issue]:
    issues: List[Issue] = []

    # Walk all element dicts
    for elem, path in iter_elements(doc):
        # First, common src/tag checks
        validate_src(elem, strict=strict, issues=issues, path=path)
        # Then, per-type contract checks
        validate_type_contract(elem, strict=strict, warn_unknown_keys=warn_unknown_keys, issues=issues, path=path)

    return issues


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Validate exam JSON(v2) contract.")
    ap.add_argument("json_path", type=str, help="Path to JSON file")
    ap.add_argument("--strict", action="store_true", help="Treat unknown types and missing tag/src as errors")
    ap.add_argument(
        "--warn-unknown-keys",
        action="store_true",
        help="Warn (or error in --strict) on unknown keys within known element types",
    )
    args = ap.parse_args(argv)

    p = Path(args.json_path)
    if not p.exists():
        print(f"ERROR: file not found: {p}", file=sys.stderr)
        return 2

    try:
        doc = json.loads(p.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"ERROR: failed to read JSON: {e}", file=sys.stderr)
        return 2

    issues = validate_document(doc, strict=args.strict, warn_unknown_keys=args.warn_unknown_keys)

    # Print issues
    errs = 0
    warns = 0
    for it in issues:
        if it.level == "ERROR":
            errs += 1
        else:
            warns += 1
        print(f"{it.level}: {it.loc}: {it.message}")

    if errs == 0:
        print(f"OK: no errors. warnings={warns}")
        return 0

    print(f"NG: errors={errs} warnings={warns}")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
