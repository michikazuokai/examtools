from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple


class ContractError(ValueError):
    """Raised when JSON does not satisfy the v2 contract."""


@dataclass
class Problem:
    level: str  # 'ERROR' or 'WARN'
    loc: str
    message: str

    def __str__(self) -> str:
        return f"{self.level}: {self.loc}: {self.message}"


# -----------------------------------------------------------------------------
# Contract definition
# -----------------------------------------------------------------------------

# Common metadata keys for traceability
COMMON_META_REQUIRED = {"tag", "src"}

# Types that are allowed to omit tag/src (pure metadata, not originating from a row)
META_TYPES_NO_SRC = {"metainfo"}


@dataclass(frozen=True)
class TypeSpec:
    required: Tuple[str, ...]
    optional: Tuple[str, ...] = ()
    # If True, unknown keys are warned (or errored in strict mode)
    warn_unknown_keys: bool = True
    # If True, tag/src are required (unless type is in META_TYPES_NO_SRC)
    require_meta: bool = True


CONTRACT: Dict[str, TypeSpec] = {
    # Plain text line / paragraph
    "text": TypeSpec(required=("type", "value"), optional=("class",)),
    # Alias often used in older JSON generators
    "sline": TypeSpec(required=("type", "value"), optional=("class",)),

    # Choices / selection
    # Canonical in your current JSON: style + values
    "choices": TypeSpec(required=("type", "style", "values"), optional=("sep", "shuffle")),

    # Code block
    "code": TypeSpec(required=("type", "lines"), optional=("linenumber", "language")),

    # Multiline (free-form multiple lines)
    "multiline": TypeSpec(required=("type", "values"), optional=("class",)),

    # Image block
    "image": TypeSpec(required=("type", "path"), optional=("width", "caption", "subimages")),

    # Spacing and pagination
    "vspace": TypeSpec(required=("type", "value_mm"), optional=()),
    "pagebreak": TypeSpec(required=("type",), optional=()),

    # Cover / header area (often a large block)
    "cover": TypeSpec(required=("type",), optional=(), warn_unknown_keys=False),

    # Meta information block (hash, createdatetime, etc.)
    "metainfo": TypeSpec(required=("type",), optional=(), warn_unknown_keys=False, require_meta=False),
}


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------

_STYLE_INLINE_RE = re.compile(r"^inline\((\d+)\)$")


def loc_from_src(src: Any) -> str:
    """Return location string like 'SHEET!R18-R23'."""
    if not isinstance(src, dict):
        return "?"
    sheet = src.get("sheet", "?")
    row = src.get("row")
    row_end = src.get("row_end")
    if isinstance(row, int) and isinstance(row_end, int):
        return f"{sheet}!R{row}-R{row_end}"
    if isinstance(row, int):
        return f"{sheet}!R{row}"
    return f"{sheet}!R?"


def element_loc(elem: Any) -> str:
    if isinstance(elem, dict) and "src" in elem:
        return loc_from_src(elem.get("src"))
    return "?"


def _is_non_empty_list(v: Any) -> bool:
    return isinstance(v, list) and len(v) > 0


def _as_int(v: Any) -> Optional[int]:
    if isinstance(v, int):
        return v
    if isinstance(v, str) and v.strip().lstrip("-").isdigit():
        try:
            return int(v.strip())
        except Exception:
            return None
    return None


# -----------------------------------------------------------------------------
# Normalization (compat + canonicalization)
# -----------------------------------------------------------------------------


def normalize_element(elem: Dict[str, Any], problems: Optional[List[Problem]] = None) -> Dict[str, Any]:
    """Normalize a single element in-place and return it.

    - choices.items -> choices.values (compat)
    - choices.style='inline(8)' -> style='inline' + sep=8
    - sline -> text (optional; we keep type but normalize keys)
    - vspace.value -> vspace.value_mm (compat)
    """
    if problems is None:
        problems = []

    etype = elem.get("type")
    loc = element_loc(elem)

    # compat: vspace value -> value_mm
    if etype == "vspace":
        if "value_mm" not in elem and "value" in elem:
            elem["value_mm"] = elem.pop("value")
            problems.append(Problem("WARN", loc, "vspace: renamed key 'value' -> 'value_mm'"))

    # compat + canonicalize: choices
    if etype == "choices":
        if "values" not in elem and "items" in elem:
            elem["values"] = elem.pop("items")
            problems.append(Problem("WARN", loc, "choices: renamed key 'items' -> 'values'"))

        style = elem.get("style")
        if isinstance(style, str):
            m = _STYLE_INLINE_RE.match(style)
            if m:
                sep = _as_int(m.group(1))
                elem["style"] = "inline"
                if sep is not None and "sep" not in elem:
                    elem["sep"] = sep
                problems.append(Problem("WARN", loc, f"choices: normalized style '{style}' -> 'inline'"))

    # No renaming of type here; keep your existing type names stable.
    return elem


def normalize_document(doc: Any) -> List[Problem]:
    """Walk the full JSON doc and normalize all elements (dicts with 'type')."""
    problems: List[Problem] = []

    def walk(node: Any) -> None:
        if isinstance(node, dict):
            if "type" in node:
                normalize_element(node, problems)
            for v in node.values():
                walk(v)
        elif isinstance(node, list):
            for it in node:
                walk(it)

    walk(doc)
    return problems


# -----------------------------------------------------------------------------
# Validation
# -----------------------------------------------------------------------------


def validate_element(elem: Dict[str, Any], strict: bool = False, warn_unknown_keys: bool = False) -> List[Problem]:
    """Validate a single element against the contract.

    Returns a list of Problem (WARN/ERROR). In strict mode, some WARN become ERROR.
    """
    problems: List[Problem] = []
    etype = elem.get("type")
    loc = element_loc(elem)

    if not isinstance(etype, str) or not etype:
        problems.append(Problem("ERROR", loc, "missing or invalid element type"))
        return problems

    spec = CONTRACT.get(etype)
    if spec is None:
        level = "ERROR" if strict else "WARN"
        problems.append(Problem(level, loc, f"unknown element type: '{etype}'"))
        return problems

    # meta requirement
    if spec.require_meta and etype not in META_TYPES_NO_SRC:
        for k in COMMON_META_REQUIRED:
            if k not in elem:
                problems.append(Problem("ERROR" if strict else "WARN", loc, f"{etype}: missing meta key: {k}"))

    # required keys
    for k in spec.required:
        if k not in elem:
            problems.append(Problem("ERROR", loc, f"{etype}: missing required key: {k}"))

    # unknown keys
    if warn_unknown_keys and spec.warn_unknown_keys:
        allowed = set(spec.required) | set(spec.optional) | {"tag", "src"}
        for k in elem.keys():
            if k not in allowed:
                level = "ERROR" if strict else "WARN"
                problems.append(Problem(level, loc, f"{etype}: unknown key: {k}"))

    # per-type constraints
    if etype in ("text", "sline"):
        if "value" in elem and not isinstance(elem["value"], str):
            problems.append(Problem("ERROR", loc, f"{etype}: value must be string"))

    if etype == "choices":
        style = elem.get("style")
        if style not in ("normal", "inline"):
            # allow 'inline(8)' before normalization
            if not (isinstance(style, str) and _STYLE_INLINE_RE.match(style)):
                problems.append(Problem("ERROR", loc, "choices: style must be 'normal' or 'inline'"))

        values = elem.get("values")
        if not _is_non_empty_list(values):
            problems.append(Problem("ERROR", loc, "choices: values must be non-empty list"))
        else:
            for idx, opt in enumerate(values):
                opt_loc = loc
                if isinstance(opt, dict) and "src" in opt:
                    opt_loc = loc_from_src(opt["src"])
                if not isinstance(opt, dict):
                    problems.append(Problem("ERROR", opt_loc, f"choices: option[{idx}] must be object"))
                    continue
                if "text" not in opt or not isinstance(opt.get("text"), str) or not opt.get("text"):
                    problems.append(Problem("ERROR", opt_loc, f"choices: option[{idx}] missing text"))
                # label is recommended; in strict mode require it
                if strict:
                    if "label" not in opt or not isinstance(opt.get("label"), str) or not opt.get("label"):
                        problems.append(Problem("ERROR", opt_loc, f"choices: option[{idx}] missing label"))
                # meta on option is recommended; warn only
                if "src" not in opt:
                    problems.append(Problem("WARN", opt_loc, f"choices: option[{idx}] missing src"))

        if "sep" in elem:
            if _as_int(elem["sep"]) is None:
                problems.append(Problem("ERROR", loc, "choices: sep must be int"))

    if etype == "code":
        lines = elem.get("lines")
        if not _is_non_empty_list(lines):
            problems.append(Problem("ERROR", loc, "code: lines must be non-empty list"))
        else:
            bad = [i for i, s in enumerate(lines) if not isinstance(s, str)]
            if bad:
                problems.append(Problem("ERROR", loc, f"code: lines[{bad[0]}] is not string"))
        if "linenumber" in elem and not isinstance(elem["linenumber"], bool):
            problems.append(Problem("ERROR", loc, "code: linenumber must be boolean"))

    if etype == "multiline":
        values = elem.get("values")
        if not _is_non_empty_list(values):
            problems.append(Problem("ERROR", loc, "multiline: values must be non-empty list"))
        else:
            bad = [i for i, s in enumerate(values) if not isinstance(s, str)]
            if bad:
                problems.append(Problem("ERROR", loc, f"multiline: values[{bad[0]}] is not string"))

    if etype == "image":
        path = elem.get("path")
        if path is not None and not isinstance(path, str):
            problems.append(Problem("ERROR", loc, "image: path must be string"))
        if "width" in elem and not isinstance(elem["width"], (int, float)):
            problems.append(Problem("ERROR", loc, "image: width must be number"))

    if etype == "vspace":
        if "value_mm" in elem:
            if _as_int(elem["value_mm"]) is None:
                problems.append(Problem("ERROR", loc, "vspace: value_mm must be int"))

    return problems


def validate_document(doc: Any, strict: bool = False, warn_unknown_keys: bool = False) -> List[Problem]:
    """Validate the whole JSON document.

    - Walks recursively and validates every dict that has a 'type' key.
    - Returns all Problems.
    - Raises ContractError if any ERROR exists.
    """
    problems: List[Problem] = []

    def walk(node: Any) -> None:
        if isinstance(node, dict):
            if "type" in node:
                problems.extend(validate_element(node, strict=strict, warn_unknown_keys=warn_unknown_keys))
            for v in node.values():
                walk(v)
        elif isinstance(node, list):
            for it in node:
                walk(it)

    walk(doc)

    has_error = any(p.level == "ERROR" for p in problems)
    if has_error:
        raise ContractError("\n".join(str(p) for p in problems if p.level == "ERROR"))
    return problems
