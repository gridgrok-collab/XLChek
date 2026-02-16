"""
sheetguard/core/references.py

Deterministic formula tokenization for dependency graph building.

Design goal (v1):
- Extract *referenced* addresses from Excel formulas without evaluating formulas.
- Return a simple, stable shape used across CLI/UI:
    List[dict] tokens with keys:
      - raw: str               (e.g., "A1", "A1:B5", "A:A", "1:1", "Table1[Col]")
      - sheet: Optional[str]   (sheet name if explicitly referenced, else None)
      - kind: str              ("cell" | "range" | "whole_row" | "whole_column" | "symbolic")
      - symbolic: bool         (True for named ranges / structured refs)
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

# Excel volatile functions (non-exhaustive but practical for v1)
_VOLATILE_FUNCS = {
    "NOW", "TODAY", "RAND", "RANDBETWEEN",
    "OFFSET", "INDIRECT", "CELL", "INFO",
}

# Matches:
#   Sheet1!A1
#   'My Sheet'!A1:B5
_SHEET_PREFIX = r"(?:'(?P<qs>[^']+)'|(?P<us>[^'!]+))!"
_CELL = r"\$?[A-Z]{1,3}\$?\d{1,7}"
_RANGE = rf"{_CELL}:{_CELL}"
_WHOLE_COL = r"\$?[A-Z]{1,3}:\$?[A-Z]{1,3}"
_WHOLE_ROW = r"\$?\d{1,7}:\$?\d{1,7}"

_RE_SHEET_RANGE = re.compile(rf"{_SHEET_PREFIX}(?P<addr>{_RANGE}|{_CELL}|{_WHOLE_COL}|{_WHOLE_ROW})")
_RE_BARE_RANGE = re.compile(rf"(?P<addr>{_RANGE}|{_CELL}|{_WHOLE_COL}|{_WHOLE_ROW})")

# Structured refs / named ranges:
#   Table1[Col]
_RE_STRUCTURED = re.compile(r"\b(?P<name>[A-Za-z_][A-Za-z0-9_]*\[[^\]]+\])\b")
# Best-effort name tokens (avoid function calls by checking next char)
_RE_NAME = re.compile(r"\b(?P<name>[A-Za-z_][A-Za-z0-9_\.]{0,255})\b")


@dataclass(frozen=True)
class RefToken:
    raw: str
    sheet: Optional[str]
    kind: str
    symbolic: bool = False

    @property
    def is_symbolic(self) -> bool:
        """Backward-compatible alias for `symbolic`."""
        return bool(self.symbolic)

    def to_dict(self) -> Dict[str, Any]:
        return {"raw": self.raw, "sheet": self.sheet, "kind": self.kind, "symbolic": bool(self.symbolic)}

    @staticmethod
    def from_any(obj: Any) -> "RefToken":
        # Backward-compat: accept dicts
        if isinstance(obj, RefToken):
            return obj
        if isinstance(obj, dict):
            return RefToken(
                raw=str(obj.get("raw") or obj.get("ref") or ""),
                sheet=obj.get("sheet"),
                kind=str(obj.get("kind") or "symbolic"),
                symbolic=bool(obj.get("symbolic") or obj.get("is_symbolic") or False),
            )
        return RefToken(raw=str(obj), sheet=None, kind="symbolic", symbolic=True)


def detect_volatile_functions(formula: str) -> List[str]:
    """
    Returns a list of volatile function names found in the formula.
    """
    if not formula:
        return []
    f = formula.upper()
    hits: List[str] = []
    for fn in _VOLATILE_FUNCS:
        if f.find(fn + "(") != -1:
            hits.append(fn)
    hits.sort()
    return hits


def _kind_from_addr(addr: str) -> str:
    if re.fullmatch(_CELL, addr):
        return "cell"
    if re.fullmatch(_RANGE, addr):
        return "range"
    if re.fullmatch(_WHOLE_COL, addr):
        return "whole_column"
    if re.fullmatch(_WHOLE_ROW, addr):
        return "whole_row"
    return "symbolic"


def extract_reference_tokens(formula: str) -> List[Dict[str, Any]]:
    """
    Extract reference tokens from an Excel formula string.
    Deterministic parsing only (no eval), best-effort for v1.
    """
    if not formula:
        return []

    f = formula[1:] if formula.startswith("=") else formula

    tokens: List[RefToken] = []

    # 1) Structured references like Table1[Col]
    for m in _RE_STRUCTURED.finditer(f):
        name = m.group("name")
        tokens.append(RefToken(raw=name, sheet=None, kind="symbolic", symbolic=True))

    # 2) Sheet-qualified addresses
    for m in _RE_SHEET_RANGE.finditer(f):
        sheet = m.group("qs") or m.group("us")
        addr = m.group("addr")
        tokens.append(RefToken(raw=addr, sheet=sheet, kind=_kind_from_addr(addr), symbolic=False))

    # 3) Bare addresses (no explicit sheet)
    for m in _RE_BARE_RANGE.finditer(f):
        addr = m.group("addr")
        tokens.append(RefToken(raw=addr, sheet=None, kind=_kind_from_addr(addr), symbolic=False))

    # 4) Named ranges (best-effort): word tokens that are not functions.
    for m in _RE_NAME.finditer(f):
        name = m.group("name")
        after = f[m.end():m.end() + 1]
        if after == "(":
            continue
        if re.fullmatch(_CELL, name) or re.fullmatch(_WHOLE_COL, name) or re.fullmatch(_WHOLE_ROW, name):
            continue
        if name.upper() in {"TRUE", "FALSE", "NA", "N", "PI", "E"}:
            continue
        if name.upper() in _VOLATILE_FUNCS:
            continue
        tokens.append(RefToken(raw=name, sheet=None, kind="symbolic", symbolic=True))

    # De-duplicate while preserving order
    seen = set()
    out: List[Dict[str, Any]] = []
    for t in tokens:
        key = (t.raw, t.sheet, t.kind, t.symbolic)
        if key in seen:
            continue
        seen.add(key)
        out.append(t.to_dict())

    return out