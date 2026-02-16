"""Hard-coded constants detector.

Detects numeric literals embedded inside Excel formulas that may represent
hidden business assumptions (e.g., *1.05, /365, NPV(0.08,...)).

Design goals:
- Read-only, deterministic.
- Config-driven (no magic thresholds in code).
- Conservative: report potential embedded assumptions, not definitive errors.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Dict, Iterable, List, Optional, Tuple


_NUM_RE = re.compile(
    r"""(?xi)
    (?P<num>
        [+-]?                    # unary sign
        (?:\d+(?:\.\d*)?|\.\d+) # int/float
        (?:e[+-]?\d+)?           # scientific
    )
    """
)


@dataclass(frozen=True)
class ConstantHit:
    literal: str
    value: float
    operator: str
    function: str
    start: int
    end: int

    def to_dict(self, sheet: str, cell: str, formula: str) -> Dict[str, Any]:
        return {
            "type": "HARDCODED_CONSTANT",
            "sheet": sheet,
            "cell": cell,
            "literal": self.literal,
            "value": self.value,
            "operator": self.operator,
            "function": self.function,
            "formula": formula,
        }


def _strip_quoted_segments(s: str) -> str:
    """Replace characters inside double-quotes with spaces to preserve indices."""
    if '"' not in s:
        return s
    out = list(s)
    in_q = False
    for i, ch in enumerate(out):
        if ch == '"':
            in_q = not in_q
            continue
        if in_q:
            out[i] = " "
    return "".join(out)


def _looks_like_cell_ref_near(s: str, start: int, end: int) -> bool:
    """Heuristic: avoid matching row numbers in refs like A10 or $B$12."""
    prev = s[start - 1] if start - 1 >= 0 else ""
    nxt = s[end] if end < len(s) else ""

    # If preceded by a letter or '$', likely part of a cell ref (A10, $12 is less likely).
    if prev.isalpha() or prev == "$":
        return True
    # If followed by a letter, could be a name token like 1E3A (rare) — treat as not a pure number.
    if nxt.isalpha() or nxt == "$":
        return True
    return False


def _nearest_operator(s: str, start: int) -> str:
    """Return the nearest non-space operator directly adjacent to the literal."""
    # Check just-left then just-right for operators
    i = start - 1
    while i >= 0 and s[i].isspace():
        i -= 1
    if i >= 0 and s[i] in "*/+-":
        return s[i]

    j = start
    while j < len(s) and s[j].isspace():
        j += 1
    if j < len(s) and s[j] in "*/+-":
        return s[j]
    return ""


def _enclosing_function(formula: str, pos: int) -> str:
    """Best-effort: find the function name whose parentheses enclose pos."""
    # Walk left, track paren depth, pick the nearest NAME( that opens the current depth.
    s = formula
    depth = 0
    i = pos
    while i >= 0:
        ch = s[i]
        if ch == ")":
            depth += 1
        elif ch == "(":
            if depth == 0:
                # Extract function name immediately before this '(' (letters + dots/underscores)
                j = i - 1
                while j >= 0 and (s[j].isalpha() or s[j] in "._"):
                    j -= 1
                name = s[j + 1 : i].strip()
                return name.upper()
            depth -= 1
        i -= 1
    return ""


def detect_hardcoded_constants(
    formula: str,
    *,
    cfg: Optional[Dict[str, Any]] = None,
) -> List[ConstantHit]:
    """Return ConstantHit items for numeric literals that are likely assumptions.

    Config (all optional):
      ignore_literals: list[number]
      ignore_functions: list[str]
      risky_operators: list[str]
      risky_functions: list[str]
    """
    cfg = cfg or {}
    ignore_literals = set(str(x) for x in (cfg.get("ignore_literals") or []))
    ignore_functions = set(str(x).upper() for x in (cfg.get("ignore_functions") or []))
    # IMPORTANT DEFAULTING:
    # If risky_operators is missing/empty due to a mis-loaded config (e.g.,
    # wrong rules.yaml path or wrong nesting), we must NOT degrade into
    # flagging *all* numeric literals (including +1, +100, etc.).
    #
    # Industry-default conservative behavior: only treat "*" and "/" adjacency
    # as embedded-assumption signals.
    risky_operators = set(str(x) for x in (cfg.get("risky_operators") or []))
    if not risky_operators:
        risky_operators = {"*", "/"}
    risky_functions = set(str(x).upper() for x in (cfg.get("risky_functions") or []))

    if not formula or not isinstance(formula, str):
        return []

    # Work on a version with quoted text blanked out.
    s = _strip_quoted_segments(formula)

    hits: List[ConstantHit] = []
    for m in _NUM_RE.finditer(s):
        lit = m.group("num")
        if not lit:
            continue
        start, end = m.start("num"), m.end("num")

        if _looks_like_cell_ref_near(s, start, end):
            continue

        # Normalize string match for ignore list (support '1' vs '1.0')
        try:
            val = float(lit)
        except Exception:
            continue

        # Ignore-list should match regardless of an explicit unary sign.
        # Example: if config ignores "1" or "100", also ignore "+1"/"-1" and "+100"/"-100".
        lit_unsigned = lit.lstrip("+-")

        if lit in ignore_literals or lit_unsigned in ignore_literals:
            continue

        if abs(val).is_integer():
            if str(int(abs(val))) in ignore_literals:
                continue

        op = _nearest_operator(s, start)
        func = _enclosing_function(s, start)

        # If the literal appears in a "safe" function, suppress.
        if func and func in ignore_functions:
            continue

        # Determine if we should flag: risky operator adjacency OR risky function argument.
        should_flag = False
        if op and op in risky_operators:
            should_flag = True
        if func and func in risky_functions:
            should_flag = True

        if not should_flag:
            continue

        hits.append(ConstantHit(literal=lit, value=val, operator=op, function=func, start=start, end=end))

    return hits
