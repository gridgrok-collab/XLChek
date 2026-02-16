# sheetguard/detectors/drift.py
from __future__ import annotations

import re
from collections import Counter, defaultdict
from typing import Any, Dict, List, Tuple


_CELL_REF_RE = re.compile(r"(?:(?P<sheet>'[^']+'|[A-Za-z0-9_]+)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?")
_NUM_RE = re.compile(r"(?<![A-Za-z_])[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?")
_WS_RE = re.compile(r"\s+")


def normalize_formula(formula: str) -> str:
    """
    Normalize a formula to a stable 'shape' for drift detection.
    - Replaces cell/range refs with 'CELL'
    - Replaces numeric literals with 'NUM'
    - Removes whitespace
    - Uppercases function names/operators for consistency
    This is intentionally conservative and deterministic.
    """
    if not isinstance(formula, str):
        return ""
    f = formula.strip()
    if not f.startswith("="):
        return ""
    # keep leading '=' to avoid collisions with non-formulas
    f = _CELL_REF_RE.sub("CELL", f)
    f = _NUM_RE.sub("NUM", f)
    f = _WS_RE.sub("", f)
    return f.upper()


def detect_formula_drift(
    formula_cells_by_sheet: Dict[str, List[Tuple[str, str]]],
    cfg: Dict[str, Any] | None = None,
) -> List[Dict[str, Any]]:
    """
    Detect copied-formula structural drift using a row-wise dominant-shape heuristic.

    Returns a list[dict] with:
      - sheet, cell, formula
      - dominant_shape, cell_shape
      - row (1-based), drift_scope ("row")
      - note
    """
    cfg = cfg or {}
    enabled = bool(cfg.get("enabled", True))
    if not enabled:
        return []

    min_group = int(cfg.get("min_row_formula_count", 6) or 6)
    dominant_ratio = float(cfg.get("dominant_ratio", 0.60) or 0.60)
    max_findings = int(cfg.get("max_findings", 50) or 50)

    findings: List[Dict[str, Any]] = []

    # build row -> list of (cell, formula, shape)
    for sheet, cells in (formula_cells_by_sheet or {}).items():
        rows: Dict[int, List[Tuple[str, str, str]]] = defaultdict(list)

        for addr, formula in cells:
            # addr like "H10"
            m = re.match(r"^[A-Z]{1,3}(\d+)$", str(addr))
            if not m:
                continue
            r = int(m.group(1))
            shape = normalize_formula(formula)
            if not shape:
                continue
            rows[r].append((addr, formula, shape))

        for r, items in rows.items():
            if len(items) < min_group:
                continue

            shapes = [it[2] for it in items]
            counts = Counter(shapes)
            dominant_shape, dom_count = counts.most_common(1)[0]
            if dom_count / max(1, len(items)) < dominant_ratio:
                # no strong 'expected' shape
                continue

            # any non-dominant shapes in this row are drift candidates
            for addr, formula, shape in items:
                if shape == dominant_shape:
                    continue
                findings.append(
                    {
                        "sheet": sheet,
                        "cell": addr,
                        "formula": formula,
                        "dominant_shape": dominant_shape,
                        "cell_shape": shape,
                        "row": r,
                        "drift_scope": "row",
                        "note": "Formula shape differs from dominant pattern in the same row (possible copy/paste drift).",
                    }
                )
                if len(findings) >= max_findings:
                    return findings

    return findings
