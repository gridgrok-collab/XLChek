from __future__ import annotations

from typing import Dict, List, Set


def _flatten_circular_cells(findings: List[Dict]) -> Set[str]:
    """
    Extract all cell node IDs involved in circular references.
    Expected format: findings contains an item with rule_id == "CIRCULAR_REFERENCE"
    and an "items" list, each item has "cells": [node_ids...]
    """
    circ = next((f for f in findings if f.get("rule_id") == "CIRCULAR_REFERENCE"), None)
    if not circ:
        return set()

    out: Set[str] = set()
    for item in circ.get("items", []):
        for c in item.get("cells", []):
            out.add(c)
    return out


def compute_executive_summary(results: Dict) -> Dict:
    """
    Deterministic v1 summary logic.

    Inputs (from results dict):
      - findings (circular)
      - total_volatile_hits
      - graph.symbolic_refs
      - top_risks
      - sheets[*].formula_cells
    """
    circular_cells = _flatten_circular_cells(results.get("findings", []))
    circular_count = len(circular_cells)

    volatile_hits = int(results.get("total_volatile_hits", 0) or 0)
    symbolic_refs = int(results.get("graph", {}).get("symbolic_refs", 0) or 0)
    top_risks = results.get("top_risks", []) or []
    high_impact_count = len(top_risks)

    # --- Workbook risk rules (as agreed) ---
    workbook_risk = "LOW"
    reasons: List[str] = []

    if circular_count >= 1:
        workbook_risk = "HIGH"
        reasons.append(f"{circular_count} circular-reference cell(s) detected")

    if high_impact_count >= 5:
        workbook_risk = "HIGH"
        reasons.append(f"{high_impact_count} high-impact cells in Top-N list")

    if workbook_risk != "HIGH":
        if volatile_hits > 0:
            workbook_risk = "MEDIUM"
            reasons.append(f"{volatile_hits} volatile function hit(s)")
        if symbolic_refs > 0:
            workbook_risk = "MEDIUM"
            reasons.append(f"{symbolic_refs} symbolic reference(s) (whole-row/whole-column/external)")

    if not reasons:
        reasons.append("No high-risk patterns detected under v1 rules")

    # --- Per-sheet diagnostics ---
    sheet_risks: List[Dict] = []
    for s in results.get("sheets", []):
        sheet = s.get("sheet", "")
        formula_cells = int(s.get("formula_cell_count", 0) or 0)

        # circular involvement: count circular cells belonging to this sheet
        prefix = f"{sheet}!"
        sheet_circular = sum(1 for c in circular_cells if c.startswith(prefix))

        # high-risk cells on this sheet (from top_risks list)
        sheet_high_risk = sum(1 for r in top_risks if str(r.get("cell", "")).startswith(prefix))

        # Sheet risk rules (as agreed)
        if sheet_circular >= 1:
            risk = "HIGH"
            sheet_reason = f"{sheet_circular} circular-reference cell(s) on this sheet"
        elif sheet_high_risk >= 3:
            risk = "MEDIUM"
            sheet_reason = f"{sheet_high_risk} high-impact cells on this sheet"
        else:
            risk = "LOW"
            sheet_reason = "No high-risk patterns detected under v1 sheet rules"

        sheet_risks.append(
            {
                "sheet": sheet,
                "risk": risk,
                "reason": sheet_reason,
                "formula_cells": formula_cells,
                "circular_cells": int(sheet_circular),
                "high_risk_cells": int(sheet_high_risk),
            }
        )

    return {
        "workbook_risk": workbook_risk,
        "reason": "; ".join(reasons),
        "sheet_risks": sheet_risks,
        "counters": {
            "circular_cells": circular_count,
            "volatile_hits": volatile_hits,
            "symbolic_refs": symbolic_refs,
            "top_risks_count": high_impact_count,
        },
    }
