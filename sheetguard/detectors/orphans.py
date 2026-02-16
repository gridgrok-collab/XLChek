from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Set, Tuple

import networkx as nx

try:
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
except Exception:  # pragma: no cover
    coordinate_from_string = None
    column_index_from_string = None


def _split_cell_ref(ref: str) -> Tuple[str, str]:
    if not ref:
        return "", ""
    if "!" in ref:
        sh, addr = ref.split("!", 1)
        return sh, addr
    return "", ref


def _parse_a1(addr: str) -> Tuple[int, int]:
    if not addr or coordinate_from_string is None or column_index_from_string is None:
        return 0, 0
    try:
        col_letters, row_num = coordinate_from_string(addr)
        col_num = column_index_from_string(col_letters)
        return int(row_num), int(col_num)
    except Exception:
        return 0, 0


@dataclass(frozen=True)
class OrphanFinding:
    sheet: str
    cell: str
    formula: str
    reason: str = "formula not referenced by any other formula cell"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "type": "ORPHAN_FORMULA",
            "sheet": self.sheet,
            "cell": self.cell,
            "formula": self.formula,
            "reason": self.reason,
        }


def detect_orphan_formulas(
    g: nx.DiGraph,
    formula_nodes: List[str],
    *,
    cfg: Optional[Dict[str, Any]] = None,
    formula_by_node: Optional[Dict[str, str]] = None,
    sheet_max_rc: Optional[Dict[str, Tuple[int, int]]] = None,
    named_cells: Optional[Set[str]] = None,
) -> List[OrphanFinding]:
    cfg = cfg or {}
    if not bool(cfg.get("enabled", True)):
        return []

    exclude_sheets = set(str(s) for s in (cfg.get("exclude_sheets") or []))
    exclude_named = bool(cfg.get("exclude_named_ranges", True))

    # Boundary exclusion flags (support both legacy + new keys)
    exclude_last_row_or_col = bool(cfg.get("exclude_last_row_or_col", False))
    exclude_last_row_col_legacy = bool(cfg.get("exclude_last_row_col", False))
    exclude_last_row = bool(cfg.get("exclude_last_row", False))
    exclude_last_column = bool(cfg.get("exclude_last_column", False))

    # Conservative interpretation:
    exclude_boundary_any = (
        exclude_last_row_or_col
        or exclude_last_row_col_legacy
        or (exclude_last_row and exclude_last_column)
    )

    max_findings = int(cfg.get("max_findings", 200) or 200)

    formula_by_node = formula_by_node or {}
    sheet_max_rc = sheet_max_rc or {}
    named_cells = named_cells or set()

    formula_set = set(formula_nodes)
    out: List[OrphanFinding] = []

    for node in formula_nodes:
        if len(out) >= max_findings:
            break

        sh, addr = _split_cell_ref(node)
        if not sh:
            continue

        if sh in exclude_sheets:
            continue

        if exclude_named and node in named_cells:
            continue

        # Boundary exclusions
        if addr:
            row, col = _parse_a1(addr)
            if row and col:
                mr, mc = sheet_max_rc.get(sh, (0, 0))

                if exclude_boundary_any and mr and mc and (row == mr or col == mc):
                    continue

                if exclude_last_row and mr and row == mr:
                    continue

                if exclude_last_column and mc and col == mc:
                    continue

        # referenced by another formula?
        used = False
        try:
            for pred in g.predecessors(node):
                if pred in formula_set:
                    used = True
                    break
        except Exception:
            used = False

        if used:
            continue

        f = str(formula_by_node.get(node, "") or "")
        out.append(OrphanFinding(sheet=sh, cell=addr, formula=f))

    return out
