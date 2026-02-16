"""
sheetguard/core/graph.py

Builds a dependency graph from scan output (produced in cli.main).
We do NOT evaluate formulas. We only model relationships between formula cells
and the references they mention.

Public API expected by other modules:
  - build_dependency_graph(scan: dict) -> dict with keys:
      graph_obj: nx.DiGraph
      formula_nodes: list[str]
      errors: list[dict]
      stats: dict(nodes=int, edges=int, symbolic_refs=int)
  - downstream_reach(g, node) -> int
  - fan_out(g, node) -> int
  - detect_cycles(g, limit=25) -> list[dict]  (each dict has type, cells)
"""

from __future__ import annotations

from typing import Any, Dict, List

import networkx as nx


def _node_id(sheet: str, addr: str) -> str:
    return f"{sheet}!{addr}"


def build_dependency_graph(*args, **kwargs):
    """
    Backward-compatible dependency graph builder.

    Supported call styles:
      1) build_dependency_graph(scan: dict) -> dict   (newer style)
      2) build_dependency_graph(sheet_name=..., cell_addr=..., refs=..., g=..., symbolic_edges=...) -> None (Jan-28 v1 style)
    """
    # Style (1): single scan dict
    if args and len(args) == 1 and isinstance(args[0], dict) and not kwargs:
        return _build_dependency_graph_from_scan(args[0])

    # Style (2): incremental edge add
    sheet_name = kwargs['sheet_name'] if 'sheet_name' in kwargs else (args[0] if len(args) > 0 else None)
    cell_addr = kwargs['cell_addr'] if 'cell_addr' in kwargs else (args[1] if len(args) > 1 else None)
    refs = kwargs['refs'] if 'refs' in kwargs else (args[2] if len(args) > 2 else None)
    g = kwargs['g'] if 'g' in kwargs else (args[3] if len(args) > 3 else None)
    symbolic_edges = kwargs['symbolic_edges'] if 'symbolic_edges' in kwargs else (kwargs.get('symbolic') if isinstance(kwargs.get('symbolic'), list) else None)
    if sheet_name is None or cell_addr is None or g is None or refs is None:
        raise TypeError('build_dependency_graph: missing required args (sheet_name, cell_addr, refs, g)')

    dep = _node_id(str(sheet_name), str(cell_addr))
    if not g.has_node(dep):
        g.add_node(dep)

    # Add edges dep -> ref (dependency direction: formula cell depends on referenced cell/range)
    for r in refs:
        # Accept RefToken-like objects or dicts
        sheet = getattr(r, 'sheet', None)
        raw = getattr(r, 'raw', None)
        kind = getattr(r, 'kind', None)
        symbolic = getattr(r, 'is_symbolic', None)
        if raw is None and isinstance(r, dict):
            raw = r.get('raw') or r.get('ref')
            sheet = r.get('sheet')
            kind = r.get('kind')
            symbolic = bool(r.get('symbolic') or r.get('is_symbolic') or False)
        raw = str(raw or '').strip()
        if not raw:
            continue

        # Symbolic refs (named ranges / structured refs) are tracked separately; we still model them as nodes for visibility
        if symbolic is None:
            symbolic = False

        ref_sheet = str(sheet) if sheet else str(sheet_name)
        ref_node = _node_id(ref_sheet, raw)
        g.add_edge(dep, ref_node)

        if bool(symbolic) and isinstance(symbolic_edges, list):
            symbolic_edges.append({'dependent': dep, 'ref': ref_node, 'raw': raw, 'sheet': ref_sheet, 'kind': str(kind or 'symbolic')})

    return None

def _build_dependency_graph_from_scan(scan: Dict[str, Any]) -> Dict[str, Any]:
    g = nx.DiGraph()
    errors: List[Dict[str, Any]] = []

    formula_nodes: List[str] = []
    symbolic_ref_count = 0

    sheets = scan.get("sheets") or []
    for sh in sheets:
        sheet_name = sh.get("name") or "Sheet"
        for cell in (sh.get("formula_cells") or []):
            cell_addr = cell.get("cell")
            if not cell_addr:
                continue

            src = _node_id(sheet_name, cell_addr)
            g.add_node(src, sheet=sheet_name, addr=cell_addr, is_formula=True)
            formula_nodes.append(src)

            refs = cell.get("refs") or []
            for r in refs:
                try:
                    # tolerate unexpected shapes
                    if not isinstance(r, dict):
                        r = {
                            "raw": getattr(r, "raw", str(r)),
                            "sheet": getattr(r, "sheet", None),
                            "symbolic": getattr(r, "symbolic", False),
                            "is_symbolic": getattr(r, "is_symbolic", False),
                        }

                    raw = str(r.get("raw") or "")
                    if not raw:
                        continue

                    ref_sheet = r.get("sheet") or sheet_name
                    is_symbolic = bool(r.get("symbolic") or r.get("is_symbolic") or False)

                    if is_symbolic:
                        symbolic_ref_count += 1
                        ref_node = f"{ref_sheet}!{raw}"
                        g.add_node(ref_node, sheet=ref_sheet, addr=raw, is_formula=False, is_symbolic=True)
                    else:
                        ref_node = _node_id(ref_sheet, raw)
                        g.add_node(ref_node, sheet=ref_sheet, addr=raw, is_formula=False, is_symbolic=False)

                    g.add_edge(src, ref_node)

                except Exception as e:
                    errors.append({
                        "scope": "graph",
                        "type": type(e).__name__,
                        "dependent": src,
                        "ref": (r.get("raw") if isinstance(r, dict) else str(r)),
                        "details": f"{type(e).__name__}: {e}",
                    })

    return {
        "graph_obj": g,
        "formula_nodes": formula_nodes,
        "errors": errors,
        "stats": {"nodes": g.number_of_nodes(), "edges": g.number_of_edges(), "symbolic_refs": symbolic_ref_count},
    }


def downstream_reach(g: nx.DiGraph, node: str) -> int:
    """Number of nodes reachable from node via directed edges."""
    try:
        return len(nx.descendants(g, node))
    except Exception:
        return 0


def fan_out(g: nx.DiGraph, node: str) -> int:
    """Out-degree (direct references)."""
    try:
        return int(g.out_degree(node))
    except Exception:
        return 0


def detect_cycles(g: nx.DiGraph, limit: int = 25) -> List[Dict[str, Any]]:
    """
    Return up to `limit` cycles. Each cycle is:
      { "type": "cycle", "cells": ["Sheet!A1", ...] }
    """
    out: List[Dict[str, Any]] = []
    try:
        for cyc in nx.simple_cycles(g):
            out.append({"type": "cycle", "cells": [str(x) for x in cyc]})
            if len(out) >= limit:
                break
    except Exception as e:
        out.append({"type": "error", "cells": [], "details": f"{type(e).__name__}: {e}"})
    return out