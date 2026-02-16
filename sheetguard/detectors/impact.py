from __future__ import annotations

from collections import deque
from typing import Dict, List, Tuple

import networkx as nx


def downstream_reach(g: nx.DiGraph, start: str, max_nodes: int = 200000) -> int:
    """
    Count how many nodes are reachable downstream from start (excluding start).
    Bounded BFS to protect runtime on large graphs.
    """
    visited = set()
    q = deque()

    for nxt in g.successors(start):
        visited.add(nxt)
        q.append(nxt)

    while q and len(visited) < max_nodes:
        node = q.popleft()
        for nxt in g.successors(node):
            if nxt not in visited:
                visited.add(nxt)
                q.append(nxt)

    return len(visited)


def rank_high_risk_cells(g: nx.DiGraph, formula_nodes: List[str], top_n: int = 10) -> List[Dict]:
    """
    Returns top N risky nodes based on downstream reach, then direct fan-out.
    """
    scored = []
    for n in formula_nodes:
        reach = downstream_reach(g, n)
        fan_out = g.out_degree(n)  # direct dependents
        scored.append((reach, fan_out, n))

    scored.sort(reverse=True)  # highest reach first
    out = []
    for reach, fan_out, n in scored[:top_n]:
        out.append(
            {
                "cell": n,
                "downstream_reach": int(reach),
                "direct_fan_out": int(fan_out),
                "explanation": "High impact radius: many downstream dependents.",
            }
        )
    return out
