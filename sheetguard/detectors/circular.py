from __future__ import annotations

from typing import Dict, List

import networkx as nx

from sheetguard.core.graph import detect_cycles


def circular_reference_findings(g: nx.DiGraph) -> Dict:
    """
    Deterministic circular reference findings using graph cycle detection.
    """
    cycles = detect_cycles(g)

    items: List[Dict] = []
    for cyc in cycles:
        if len(cyc) == 1:
            items.append(
                {
                    "cycle_type": "self_reference",
                    "cells": cyc,
                    "explanation": "Formula references itself directly (self-loop).",
                }
            )
        else:
            items.append(
                {
                    "cycle_type": "multi_cell_cycle",
                    "cells": cyc,
                    "explanation": "Cells form a dependency loop (multi-node cycle).",
                }
            )

    return {
        "rule_id": "CIRCULAR_REFERENCE",
        "severity": "HIGH" if cycles else "LOW",
        "cycle_count": len(cycles),
        "items": items,
    }
