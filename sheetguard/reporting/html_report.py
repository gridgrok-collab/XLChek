# sheetguard/reporting/html_report.py
from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List


def _default_css() -> str:
    """Load embedded CSS for the HTML report.

    Prefers templates/styles.css (kept under version control).
    Falls back to minimal CSS if missing.
    """
    css_path = Path(__file__).resolve().parent / "templates" / "styles.css"
    try:
        return css_path.read_text(encoding="utf-8")
    except Exception:
        return "body{font-family:system-ui,Segoe UI,Roboto,Arial;margin:16px;} .card{border:1px solid #ddd;padding:12px;border-radius:10px;}"


def _risk_badge(risk: str) -> str:
    r = (risk or "UNKNOWN").upper()
    cls = "med"
    if r == "LOW":
        cls = "low"
    elif r == "HIGH":
        cls = "high"
    elif r == "UNKNOWN":
        cls = "med"
    return f'<span class="pill {cls}">{r}</span>'


def _table_rows(items: List[Dict[str, Any]], columns: List[str]) -> str:
    if not items:
        return '<tr><td colspan="99" class="muted">No data</td></tr>'
    out = []
    for it in items:
        tds = []
        for c in columns:
            v = it.get(c, "")
            tds.append(f"<td>{v}</td>")
        out.append("<tr>" + "".join(tds) + "</tr>")
    return "".join(out)


def _render_sheet_table(sheets: List[Dict[str, Any]]) -> str:
    if not sheets:
        return '<tr><td colspan="6">No sheets</td></tr>'

    rows = []
    for s in sheets:
        rows.append(
            "<tr>"
            f"<td>{s.get('sheet','')}</td>"
            f"<td>{_risk_badge(s.get('risk','UNKNOWN'))}</td>"
            f"<td>{int(s.get('formulas',0) or 0)}</td>"
            f"<td>{int(s.get('circular_cells',0) or 0)}</td>"
            f"<td>{int(s.get('high_risk_cells',0) or 0)}</td>"
            f"<td>{s.get('reason','')}</td>"
            "</tr>"
        )
    return "".join(rows)


def _render_top_risks(top: List[Dict[str, Any]]) -> str:
    if not top:
        return '<tr><td colspan="3">No high-risk cells found.</td></tr>'
    rows = []
    for r in top:
        rows.append(
            "<tr>"
            f"<td>{r.get('cell','')}</td>"
            f"<td>{int(r.get('downstream_reach',0) or 0)}</td>"
            f"<td>{int((r.get('fan_out', None) if r.get('fan_out', None) is not None else r.get('direct_fan_out',0)) or 0)}</td>"
            "</tr>"
        )
    return "".join(rows)


def _render_cycles(cycles: List[Dict[str, Any]]) -> str:
    if not cycles:
        return '<tr><td colspan="2">No circular references found.</td></tr>'
    rows = []
    for c in cycles:
        ctype = c.get("type", "cycle")
        cells = ", ".join(c.get("cells", []) or [])
        rows.append(f"<tr><td>{ctype}</td><td>{cells}</td></tr>")
    return "".join(rows)


def _render_hardcoded_constants(items: List[Dict[str, Any]]) -> str:
    if not items:
        return '<tr><td colspan="5">No embedded constants detected (or detector disabled).</td></tr>'
    rows = []
    for it in items:
        sheet = it.get("sheet", "")
        cell = it.get("cell", "")
        lit = it.get("literal", "")
        op = it.get("operator", "")
        func = it.get("function", "")
        fml = it.get("formula", "")
        # Compact formula display for readability
        if isinstance(fml, str) and len(fml) > 120:
            fml = fml[:117] + "..."
        where = f"{sheet}!{cell}" if sheet and cell else (cell or "")
        ctx = func if func else "-"
        rows.append(
            "<tr>"
            f"<td>{where}</td>"
            f"<td>{lit}</td>"
            f"<td>{op or '-'}" + "</td>"
            f"<td>{ctx}</td>"
            f"<td style=\"font-family:ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;\">{fml}</td>"
            "</tr>"
        )
    return "".join(rows)


def _render_orphans(items: List[Dict[str, Any]]) -> str:
    if not items:
        return '<tr><td colspan="3">No orphan formulas detected (or all excluded by config).</td></tr>'
    rows = []
    for it in items:
        sheet = it.get("sheet", "")
        cell = it.get("cell", "")
        fml = it.get("formula", "") or ""
        reason = it.get("reason", "") or ""
        if isinstance(fml, str) and len(fml) > 180:
            fml = fml[:177] + "..."
        where = f"{sheet}!{cell}" if sheet and cell else (cell or "")
        rows.append(
            "<tr>"
            f"<td>{where}</td>"
            f"<td style=\"font-family:ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;\">{fml}</td>"
            f"<td>{reason}</td>"
            "</tr>"
        )
    return "".join(rows)




def _render_drift(items: List[Dict[str, Any]]) -> str:
    if not items:
        return '<tr><td colspan="4">No formula drift detected (or detector disabled).</td></tr>'
    rows = []
    for it in items:
        sheet = it.get("sheet", "")
        cell = it.get("cell", "")
        row = it.get("row", "")
        fml = it.get("formula", "") or ""
        note = it.get("note", "") or ""
        if isinstance(fml, str) and len(fml) > 180:
            fml = fml[:177] + "..."
        where = f"{sheet}!{cell}" if sheet and cell else (cell or "")
        rows.append(
            "<tr>"
            f"<td>{where}</td>"
            f"<td>{row}</td>"
            f"<td style=\"font-family:ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;\">{fml}</td>"
            f"<td>{note}</td>"
            "</tr>"
        )
    return "".join(rows)

def _render_trial_banner(trial: bool) -> str:
    if not trial:
        return ""
    return '<div class="badge" style="margin-top:10px;">TRIAL MODE</div>'


def render_html_report(scan: Dict[str, Any], **_ignored_kwargs: Any) -> str:
    """
    Compatible with callers that previously tried:
      render_html_report(scan=..., errors=...)
    We accept **kwargs and ignore unknown keys.
    """
    template_path = Path(__file__).resolve().parent / "templates" / "report.html"
    template = template_path.read_text(encoding="utf-8")

    signals = scan.get("signals", {}) or {}
    sheets = scan.get("sheets", []) or []
    top = scan.get("top_risks", []) or []
    cycles = scan.get("cycles", []) or []
    constants = scan.get("hardcoded_constants", []) or []
    orphans = scan.get("orphan_formulas", []) or []

    # Fill placeholders used by report.html
    html = template
    html = html.replace("$CSS", _default_css())

    # IMPORTANT: Replace longer placeholders *before* shorter prefixes.
    # Otherwise a placeholder like "$WORKBOOK_BADGE" gets partially replaced
    # by the "$WORKBOOK" replacement, producing "<workbook>_BADGE" in the HTML.
    wb_risk = str(scan.get("workbook_risk", "UNKNOWN"))
    html = html.replace("$WORKBOOK_BADGE", _risk_badge(wb_risk))

    html = html.replace("$WORKBOOK", str(scan.get("workbook", "")))
    html = html.replace("$GENERATED", str(scan.get("generated", "")))
    html = html.replace("$TRIAL_BANNER", _render_trial_banner(bool(scan.get("trial", False))))

    html = html.replace("$REASON", str(scan.get("reason", "")))

    html = html.replace("$CIRCULAR", str(int(signals.get("circular", 0) or 0)))
    html = html.replace("$VOLATILE", str(int(signals.get("volatile", 0) or 0)))
    html = html.replace("$SYMBOLIC", str(int(signals.get("symbolic", 0) or 0)))
    html = html.replace("$TOP_RISKS", str(int(signals.get("top_risks", 0) or 0)))
    html = html.replace("$HARDCODED", str(int(signals.get("hardcoded_constants", 0) or 0)))
    html = html.replace("$ORPHANS", str(int(signals.get("orphan_formulas", 0) or 0)))
    html = html.replace("$DRIFT_FORMULAS", str(int(signals.get("drift_formulas", 0) or 0)))

    # These are also present in some versions of the template
    html = html.replace("$CIRCULAR_CYCLES", str(int(signals.get("circular", 0) or 0)))
    html = html.replace("$VOLATILE_HITS", str(int(signals.get("volatile", 0) or 0)))
    html = html.replace("$SYMBOLIC_REFS", str(int(signals.get("symbolic", 0) or 0)))
    html = html.replace("$HARDCODED_CONSTANTS", str(int(signals.get("hardcoded_constants", 0) or 0)))
    html = html.replace("$ORPHAN_FORMULAS", str(int(signals.get("orphan_formulas", 0) or 0)))
    html = html.replace("$DRIFT_FORMULAS", str(int(signals.get("drift_formulas", 0) or 0)))

    html = html.replace("$TABLE_SHEETS", _render_sheet_table(sheets))
    html = html.replace("$TABLE_TOP_RISKS", _render_top_risks(top))
    html = html.replace("$TABLE_CIRCULAR", _render_cycles(cycles))
    html = html.replace("$TABLE_HARDCODED_CONSTANTS", _render_hardcoded_constants(constants))
    html = html.replace("$TABLE_ORPHANS", _render_orphans(orphans))
    html = html.replace("$TABLE_DRIFT", _render_drift((scan.get("formula_drift", []) or [])))

    # Optional errors block if template contains it
    if "${ERRORS_BLOCK}" in html or "$ERRORS_BLOCK" in html:
        # We are NOT changing report.html; if placeholder exists, fill it, else do nothing.
        errors = scan.get("errors", []) or []
        if errors:
            # simple minimal rendering
            rows = []
            for e in errors:
                rows.append(
                    "<tr>"
                    f"<td>{e.get('scope','')}</td>"
                    f"<td>{e.get('type','')}</td>"
                    f"<td>{e.get('dependent','')}</td>"
                    f"<td>{e.get('ref','')}</td>"
                    f"<td>{e.get('details','')}</td>"
                    "</tr>"
                )
            block = (
                '<div class="section"><h2>Errors &amp; Warnings</h2>'
                '<div class="card"><table class="table">'
                "<thead><tr><th>Scope</th><th>Type</th><th>Dependent</th><th>Ref</th><th>Details</th></tr></thead>"
                f"<tbody>{''.join(rows)}</tbody></table></div></div>"
            )
        else:
            block = ""
        html = html.replace("${ERRORS_BLOCK}", block)
        html = html.replace("$ERRORS_BLOCK", block)

    return html
