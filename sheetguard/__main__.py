# sheetguard/__main__.py
from __future__ import annotations

import sys

from sheetguard.cli.main import main as cli_main


def main() -> int:
    """
    Entry behavior:
      - python -m sheetguard <file> ...  -> CLI
      - python -m sheetguard            -> UI (if available)
    """
    if len(sys.argv) > 1:
        return cli_main()

    # No args: try UI. If UI isn't present, fall back to CLI help.
    try:
        from sheetguard.ui.launcher import launch_ui
        from sheetguard.cli.main import run_scan

        def _run_analysis(workbook_path: str, out_dir: str, top_n: int, trial: bool):
            return run_scan(workbook_path=workbook_path, out_dir=out_dir, top_n=top_n, trial=trial)

        return launch_ui(_run_analysis, default_top=10)
    except Exception:
        return cli_main([])


if __name__ == "__main__":
    raise SystemExit(main())
