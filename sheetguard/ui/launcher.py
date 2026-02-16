from __future__ import annotations

import os
import threading
import webbrowser
from pathlib import Path

from tkinter import StringVar, filedialog, messagebox
from tkinter.ttk import Button, Frame, Label, Progressbar, Separator

try:
    # Drag & drop support
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:  # pragma: no cover
    TkinterDnD = None
    DND_FILES = None


def _is_excel(path: str) -> bool:
    p = path.lower()
    return p.endswith(".xlsx") or p.endswith(".xlsm")


def _normalize_drop_path(data: str) -> str:
    r"""
    tkinterdnd2 gives paths like:
    - {C:\path with spaces\file.xlsx}
    - C:\path\file.xlsx
    """
    s = data.strip()
    if s.startswith("{") and s.endswith("}"):
        s = s[1:-1]
    return s


def launch_ui(run_analysis_fn, default_top: int = 10) -> int:
    """
    run_analysis_fn: callable like run_analysis_fn(workbook_path, out_dir, top_n, trial)
    It should return (json_path, html_path).
    """

    if TkinterDnD is None:
        # Fallback: no drag/drop, only browse.
        from tkinter import Tk  # type: ignore
        root = Tk()
    else:
        root = TkinterDnD.Tk()

    root.title("XLChek")
    root.geometry("640x360")
    root.minsize(640, 360)

    status = StringVar(value="Drop an Excel file here, or click Browse…")
    selected = StringVar(value="")
    progress = StringVar(value="")

    def set_busy(is_busy: bool) -> None:
        btn_browse.config(state=("disabled" if is_busy else "normal"))
        btn_run.config(state=("disabled" if is_busy else "normal"))
        if is_busy:
            bar.start(10)
        else:
            bar.stop()

    def pick_file() -> None:
        p = filedialog.askopenfilename(
            title="Select an Excel file",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if p:
            selected.set(p)
            status.set(f"Selected: {p}")

    def run_now(path: str) -> None:
        path = path.strip().strip('"')
        if not path:
            messagebox.showwarning("XLChek", "Please select an Excel file.")
            return
        if not os.path.exists(path):
            messagebox.showerror("XLChek", f"File not found:\n{path}")
            return
        if not _is_excel(path):
            messagebox.showerror("XLChek", "Please select a .xlsx or .xlsm file.")
            return

        wb_path = Path(path)
        out_dir = str(wb_path.parent)

        set_busy(True)
        progress.set("Running analysis… this may take a moment.")
        status.set(f"Analyzing: {path}")

        def worker() -> None:
            try:
                json_path, html_path = run_analysis_fn(
                    workbook_path=str(wb_path),
                    out_dir=out_dir,
                    top_n=default_top,
                    trial=False,
                )
                progress.set("Done. Opening HTML report…")
                if html_path and os.path.exists(html_path):
                    # UI opens exactly once
                    webbrowser.open_new_tab(Path(html_path).as_uri())
                status.set("Completed successfully.")
            except Exception as e:
                messagebox.showerror("XLChek", f"Analysis failed:\n\n{e}")
                status.set("Failed. See error message.")
            finally:
                set_busy(False)
                progress.set("")

        threading.Thread(target=worker, daemon=True).start()

    def on_run_clicked() -> None:
        run_now(selected.get())

    # ----- Layout -----
    outer = Frame(root, padding=18)
    outer.pack(fill="both", expand=True)

    Label(outer, text="XLChek", font=("Segoe UI", 22, "bold")).pack(anchor="w")
    Label(
        outer,
        text="Offline Excel Diagnostic Report (read-only)",
        foreground="#6b7280",
        font=("Segoe UI", 10),
    ).pack(anchor="w", pady=(2, 14))

    drop = Frame(outer, padding=16)
    drop.pack(fill="both", expand=True)

    drop_box = Frame(drop, padding=18, style="Card.TFrame")
    drop_box.pack(fill="both", expand=True)

    Label(drop_box, text="Drop an Excel file here", font=("Segoe UI", 14, "bold")).pack(anchor="center", pady=(18, 6))
    Label(
        drop_box,
        text="(.xlsx / .xlsm)  •  Report opens automatically",
        foreground="#6b7280",
        font=("Segoe UI", 10),
    ).pack(anchor="center")

    Separator(drop_box).pack(fill="x", pady=18)

    Label(drop_box, textvariable=status, font=("Segoe UI", 10)).pack(anchor="w")
    Label(drop_box, textvariable=progress, foreground="#2563eb", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(6, 0))

    bar = Progressbar(drop_box, mode="indeterminate")
    bar.pack(fill="x", pady=(10, 0))

    actions = Frame(outer)
    actions.pack(fill="x", pady=(12, 0))

    btn_browse = Button(actions, text="Browse…", command=pick_file)
    btn_browse.pack(side="left")

    btn_run = Button(actions, text="Run Analysis", command=on_run_clicked)
    btn_run.pack(side="left", padx=(10, 0))

    Button(actions, text="Close", command=root.destroy).pack(side="right")

    # Drag & Drop hookup
    if TkinterDnD is not None and DND_FILES is not None:
        def on_drop(event):
            p = _normalize_drop_path(event.data)
            selected.set(p)
            status.set(f"Selected: {p}")
            run_now(p)

        drop_box.drop_target_register(DND_FILES)
        drop_box.dnd_bind("<<Drop>>", on_drop)

    set_busy(False)
    root.mainloop()
    return 0
