"""
main.py — Tkinter GUI only.

Responsibilities:
    - Build and display the GUI
    - Validate that required files have been selected
    - Collect user inputs into a DashboardRequest
    - Show a save dialog to get the output path
    - Spawn the background thread that calls orchestrator.run()
    - Relay results / errors back to the GUI via root.after()

Zero business logic. Zero data processing. Zero Excel writing.
"""

import logging
import threading
from typing import Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry

from orchestrator import DashboardRequest, DashboardResult, run as run_dashboard

# Configure logging once here at the application entry point.
# All other modules call logging.getLogger(__name__) and inherit this config.
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(name)s - %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


def select_file(label_var, filetypes=None):
    if filetypes is None:
        filetypes = [
            ("Data Files", "*.csv *.xlsx *.xls"),
            ("CSV Files", "*.csv"),
            ("Excel Files", "*.xlsx *.xls"),
        ]
    path = filedialog.askopenfilename(filetypes=filetypes)
    if path:
        label_var.set(path)


def _run_in_thread(request, progress):
    try:
        log.info("Background thread started")
        result = run_dashboard(request)

        if result.success:
            msg = result.message
            if result.trend_summary:
                ts = result.trend_summary
                msg += (
                    "\n\nTrend vs previous report:"
                    f"\n  \u25b2 {ts['new_cve_count']:,} new CVE types   "
                    f"\u25bc {ts['resolved_cve_count']:,} resolved   "
                    f"\u23f3 {ts['persisting_cve_count']:,} persisting"
                )
            if result.warnings:
                msg += "\n\nWarnings:\n" + "\n".join(f"  - {w}" for w in result.warnings)

            _msg = msg  # capture value, not reference

            def _on_success():
                progress.stop()
                progress.destroy()
                generate_btn.config(state="normal")
                messagebox.showinfo("Done", _msg)

            root.after(0, _on_success)

        else:
            _err = result.message

            def _on_failure():
                progress.stop()
                progress.destroy()
                generate_btn.config(state="normal")
                messagebox.showerror("Error", f"Processing failed:\n{_err}")

            root.after(0, _on_failure)

    except Exception as exc:
        import traceback
        tb = traceback.format_exc()
        log.exception("Unexpected error in background thread")
        _exc_msg = f"Unexpected error:\n{exc}\n\n{tb}"

        def _on_exception():
            progress.stop()
            progress.destroy()
            generate_btn.config(state="normal")
            messagebox.showerror("Error", _exc_msg)

        root.after(0, _on_exception)


def process_reports():
    vuln_path        = vuln_var.get()
    rmm_path         = rmm_var.get()
    skip_rmm         = skip_rmm_var.get()
    include_patch    = include_patch_var.get()
    patch_path       = patch_var.get()
    include_trend    = include_trend_var.get()
    prev_report_path = prev_report_var.get()

    if not vuln_path:
        messagebox.showerror("Error", "Please select the Vulnerability Report.")
        return
    if not skip_rmm and not rmm_path:
        messagebox.showerror("Error", "Please select the Device Inventory / RMM Report.")
        return
    if include_patch and not patch_path:
        messagebox.showerror("Error",
            "Patch Report matching is enabled but no file selected.\n"
            "Please browse for a Patch Report or uncheck the option.")
        return
    if include_trend and not prev_report_path:
        messagebox.showerror("Error",
            "Trend tracking is enabled but no previous report selected.\n"
            "Please browse for a previous dashboard or uncheck the option.")
        return

    output_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")]
    )
    if not output_path:
        log.info("User cancelled save dialog")
        return

    cutoff_date = None if show_all_dates_var.get() else date_var.get()
    request = DashboardRequest(
        vuln_path              = vuln_path,
        output_path            = output_path,
        rmm_path               = rmm_path or None,
        skip_rmm               = skip_rmm,
        patch_path             = patch_path or None,
        include_patch          = include_patch,
        failure_report_path    = failure_var.get() or None,
        include_failure_report = include_failure_var.get(),
        prev_report_path       = prev_report_path or None,
        include_trend          = include_trend,
        threshold              = float(score_var.get()),
        cutoff_date            = cutoff_date,
        show_all_dates         = show_all_dates_var.get(),
    )

    log.info("Starting dashboard generation: %s", output_path)
    generate_btn.config(state="disabled")
    progress = ttk.Progressbar(root, mode="indeterminate")
    progress.pack(pady=5)
    progress.start()

    threading.Thread(target=_run_in_thread, args=(request, progress), daemon=True).start()


def toggle_rmm_state():
    state = tk.DISABLED if skip_rmm_var.get() else tk.NORMAL
    rmm_entry.config(state=state)
    rmm_browse_btn.config(state=state)

def toggle_date_state():
    date_entry.config(state=tk.DISABLED if show_all_dates_var.get() else tk.NORMAL)

def toggle_patch_state():
    state = tk.NORMAL if include_patch_var.get() else tk.DISABLED
    patch_entry.config(state=state)
    patch_browse_btn.config(state=state)

def toggle_trend_state():
    state = tk.NORMAL if include_trend_var.get() else tk.DISABLED
    prev_report_entry.config(state=state)
    prev_report_browse_btn.config(state=state)


root = tk.Tk()
root.title("N-able CVE Dashboard & Triage Tool")
root.geometry("570x800")
root.resizable(False, True)

tk.Label(root, text="N-able CVE Dashboard & Triage Tool",
         font=("Arial", 13, "bold")).pack(pady=(12, 4))

tk.Label(root, text="Vulnerability / CVE Report  (CSV or XLSX)",
         font=("Arial", 9, "bold")).pack(anchor="w", padx=14)
vuln_var = tk.StringVar()
vuln_entry = tk.Entry(root, textvariable=vuln_var, width=55, state="readonly")
vuln_entry.pack(padx=14)
tk.Button(root, text="Browse", command=lambda: select_file(vuln_var)).pack()

tk.Label(root, text="Device Inventory / RMM Report  (CSV or XLSX)",
         font=("Arial", 9, "bold")).pack(anchor="w", padx=14, pady=(8, 0))
rmm_var = tk.StringVar()
rmm_frame = tk.Frame(root)
rmm_frame.pack(fill="x", padx=14)
rmm_entry = tk.Entry(rmm_frame, textvariable=rmm_var, width=44, state="readonly")
rmm_entry.pack(side=tk.LEFT)
rmm_browse_btn = tk.Button(rmm_frame, text="Browse", command=lambda: select_file(rmm_var))
rmm_browse_btn.pack(side=tk.LEFT, padx=4)
skip_rmm_var = tk.BooleanVar()
tk.Checkbutton(root, text="Skip RMM (CVE export includes device info)",
               variable=skip_rmm_var, command=toggle_rmm_state).pack(anchor="w", padx=14)

score_frame = tk.Frame(root)
score_frame.pack(anchor="w", padx=14, pady=(8, 0))
tk.Label(score_frame, text="Minimum CVE Score:", font=("Arial", 9, "bold")).pack(side=tk.LEFT)
score_var = tk.StringVar(value="9.0")
tk.Entry(score_frame, textvariable=score_var, width=6).pack(side=tk.LEFT, padx=6)

date_frame = tk.Frame(root)
date_frame.pack(anchor="w", padx=14, pady=(6, 0))
tk.Label(date_frame, text="Filter: detections since", font=("Arial", 9, "bold")).pack(side=tk.LEFT)
date_var = tk.StringVar()
date_entry = DateEntry(date_frame, textvariable=date_var, date_pattern="yyyy-mm-dd", width=12)
date_entry.pack(side=tk.LEFT, padx=6)
show_all_dates_var = tk.BooleanVar()
tk.Checkbutton(date_frame, text="Show All Dates",
               variable=show_all_dates_var, command=toggle_date_state).pack(side=tk.LEFT)
toggle_date_state()

tk.Label(root, text="Patch Report  (optional, CSV or XLSX)",
         font=("Arial", 9, "bold")).pack(anchor="w", padx=14, pady=(10, 0))
patch_var = tk.StringVar()
patch_frame = tk.Frame(root)
patch_frame.pack(fill="x", padx=14)
patch_entry = tk.Entry(patch_frame, textvariable=patch_var, width=44, state="disabled")
patch_entry.pack(side=tk.LEFT)
patch_browse_btn = tk.Button(patch_frame, text="Browse",
                              command=lambda: select_file(patch_var), state="disabled")
patch_browse_btn.pack(side=tk.LEFT, padx=4)
include_patch_var = tk.BooleanVar()
tk.Checkbutton(root, text="Include Patch Report matching",
               variable=include_patch_var, command=toggle_patch_state).pack(anchor="w", padx=14)

# ── Patch Failure Report (optional) ───────────────────────────────────────────
tk.Label(root, text="Patch Failure Report  (optional, CSV)",
         font=("Arial", 9, "bold")).pack(anchor="w", padx=14, pady=(6, 0))
failure_var = tk.StringVar()
failure_frame = tk.Frame(root)
failure_frame.pack(fill="x", padx=14)
failure_entry = tk.Entry(failure_frame, textvariable=failure_var, width=44, state="disabled")
failure_entry.pack(side=tk.LEFT)
failure_browse_btn = tk.Button(failure_frame, text="Browse",
                               command=lambda: select_file(failure_var, [("CSV Files","*.csv")]),
                               state="disabled")
failure_browse_btn.pack(side=tk.LEFT, padx=4)
include_failure_var = tk.BooleanVar()
tk.Checkbutton(root, text="Include Patch Failure analysis",
               variable=include_failure_var,
               command=lambda: [
                   failure_entry.config(state=tk.NORMAL if include_failure_var.get() else tk.DISABLED),
                   failure_browse_btn.config(state=tk.NORMAL if include_failure_var.get() else tk.DISABLED),
               ]).pack(anchor="w", padx=14)

tk.Label(root, text="Previous Dashboard  (optional, for M-o-M trends)",
         font=("Arial", 9, "bold")).pack(anchor="w", padx=14, pady=(10, 0))
prev_report_var = tk.StringVar()
prev_frame = tk.Frame(root)
prev_frame.pack(fill="x", padx=14)
prev_report_entry = tk.Entry(prev_frame, textvariable=prev_report_var, width=44, state="disabled")
prev_report_entry.pack(side=tk.LEFT)
prev_report_browse_btn = tk.Button(
    prev_frame, text="Browse",
    command=lambda: select_file(prev_report_var, [("Excel Files", "*.xlsx")]),
    state="disabled",
)
prev_report_browse_btn.pack(side=tk.LEFT, padx=4)
include_trend_var = tk.BooleanVar()
tk.Checkbutton(root, text="Include month-over-month trend analysis",
               variable=include_trend_var, command=toggle_trend_state).pack(anchor="w", padx=14)

generate_btn = tk.Button(
    root,
    text="GENERATE COMPLETE DASHBOARD",
    command=process_reports,
    bg="#0078D7", fg="white",
    font=("Arial", 10, "bold"),
    height=2,
)
generate_btn.pack(pady=14)

root.mainloop()
