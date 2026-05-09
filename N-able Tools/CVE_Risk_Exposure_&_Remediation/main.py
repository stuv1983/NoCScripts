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
from datetime import date, timedelta, datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from orchestrator import DashboardRequest, DashboardResult, run as run_dashboard

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

            _msg = msg 

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
        sync_baselines         = sync_baselines_var.get(),
        report_month           = report_month_var.get().strip(),
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

# Create the main window and all GUI components 
# (labels, entries, buttons, checkboxes) with appropriate layout and styling.
# Each "Browse" button calls select_file() with the corresponding StringVar to update the entry field.
# The "Generate" button calls process_reports() to validate inputs and start the background processing thread.
# Checkboxes toggle the state of related input fields (e.g. skipping RMM disables the RMM file input).
# The background thread runs _run_in_thread(), which calls the orchestrator and then uses root.after() to update the GUI with results or errors once processing is complete.
# No business logic or data processing should be done in this file - it should only handle the GUI and user interactions.
root = tk.Tk()
root.title("N-able CVE Dashboard & Triage Tool")
root.geometry("570x830")
# root.configure(bg="black")
# root.iconbitmap("icon.ico")
root.resizable(False, True)
# Set a consistent font for all widgets
default_font = ("Arial", 10)
# style = ttk.Style(root)
# style.configure("TLabel", font=default_font, background="black", foreground="white")
# style.configure("TButton", font=default_font, background="#0078D7", foreground="white")   
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
# By default, we expect RMM data to be included since it's needed for the most comprehensive analysis and remediation guidance.
# If the user checks "Skip RMM", we disable the RMM file input since it's not needed. This allows users who only have a CVE report to still use the tool with limited functionality, while encouraging those with RMM data to include it for the best results.
# Note: If "Skip RMM" is checked, the orchestrator should still be able to run and generate a dashboard, but it will not have device-specific insights or remediation steps that rely on RMM data. The CVE analysis will be based solely on the vulnerability report, and any sections of the dashboard that require RMM data should be hidden or show a message indicating that RMM data was not included.
# This option is intended for users who may have a CVE report from a source other than RMM, or who want to test the tool without providing RMM data. However, for the most accurate and actionable dashboard, including RMM data is recommended.
# The toggle_rmm_state() function will enable or disable the RMM file input fields based on the state of the "Skip RMM" checkbox. When "Skip RMM" is checked, the RMM entry and browse button will be disabled, and the background processing logic should be designed to handle the case where no RMM data is provided (e.g. by skipping any steps that require RMM data and adjusting the dashboard output accordingly).
# This design allows for flexibility in how the tool can be used, while still guiding users towards providing the most comprehensive data for the best results.
# The default state is to include RMM data, so the RMM file input is enabled by default. If the user decides to skip RMM, they can check the box and the input will be disabled to reflect that it's not needed.
# The orchestrator should be designed to handle both cases (with or without RMM data) gracefully, ensuring that the tool remains functional and provides useful insights even if RMM data is not included, while still encouraging users to include it for the best possible analysis and remediation guidance.
# The "Skip RMM" option is essentially a way to allow users to use the tool with just a CVE report, while making it clear that including RMM data will provide a richer and more actionable dashboard. The GUI should reflect this by enabling/disabling the RMM file input based on the state of the checkbox, and the background processing logic should be robust enough to handle either scenario without errors.
# This approach provides flexibility for different user needs and data availability, while still promoting the inclusion of RMM data for the best results.
# The toggle_rmm_state() function is a simple way to manage the state of the RMM file input fields based on the user's choice to include or skip RMM data. By default, the tool expects RMM data to be included for the most comprehensive analysis, but it can still function with just a CVE report if the user chooses to skip RMM. The GUI should make it clear that including RMM data is recommended for the best results, while still allowing users to proceed without it if necessary.
score_frame = tk.Frame(root)
score_frame.pack(anchor="w", padx=14, pady=(8, 0))
tk.Label(score_frame, text="Minimum CVE Score:", font=("Arial", 9, "bold")).pack(side=tk.LEFT)
score_var = tk.StringVar(value="9.0")
tk.Entry(score_frame, textvariable=score_var, width=6).pack(side=tk.LEFT, padx=6)

date_frame = tk.Frame(root)
date_frame.pack(anchor="w", padx=14, pady=(6, 0))
tk.Label(date_frame, text="Exclude stale devices last seen before", font=("Arial", 9, "bold")).pack(side=tk.LEFT)
date_var = tk.StringVar()
_default_since = (date.today() - timedelta(days=90)).strftime('%d/%m/%Y')
date_var.set(_default_since)
date_entry = tk.Entry(date_frame, textvariable=date_var, width=12)
date_entry.pack(side=tk.LEFT, padx=6)
tk.Label(date_frame, text="(dd/mm/yyyy)").pack(side=tk.LEFT, padx=4)
show_all_dates_var = tk.BooleanVar()
tk.Checkbutton(date_frame, text="Show All Dates",
               variable=show_all_dates_var, command=toggle_date_state).pack(side=tk.LEFT)
toggle_date_state()

# New GUI Field for Report Month
month_frame = tk.Frame(root)
month_frame.pack(anchor="w", padx=14, pady=(6, 0))
tk.Label(month_frame, text="Report Month:", font=("Arial", 9, "bold")).pack(side=tk.LEFT)
report_month_var = tk.StringVar(value=datetime.now().strftime('%B %Y'))
tk.Entry(month_frame, textvariable=report_month_var, width=15).pack(side=tk.LEFT, padx=6)


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

sync_baselines_var = tk.BooleanVar()
tk.Checkbutton(root, text="Refresh product baselines before run",
               variable=sync_baselines_var).pack(anchor="w", padx=14, pady=(6, 0))

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