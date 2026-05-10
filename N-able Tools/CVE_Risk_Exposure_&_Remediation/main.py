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
import subprocess
import sys
import threading
from pathlib import Path
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

_CVE_REPO_DEFAULT = r"C:\NoCScripts\N-able Tools\CVE_Risk_Exposure_&_Remediation\cvelistV5"


# ===========================================================================
# FILE HELPER
# ===========================================================================

def select_file(label_var, filetypes=None):
    if filetypes is None:
        filetypes = [
            ("Data Files",  "*.csv *.xlsx *.xls"),
            ("CSV Files",   "*.csv"),
            ("Excel Files", "*.xlsx *.xls"),
        ]
    path = filedialog.askopenfilename(filetypes=filetypes)
    if path:
        label_var.set(path)


# ===========================================================================
# BACKGROUND WORKER
# ===========================================================================

def _run_in_thread(request, progress_bar):
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
                progress_bar.stop()
                progress_bar.grid_remove()
                generate_btn.config(state="normal")
                messagebox.showinfo("Done", _msg)
            root.after(0, _on_success)

        else:
            _err = result.message

            def _on_failure():
                progress_bar.stop()
                progress_bar.grid_remove()
                generate_btn.config(state="normal")
                messagebox.showerror("Error", f"Processing failed:\n{_err}")
            root.after(0, _on_failure)

    except Exception as exc:
        import traceback
        tb = traceback.format_exc()
        log.exception("Unexpected error in background thread")
        _exc_msg = f"Unexpected error:\n{exc}\n\n{tb}"

        def _on_exception():
            progress_bar.stop()
            progress_bar.grid_remove()
            generate_btn.config(state="normal")
            messagebox.showerror("Error", _exc_msg)
        root.after(0, _on_exception)


# ===========================================================================
# MAIN ACTION
# ===========================================================================

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
            "Please browse via Help > Advanced or uncheck the option.")
        return
    if include_trend and not prev_report_path:
        messagebox.showerror("Error",
            "Trend tracking is enabled but no previous report selected.\n"
            "Please browse for a previous dashboard or uncheck the option.")
        return

    try:
        threshold = float(score_var.get())
    except ValueError:
        messagebox.showerror("Error",
            f"Minimum CVE Score must be a number (e.g. 9.0).\nCurrent value: {score_var.get()!r}")
        return

    if not show_all_dates_var.get() and date_var.get().strip():
        try:
            datetime.strptime(date_var.get().strip(), "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Error",
                f"Stale date must be in dd/mm/yyyy format.\nCurrent value: {date_var.get()!r}")
            return

    output_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")]
    )
    if not output_path:
        log.info("User cancelled save dialog")
        return

    cutoff_date = None if show_all_dates_var.get() else date_var.get().strip() or None

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
        threshold              = threshold,
        cutoff_date            = cutoff_date,
        show_all_dates         = show_all_dates_var.get(),
        sync_baselines         = sync_baselines_var.get(),
        report_month           = report_month_var.get().strip(),
    )

    log.info("Starting dashboard generation: %s", output_path)
    generate_btn.config(state="disabled")
    progress_bar.grid()
    progress_bar.start(12)
    threading.Thread(target=_run_in_thread, args=(request, progress_bar), daemon=True).start()


# ===========================================================================
# TOGGLE HELPERS
# ===========================================================================

def toggle_rmm_state():
    state = tk.DISABLED if skip_rmm_var.get() else tk.NORMAL
    rmm_entry.config(state=state)
    rmm_browse_btn.config(state=state)

def toggle_date_state():
    date_entry.config(state=tk.DISABLED if show_all_dates_var.get() else tk.NORMAL)

def toggle_trend_state():
    state = tk.NORMAL if include_trend_var.get() else tk.DISABLED
    prev_report_entry.config(state=state)
    prev_report_browse_btn.config(state=state)


# ===========================================================================
# HELP MENU ACTIONS
# ===========================================================================

def _find_cve_repo() -> Path:
    default = Path(_CVE_REPO_DEFAULT)
    if default.exists():
        return default
    here = Path(sys.argv[0]).resolve().parent
    for c in (here / "cvelistV5", here.parent / "cvelistV5"):
        if c.exists():
            return c
    return default


def update_cve_list():
    repo = _find_cve_repo()

    def _do_pull():
        try:
            r = subprocess.run(
                ["git", "-C", str(repo), "pull"],
                capture_output=True, text=True, timeout=120,
            )
            out = r.stdout.strip() or r.stderr.strip() or "(no output)"
            ok  = r.returncode == 0
            def _show():
                if ok:
                    messagebox.showinfo("Update CVEs", f"\u2714  CVE list updated.\n\n{out}")
                else:
                    messagebox.showerror("Update CVEs",
                        f"git pull returned exit code {r.returncode}.\n\n{out}")
            root.after(0, _show)
        except FileNotFoundError:
            root.after(0, lambda: messagebox.showerror(
                "Update CVEs", "git not found.\nEnsure Git is installed and on your PATH."))
        except subprocess.TimeoutExpired:
            root.after(0, lambda: messagebox.showerror(
                "Update CVEs", "git pull timed out after 120 seconds."))
        except Exception as exc:
            _m = str(exc)
            root.after(0, lambda: messagebox.showerror("Update CVEs", f"Unexpected error:\n{_m}"))

    threading.Thread(target=_do_pull, daemon=True).start()
    messagebox.showinfo("Update CVEs",
        f"Pulling latest CVEs from:\n{repo}\n\nThis runs in the background\u2026")


def show_about():
    messagebox.showinfo(
        "About \u2014 N-able CVE Dashboard",
        "N-able CVE Dashboard & Triage Tool\n\n"
        "Automates month-over-month vulnerability triage from N-able exports.\n\n"
        "Features:\n"
        "  \u2022 Patch match & evidence scoring\n"
        "  \u2022 Stale device purge from trend math\n"
        "  \u2022 CVE enrichment via NVD / cvelistV5\n"
        "  \u2022 Redetection tracking & root-cause diagnostics\n\n"
        "\u00a9 2026 Stuart Villanti \u2014 MIT Licence",
    )


def open_advanced_dialog():
    """
    Help > Advanced  --  Patch Report options in a modal dialog.
    Patch data is optional for most runs so it lives here rather than
    cluttering the main window.
    """
    dlg = tk.Toplevel(root)
    dlg.title("Advanced \u2014 Patch Report Options")
    dlg.resizable(False, False)
    dlg.grab_set()   # modal

    dlg.update_idletasks()
    pw = root.winfo_x() + root.winfo_width()  // 2
    ph = root.winfo_y() + root.winfo_height() // 2
    dlg.geometry(f"520x300+{pw - 260}+{ph - 150}")

    PAD = {"padx": 14, "pady": (6, 0)}

    tk.Label(dlg, text="Patch Report Options",
             font=("Arial", 11, "bold")).pack(pady=(12, 6))

    # ── Patch Report ──────────────────────────────────────────────────────────
    tk.Label(dlg, text="Patch Report  (CSV or XLSX)",
             font=("Arial", 9, "bold")).pack(anchor="w", **PAD)
    pf  = tk.Frame(dlg); pf.pack(fill="x", padx=14)
    _pe = tk.Entry(pf, textvariable=patch_var, width=44,
                   state=tk.NORMAL if include_patch_var.get() else tk.DISABLED)
    _pe.pack(side=tk.LEFT)
    _pb = tk.Button(pf, text="Browse", command=lambda: select_file(patch_var),
                    state=tk.NORMAL if include_patch_var.get() else tk.DISABLED)
    _pb.pack(side=tk.LEFT, padx=4)

    def _toggle_p():
        s = tk.NORMAL if include_patch_var.get() else tk.DISABLED
        _pe.config(state=s); _pb.config(state=s)
        _refresh_status()

    tk.Checkbutton(dlg, text="Include Patch Report matching",
                   variable=include_patch_var, command=_toggle_p).pack(anchor="w", padx=14)

    # ── Patch Failure Report ──────────────────────────────────────────────────
    tk.Label(dlg, text="Patch Failure Report  (CSV)",
             font=("Arial", 9, "bold")).pack(anchor="w", **PAD)
    ff  = tk.Frame(dlg); ff.pack(fill="x", padx=14)
    _fe = tk.Entry(ff, textvariable=failure_var, width=44,
                   state=tk.NORMAL if include_failure_var.get() else tk.DISABLED)
    _fe.pack(side=tk.LEFT)
    _fb = tk.Button(ff, text="Browse",
                    command=lambda: select_file(failure_var, [("CSV Files", "*.csv")]),
                    state=tk.NORMAL if include_failure_var.get() else tk.DISABLED)
    _fb.pack(side=tk.LEFT, padx=4)

    def _toggle_f():
        s = tk.NORMAL if include_failure_var.get() else tk.DISABLED
        _fe.config(state=s); _fb.config(state=s)
        _refresh_status()

    tk.Checkbutton(dlg, text="Include Patch Failure analysis",
                   variable=include_failure_var, command=_toggle_f).pack(anchor="w", padx=14)

    # ── Status ────────────────────────────────────────────────────────────────
    _dlg_status_var = tk.StringVar()

    def _refresh_status(*_):
        parts = []
        if include_patch_var.get() and patch_var.get():
            parts.append(f"Patch: {Path(patch_var.get()).name}")
        if include_failure_var.get() and failure_var.get():
            parts.append(f"Failure: {Path(failure_var.get()).name}")
        txt = "  |  ".join(parts) if parts else "No patch data selected"
        _dlg_status_var.set(txt)
        # Keep main-window indicator in sync
        _update_patch_status()

    patch_var.trace_add("write",          _refresh_status)
    failure_var.trace_add("write",        _refresh_status)
    _refresh_status()

    tk.Label(dlg, textvariable=_dlg_status_var,
             font=("Arial", 8), fg="#595959").pack(pady=(8, 0))
    tk.Button(dlg, text="Close", width=10, command=dlg.destroy).pack(pady=(10, 14))


# ===========================================================================
# ROOT WINDOW & WIDGETS
# ===========================================================================

root = tk.Tk()
root.title("N-able CVE Dashboard & Triage Tool")
root.geometry("570x730")
root.resizable(True, True)
root.minsize(520, 600)
root.state("zoomed")   # start maximised on Windows; harmless on other platforms

# ── Menu bar ──────────────────────────────────────────────────────────────────
menubar   = tk.Menu(root)
help_menu = tk.Menu(menubar, tearoff=0)
help_menu.add_command(label="Advanced \u2014 Patch Report Options\u2026", command=open_advanced_dialog)
help_menu.add_separator()
help_menu.add_command(label="Update CVE Data  (git pull cvelistV5)", command=update_cve_list)
help_menu.add_separator()
help_menu.add_command(label="About", command=show_about)
menubar.add_cascade(label="Help", menu=help_menu)
root.config(menu=menubar)

# ── Title ─────────────────────────────────────────────────────────────────────
tk.Label(root, text="N-able CVE Dashboard & Triage Tool",
         font=("Arial", 13, "bold")).pack(pady=(12, 4))

# ── Vulnerability report ──────────────────────────────────────────────────────
tk.Label(root, text="Vulnerability / CVE Report  (CSV or XLSX)",
         font=("Arial", 9, "bold")).pack(anchor="w", padx=14)
vuln_var   = tk.StringVar()
vuln_entry = tk.Entry(root, textvariable=vuln_var, width=55, state="readonly")
vuln_entry.pack(padx=14)
tk.Button(root, text="Browse", command=lambda: select_file(vuln_var)).pack()

# ── RMM / Device inventory ────────────────────────────────────────────────────
tk.Label(root, text="Device Inventory / RMM Report  (CSV or XLSX)",
         font=("Arial", 9, "bold")).pack(anchor="w", padx=14, pady=(8, 0))
rmm_var        = tk.StringVar()
rmm_frame      = tk.Frame(root)
rmm_frame.pack(fill="x", padx=14)
rmm_entry      = tk.Entry(rmm_frame, textvariable=rmm_var, width=44, state="readonly")
rmm_entry.pack(side=tk.LEFT)
rmm_browse_btn = tk.Button(rmm_frame, text="Browse", command=lambda: select_file(rmm_var))
rmm_browse_btn.pack(side=tk.LEFT, padx=4)
skip_rmm_var   = tk.BooleanVar()
tk.Checkbutton(root, text="Skip RMM (CVE export includes device info)",
               variable=skip_rmm_var, command=toggle_rmm_state).pack(anchor="w", padx=14)

# ── CVE score threshold ───────────────────────────────────────────────────────
score_frame = tk.Frame(root)
score_frame.pack(anchor="w", padx=14, pady=(8, 0))
tk.Label(score_frame, text="Minimum CVE Score:", font=("Arial", 9, "bold")).pack(side=tk.LEFT)
score_var = tk.StringVar(value="9.0")
tk.Entry(score_frame, textvariable=score_var, width=6).pack(side=tk.LEFT, padx=6)

# ── Stale-device cutoff date ──────────────────────────────────────────────────
date_frame = tk.Frame(root)
date_frame.pack(anchor="w", padx=14, pady=(6, 0))
tk.Label(date_frame, text="Exclude stale devices last seen before",
         font=("Arial", 9, "bold")).pack(side=tk.LEFT)
date_var = tk.StringVar(value=(date.today() - timedelta(days=90)).strftime('%d/%m/%Y'))
date_entry = tk.Entry(date_frame, textvariable=date_var, width=12)
date_entry.pack(side=tk.LEFT, padx=6)
tk.Label(date_frame, text="(dd/mm/yyyy)").pack(side=tk.LEFT, padx=4)
show_all_dates_var = tk.BooleanVar()
tk.Checkbutton(date_frame, text="Show All Dates",
               variable=show_all_dates_var, command=toggle_date_state).pack(side=tk.LEFT)
toggle_date_state()

# ── Report month ──────────────────────────────────────────────────────────────
month_frame = tk.Frame(root)
month_frame.pack(anchor="w", padx=14, pady=(6, 0))
tk.Label(month_frame, text="Report Month:", font=("Arial", 9, "bold")).pack(side=tk.LEFT)
report_month_var = tk.StringVar(value=datetime.now().strftime('%B %Y'))
tk.Entry(month_frame, textvariable=report_month_var, width=15).pack(side=tk.LEFT, padx=6)

# ── Previous dashboard (trend) ────────────────────────────────────────────────
tk.Label(root, text="Previous Dashboard  (optional, for M-o-M trends)",
         font=("Arial", 9, "bold")).pack(anchor="w", padx=14, pady=(10, 0))
prev_report_var = tk.StringVar()
prev_frame      = tk.Frame(root)
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

# ── Sync baselines ────────────────────────────────────────────────────────────
sync_baselines_var = tk.BooleanVar()
tk.Checkbutton(root, text="Refresh product baselines before run",
               variable=sync_baselines_var).pack(anchor="w", padx=14, pady=(6, 0))

# ── Patch StringVars / hidden widgets (surfaced via Advanced dialog) ──────────
# StringVars persist across dialog opens so paths are remembered session-wide.
patch_var           = tk.StringVar()
failure_var         = tk.StringVar()
include_patch_var   = tk.BooleanVar()
include_failure_var = tk.BooleanVar()

# Proxy widgets referenced by toggle helpers — never packed into root directly.
patch_entry        = tk.Entry(root, textvariable=patch_var,   state="disabled")
patch_browse_btn   = tk.Button(root, text="Browse")
failure_entry      = tk.Entry(root, textvariable=failure_var, state="disabled")
failure_browse_btn = tk.Button(root, text="Browse")

# ── Patch status indicator ────────────────────────────────────────────────────
patch_status_var = tk.StringVar(value="No patch data  (Help \u25b8 Advanced to configure)")

def _update_patch_status(*_):
    parts = []
    if include_patch_var.get() and patch_var.get():
        parts.append(f"Patch: {Path(patch_var.get()).name}")
    if include_failure_var.get() and failure_var.get():
        parts.append(f"Failure: {Path(failure_var.get()).name}")
    patch_status_var.set(
        "  |  ".join(parts) if parts
        else "No patch data  (Help \u25b8 Advanced to configure)"
    )

patch_var.trace_add("write",           _update_patch_status)
failure_var.trace_add("write",         _update_patch_status)
include_patch_var.trace_add("write",   _update_patch_status)
include_failure_var.trace_add("write", _update_patch_status)

tk.Label(root, textvariable=patch_status_var,
         font=("Arial", 8), fg="#595959").pack(anchor="w", padx=14, pady=(4, 0))

# ── Generate button ───────────────────────────────────────────────────────────
generate_btn = tk.Button(
    root,
    text="GENERATE COMPLETE DASHBOARD",
    command=process_reports,
    bg="#0078D7", fg="white",
    font=("Arial", 10, "bold"),
    height=2,
)
generate_btn.pack(pady=14)

# ── Fixed-height progress bar slot — no layout jitter ────────────────────────
_prog_frame = tk.Frame(root, height=24)
_prog_frame.pack(fill="x", padx=14, pady=(0, 6))
_prog_frame.pack_propagate(False)
progress_bar = ttk.Progressbar(_prog_frame, mode="indeterminate")
progress_bar.grid(row=0, column=0, sticky="ew")
_prog_frame.columnconfigure(0, weight=1)
progress_bar.grid_remove()

root.mainloop()
