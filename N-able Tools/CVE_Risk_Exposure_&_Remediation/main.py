"""
main.py — CustomTkinter GUI.

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
from tkinter import filedialog, messagebox

import customtkinter as ctk

from orchestrator import DashboardRequest, DashboardResult, run as run_dashboard

# ── Appearance ────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(name)s - %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

_CVE_REPO_DEFAULT = r"C:\NoCScripts\N-able Tools\CVE_Risk_Exposure_&_Remediation\cvelistV5"

# ── Colour tokens ─────────────────────────────────────────────────────────────
_BLUE        = "#1f6aa5"
_BLUE_HOVER  = "#144f7a"
_GREEN       = "#2fa84f"
_GREEN_HOVER = "#237a3b"
_MUTED_FG    = "gray60"


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
                generate_btn.configure(state="normal")
                messagebox.showinfo("Done", _msg)
            root.after(0, _on_success)

        else:
            _err = result.message

            def _on_failure():
                progress_bar.stop()
                progress_bar.grid_remove()
                generate_btn.configure(state="normal")
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
            generate_btn.configure(state="normal")
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
    generate_btn.configure(state="disabled")
    progress_bar.grid()
    progress_bar.start()
    threading.Thread(target=_run_in_thread, args=(request, progress_bar), daemon=True).start()


# ===========================================================================
# TOGGLE HELPERS
# ===========================================================================

def toggle_rmm_state():
    state = "disabled" if skip_rmm_var.get() else "normal"
    rmm_entry.configure(state=state)
    rmm_browse_btn.configure(state=state)

def toggle_date_state():
    date_entry.configure(state="disabled" if show_all_dates_var.get() else "normal")

def toggle_trend_state():
    state = "normal" if include_trend_var.get() else "disabled"
    prev_report_entry.configure(state=state)
    prev_report_browse_btn.configure(state=state)


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
    """
    dlg = ctk.CTkToplevel(root)
    dlg.title("Advanced \u2014 Patch Report Options")
    dlg.resizable(False, False)
    dlg.grab_set()

    dlg.update_idletasks()
    pw = root.winfo_x() + root.winfo_width()  // 2
    ph = root.winfo_y() + root.winfo_height() // 2
    dlg.geometry(f"560x320+{pw - 280}+{ph - 160}")

    PAD = {"padx": 16, "pady": (6, 0)}

    ctk.CTkLabel(dlg, text="Patch Report Options",
                 font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(14, 8))

    # ── Patch Report ──────────────────────────────────────────────────────────
    ctk.CTkLabel(dlg, text="Patch Report  (CSV or XLSX)",
                 font=ctk.CTkFont(weight="bold")).pack(anchor="w", **PAD)
    pf = ctk.CTkFrame(dlg, fg_color="transparent")
    pf.pack(fill="x", padx=16)
    _pe = ctk.CTkEntry(pf, textvariable=patch_var, width=380,
                       state="normal" if include_patch_var.get() else "disabled")
    _pe.pack(side="left")
    _pb = ctk.CTkButton(pf, text="Browse", width=80,
                        command=lambda: select_file(patch_var),
                        state="normal" if include_patch_var.get() else "disabled")
    _pb.pack(side="left", padx=6)

    def _toggle_p():
        s = "normal" if include_patch_var.get() else "disabled"
        _pe.configure(state=s)
        _pb.configure(state=s)
        _refresh_status()

    ctk.CTkCheckBox(dlg, text="Include Patch Report matching",
                    variable=include_patch_var, command=_toggle_p).pack(anchor="w", padx=16)

    # ── Patch Failure Report ──────────────────────────────────────────────────
    ctk.CTkLabel(dlg, text="Patch Failure Report  (CSV)",
                 font=ctk.CTkFont(weight="bold")).pack(anchor="w", **PAD)
    ff = ctk.CTkFrame(dlg, fg_color="transparent")
    ff.pack(fill="x", padx=16)
    _fe = ctk.CTkEntry(ff, textvariable=failure_var, width=380,
                       state="normal" if include_failure_var.get() else "disabled")
    _fe.pack(side="left")
    _fb = ctk.CTkButton(ff, text="Browse", width=80,
                        command=lambda: select_file(failure_var, [("CSV Files", "*.csv")]),
                        state="normal" if include_failure_var.get() else "disabled")
    _fb.pack(side="left", padx=6)

    def _toggle_f():
        s = "normal" if include_failure_var.get() else "disabled"
        _fe.configure(state=s)
        _fb.configure(state=s)
        _refresh_status()

    ctk.CTkCheckBox(dlg, text="Include Patch Failure analysis",
                    variable=include_failure_var, command=_toggle_f).pack(anchor="w", padx=16)

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
        _update_patch_status()

    patch_var.trace_add("write",   _refresh_status)
    failure_var.trace_add("write", _refresh_status)
    _refresh_status()

    ctk.CTkLabel(dlg, textvariable=_dlg_status_var,
                 text_color=_MUTED_FG, font=ctk.CTkFont(size=11)).pack(pady=(10, 0))
    ctk.CTkButton(dlg, text="Close", width=100,
                  fg_color="gray40", hover_color="gray30",
                  command=dlg.destroy).pack(pady=(12, 16))


# ===========================================================================
# ROOT WINDOW
# ===========================================================================

root = ctk.CTk()
root.title("N-able CVE Dashboard & Triage Tool")
root.geometry("600x680")
root.resizable(True, True)
root.minsize(520, 600)
root.state("zoomed")

# ── Menu bar (no CTk equivalent — plain tk.Menu works fine on CTk root) ───────
menubar   = tk.Menu(root)
help_menu = tk.Menu(menubar, tearoff=0)
help_menu.add_command(label="Advanced \u2014 Patch Report Options\u2026", command=open_advanced_dialog)
help_menu.add_separator()
help_menu.add_command(label="Update CVE Data  (git pull cvelistV5)", command=update_cve_list)
help_menu.add_separator()
help_menu.add_command(label="About", command=show_about)
menubar.add_cascade(label="Help", menu=help_menu)
root.configure(menu=menubar)

# ── Scrollable main container ─────────────────────────────────────────────────
# Wrapping everything in a CTkScrollableFrame means the window stays usable
# even on small screens without clipping any widgets.
_scroll = ctk.CTkScrollableFrame(root, fg_color="transparent")
_scroll.pack(fill="both", expand=True, padx=0, pady=0)

def _section_label(parent, text):
    """Bold section header with consistent top padding."""
    ctk.CTkLabel(parent, text=text,
                 font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=16, pady=(10, 2))

def _muted_label(parent, textvariable=None, text=""):
    """Small muted-colour helper/status label."""
    kwargs = dict(text_color=_MUTED_FG, font=ctk.CTkFont(size=11))
    if textvariable:
        kwargs["textvariable"] = textvariable
    else:
        kwargs["text"] = text
    ctk.CTkLabel(parent, **kwargs).pack(anchor="w", padx=16, pady=(2, 0))

# ── Title ─────────────────────────────────────────────────────────────────────
ctk.CTkLabel(
    _scroll,
    text="N-able CVE Dashboard & Triage Tool",
    font=ctk.CTkFont(size=18, weight="bold"),
).pack(pady=(16, 6))

# ── Vulnerability report ──────────────────────────────────────────────────────
_section_label(_scroll, "Vulnerability / CVE Report  (CSV or XLSX)")
vuln_var = tk.StringVar()
_vf = ctk.CTkFrame(_scroll, fg_color="transparent")
_vf.pack(fill="x", padx=16)
vuln_entry = ctk.CTkEntry(_vf, textvariable=vuln_var, state="disabled", width=440)
vuln_entry.pack(side="left")
ctk.CTkButton(_vf, text="Browse", width=90,
              command=lambda: select_file(vuln_var)).pack(side="left", padx=6)

# ── RMM / Device inventory ────────────────────────────────────────────────────
_section_label(_scroll, "Device Inventory / RMM Report  (CSV or XLSX)")
rmm_var = tk.StringVar()
_rf = ctk.CTkFrame(_scroll, fg_color="transparent")
_rf.pack(fill="x", padx=16)
rmm_entry = ctk.CTkEntry(_rf, textvariable=rmm_var, state="disabled", width=380)
rmm_entry.pack(side="left")
rmm_browse_btn = ctk.CTkButton(_rf, text="Browse", width=90,
                                command=lambda: select_file(rmm_var))
rmm_browse_btn.pack(side="left", padx=6)
skip_rmm_var = tk.BooleanVar()
ctk.CTkCheckBox(_scroll, text="Skip RMM (CVE export includes device info)",
                variable=skip_rmm_var, command=toggle_rmm_state).pack(anchor="w", padx=16, pady=(4, 0))

# ── CVE score threshold ───────────────────────────────────────────────────────
_section_label(_scroll, "Filters")
_sf = ctk.CTkFrame(_scroll, fg_color="transparent")
_sf.pack(anchor="w", padx=16)
ctk.CTkLabel(_sf, text="Minimum CVE Score:").pack(side="left")
score_var = tk.StringVar(value="9.0")
ctk.CTkEntry(_sf, textvariable=score_var, width=70).pack(side="left", padx=8)

# ── Stale-device cutoff date ──────────────────────────────────────────────────
_df = ctk.CTkFrame(_scroll, fg_color="transparent")
_df.pack(anchor="w", padx=16, pady=(6, 0))
ctk.CTkLabel(_df, text="Exclude stale devices last seen before:").pack(side="left")
date_var = tk.StringVar(value=(date.today() - timedelta(days=90)).strftime('%d/%m/%Y'))
date_entry = ctk.CTkEntry(_df, textvariable=date_var, width=110)
date_entry.pack(side="left", padx=8)
ctk.CTkLabel(_df, text="(dd/mm/yyyy)", text_color=_MUTED_FG,
             font=ctk.CTkFont(size=11)).pack(side="left")
show_all_dates_var = tk.BooleanVar()
ctk.CTkCheckBox(_df, text="Show All Dates",
                variable=show_all_dates_var, command=toggle_date_state).pack(side="left", padx=10)
toggle_date_state()

# ── Report month ──────────────────────────────────────────────────────────────
_mf = ctk.CTkFrame(_scroll, fg_color="transparent")
_mf.pack(anchor="w", padx=16, pady=(6, 0))
ctk.CTkLabel(_mf, text="Report Month:").pack(side="left")
report_month_var = tk.StringVar(value=datetime.now().strftime('%B %Y'))
ctk.CTkEntry(_mf, textvariable=report_month_var, width=130).pack(side="left", padx=8)

# ── Previous dashboard (trend) ────────────────────────────────────────────────
_section_label(_scroll, "Previous Dashboard  (optional — month-over-month trends)")
prev_report_var = tk.StringVar()
_pf = ctk.CTkFrame(_scroll, fg_color="transparent")
_pf.pack(fill="x", padx=16)
prev_report_entry = ctk.CTkEntry(_pf, textvariable=prev_report_var,
                                  state="disabled", width=380)
prev_report_entry.pack(side="left")
prev_report_browse_btn = ctk.CTkButton(
    _pf, text="Browse", width=90, state="disabled",
    command=lambda: select_file(prev_report_var, [("Excel Files", "*.xlsx")]),
)
prev_report_browse_btn.pack(side="left", padx=6)
include_trend_var = tk.BooleanVar()
ctk.CTkCheckBox(_scroll, text="Include month-over-month trend analysis",
                variable=include_trend_var, command=toggle_trend_state).pack(anchor="w", padx=16, pady=(4, 0))

# ── Sync baselines ────────────────────────────────────────────────────────────
sync_baselines_var = tk.BooleanVar()
ctk.CTkCheckBox(_scroll, text="Refresh product baselines before run",
                variable=sync_baselines_var).pack(anchor="w", padx=16, pady=(6, 0))

# ── Patch StringVars / hidden state ──────────────────────────────────────────
patch_var           = tk.StringVar()
failure_var         = tk.StringVar()
include_patch_var   = tk.BooleanVar()
include_failure_var = tk.BooleanVar()

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

_muted_label(_scroll, textvariable=patch_status_var)

# ── Separator ─────────────────────────────────────────────────────────────────
ctk.CTkFrame(_scroll, height=2, fg_color="gray30").pack(
    fill="x", padx=16, pady=(14, 0))

# ── Generate button ───────────────────────────────────────────────────────────
generate_btn = ctk.CTkButton(
    _scroll,
    text="GENERATE COMPLETE DASHBOARD",
    command=process_reports,
    fg_color=_GREEN,
    hover_color=_GREEN_HOVER,
    font=ctk.CTkFont(size=14, weight="bold"),
    height=44,
    corner_radius=18,
)
generate_btn.pack(padx=16, pady=14, fill="x")

# ── Progress bar ──────────────────────────────────────────────────────────────
_prog_frame = ctk.CTkFrame(_scroll, height=20, fg_color="transparent")
_prog_frame.pack(fill="x", padx=16, pady=(0, 10))
_prog_frame.pack_propagate(False)
progress_bar = ctk.CTkProgressBar(_prog_frame, mode="indeterminate")
progress_bar.grid(row=0, column=0, sticky="ew")
_prog_frame.columnconfigure(0, weight=1)
progress_bar.grid_remove()

root.mainloop()