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

Author : Stu Villanti <s.villanti@kenstra.com>
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
from tkcalendar import Calendar

from orchestrator import DashboardRequest, run as run_dashboard

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
_CARD_FG     = "#1f1f1f"
_CARD_BORDER = "#333333"
_DANGER      = "#d9534f"


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

def _run_in_thread(request):
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
                hide_progress()
                status_var.set("Completed successfully")
                generate_btn.configure(state="normal")
                messagebox.showinfo("Done", _msg)
            root.after(0, _on_success)

        else:
            _err = result.message

            def _on_failure():
                hide_progress()
                status_var.set("Processing failed")
                generate_btn.configure(state="normal")
                messagebox.showerror("Error", f"Processing failed:\n{_err}")
            root.after(0, _on_failure)

    except Exception as exc:
        import traceback
        tb = traceback.format_exc()
        log.exception("Unexpected error in background thread")
        _exc_msg = f"Unexpected error:\n{exc}\n\n{tb}"

        def _on_exception():
            hide_progress()
            status_var.set("Unexpected error")
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
        # date_var is always set by the calendar picker, so format is guaranteed valid.
        # This guard is kept only for safety in case date_var is seeded programmatically.
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
        browser_audit_path     = browser_audit_var.get() or None,
        include_browser_audit  = include_browser_audit_var.get(),
        patch_check_report_path    = patch_check_var.get() or None,
        include_patch_check_report = include_patch_check_var.get(),
        prev_report_path       = prev_report_path or None,
        include_trend          = include_trend,
        threshold              = threshold,
        cutoff_date            = cutoff_date,
        show_all_dates         = show_all_dates_var.get(),
        sync_baselines         = sync_baselines_var.get(),
        report_month           = report_month_var.get().strip(),
        include_health_score   = include_health_score_var.get(),
        advanced_summary       = advanced_summary_var.get(),
    )

    log.info("Starting dashboard generation: %s", output_path)
    generate_btn.configure(state="disabled")
    status_var.set("Generating dashboard... please wait")
    show_progress()
    threading.Thread(target=_run_in_thread, args=(request,), daemon=True).start()


# ===========================================================================
# TOGGLE HELPERS
# ===========================================================================

def toggle_rmm_state():
    state = "disabled" if skip_rmm_var.get() else "normal"
    rmm_entry.configure(state=state)
    rmm_browse_btn.configure(state=state)

def toggle_date_state():
    state = "disabled" if show_all_dates_var.get() else "normal"
    date_entry.configure(state=state)
    # _warn_days_entry no longer exists — "Highlight active devices" UI is disabled.

def toggle_trend_state():
    state = "normal" if include_trend_var.get() else "disabled"
    prev_report_entry.configure(state=state)
    prev_report_browse_btn.configure(state=state)


def open_date_picker():
    """Open a calendar popup and write the chosen date back to date_var (dd/mm/yyyy)."""
    dlg = ctk.CTkToplevel(root)
    dlg.title("Select cutoff date")
    dlg.resizable(False, False)
    dlg.grab_set()

    # Centre over main window
    dlg.update_idletasks()
    pw = root.winfo_x() + root.winfo_width()  // 2
    ph = root.winfo_y() + root.winfo_height() // 2
    dlg.geometry(f"300x310+{pw - 150}+{ph - 155}")

    # Parse existing value so the calendar opens on the right month
    try:
        _initial = datetime.strptime(date_var.get().strip(), "%d/%m/%Y")
        _y, _m, _d = _initial.year, _initial.month, _initial.day
    except ValueError:
        _today = date.today()
        _y, _m, _d = _today.year, _today.month, _today.day

    cal = Calendar(
        dlg,
        selectmode="day",
        year=_y, month=_m, day=_d,
        date_pattern="dd/mm/yyyy",
        background="#1f1f1f",
        foreground="white",
        headersbackground="#1f6aa5",
        headersforeground="white",
        selectbackground="#2fa84f",
        selectforeground="white",
        normalbackground="#1f1f1f",
        normalforeground="white",
        weekendbackground="#1f1f1f",
        weekendforeground="#aaaaaa",
        othermonthbackground="#161616",
        othermonthforeground="gray40",
        bordercolor="#333333",
        font="TkDefaultFont 10",
    )
    cal.pack(padx=12, pady=(12, 4), fill="both", expand=True)

    def _confirm():
        date_var.set(cal.get_date())
        # Update the display label text
        _date_display_var.set(date_var.get())
        dlg.destroy()

    btn_frame = ctk.CTkFrame(dlg, fg_color="transparent")
    btn_frame.pack(fill="x", padx=12, pady=(0, 12))
    ctk.CTkButton(btn_frame, text="Select", command=_confirm,
                  fg_color=_GREEN, hover_color=_GREEN_HOVER).pack(side="left", expand=True, fill="x", padx=(0, 4))
    ctk.CTkButton(btn_frame, text="Cancel", command=dlg.destroy,
                  fg_color="gray40", hover_color="gray30").pack(side="left", expand=True, fill="x", padx=(4, 0))


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
    Advanced Options dialog — baseline refresh, the (beta) Health Score
    toggle, and (commented out for now) patch-evidence options.
    """
    dlg = ctk.CTkToplevel(root)
    dlg.title("Advanced Options")
    dlg.resizable(False, False)
    dlg.grab_set()

    dlg.update_idletasks()
    pw = root.winfo_x() + root.winfo_width()  // 2
    ph = root.winfo_y() + root.winfo_height() // 2
    dlg.geometry(f"560x340+{pw - 280}+{ph - 170}")

    PAD = {"padx": 16, "pady": (6, 0)}

    ctk.CTkLabel(dlg, text="Advanced Options",
                 font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(14, 8))

    # ── Refresh product baselines ────────────────────────────────────────────────
    # Moved here from the main Filters card.
    _baselines_cb = ctk.CTkCheckBox(dlg, text="Refresh product baselines before run",
                                    variable=sync_baselines_var)
    _baselines_cb.pack(anchor="w", padx=16, pady=(4, 4))
    Tooltip(_baselines_cb, "Pulls the latest known-good software versions before "
                           "generating. Only matters if you're using patch-evidence "
                           "matching — adds a little time to the run.")

    # ── Patching Health Score ────────────────────────────────────────────────────
    # Moved here from the main Filters card.
    _health_score_cb = ctk.CTkCheckBox(dlg, text="Show Patching Health Score on Summary sheet  ⚠ Beta",
                                       variable=include_health_score_var)
    _health_score_cb.pack(anchor="w", padx=16, pady=(4, 0))
    Tooltip(_health_score_cb, "Adds an experimental A–F grade to the Summary sheet, "
                              "combining resolution rate, critical-CVE coverage, and "
                              "known-exploit coverage. Methodology may still change — "
                              "not recommended for formal client reporting yet.")
    ctk.CTkLabel(dlg, text="  Scoring methodology is experimental — not for formal reporting",
                 font=("", 11), text_color="#7F6000").pack(anchor="w", padx=16, pady=(0, 4))

    # ── Advanced Summary sections (preview) ──────────────────────────────────────
    _adv_summary_cb = ctk.CTkCheckBox(dlg, text="Advanced Summary sections  ⚠ Preview",
                                      variable=advanced_summary_var)
    _adv_summary_cb.pack(anchor="w", padx=16, pady=(4, 0))
    Tooltip(_adv_summary_cb, "Adds preview sections to the Summary sheet: Multi-Month "
                             "Trend (from the snapshots folder), N-Day Exposure Age "
                             "banding, extra CVE-type rows in Month-over-Month, and Top "
                             "Patch-Gap Root Causes (patch-evidence runs). Also records "
                             "health-score history in snapshots/ so trends accumulate. "
                             "Layout may change while the feature settles.")
    ctk.CTkLabel(dlg, text="  Layout and metrics may change — feedback welcome",
                 font=("", 11), text_color="#7F6000").pack(anchor="w", padx=16, pady=(0, 4))

    # ── Patch Status Check Report ────────────────────────────────────────────────
    # RMM monitoring-check export (e.g. N-able "Failing Checks") listing devices
    # where the Patch Status Check itself is failing to report — distinct from
    # a specific patch failing to install. Adds a "Patch Check Failures" sheet,
    # highlights matching active devices in Top At-Risk Devices, and adds a
    # Summary table so these show up next to the CVE data they may be masking.
    ctk.CTkLabel(dlg, text="Patch Status Check Report  (CSV or XLSX)",
                 font=ctk.CTkFont(weight="bold")).pack(anchor="w", **PAD)
    pcf = ctk.CTkFrame(dlg, fg_color="transparent")
    pcf.pack(fill="x", padx=16)
    _pce = ctk.CTkEntry(pcf, textvariable=patch_check_var, width=380,
                        state="normal" if include_patch_check_var.get() else "disabled")
    _pce.pack(side="left")
    _pcb = ctk.CTkButton(pcf, text="Browse", width=80,
                         command=lambda: select_file(patch_check_var),
                         state="normal" if include_patch_check_var.get() else "disabled")
    _pcb.pack(side="left", padx=6)

    def _toggle_pc():
        s = "normal" if include_patch_check_var.get() else "disabled"
        _pce.configure(state=s)
        _pcb.configure(state=s)

    ctk.CTkCheckBox(dlg, text="Highlight active devices failing patch status check",
                    variable=include_patch_check_var, command=_toggle_pc).pack(anchor="w", padx=16)
    Tooltip(_pce, "A device here has passed CVE detection but RMM's own automated "
                  "patch-status check is failing for it — meaning RMM can't confirm "
                  "whether it's actually patched. Import to flag these on the Summary "
                  "sheet, in Top At-Risk Devices, and in a dedicated 'Patch Check "
                  "Failures' sheet.")

    # ── Patch Report / Patch Failure Report / Browser Audit ──────────────────────
    # Disabled for now (commented out). The include_* / *_var names stay
    # declared elsewhere (defaulting to unset/False) so process_reports(),
    # _update_ready_hint(), and _update_patch_status() don't need changes —
    # they just always see "not configured" while this is off.
    #
    # ctk.CTkLabel(dlg, text="Patch Report  (CSV or XLSX)",
    #              font=ctk.CTkFont(weight="bold")).pack(anchor="w", **PAD)
    # pf = ctk.CTkFrame(dlg, fg_color="transparent")
    # pf.pack(fill="x", padx=16)
    # _pe = ctk.CTkEntry(pf, textvariable=patch_var, width=380,
    #                    state="normal" if include_patch_var.get() else "disabled")
    # _pe.pack(side="left")
    # _pb = ctk.CTkButton(pf, text="Browse", width=80,
    #                     command=lambda: select_file(patch_var),
    #                     state="normal" if include_patch_var.get() else "disabled")
    # _pb.pack(side="left", padx=6)
    #
    # def _toggle_p():
    #     s = "normal" if include_patch_var.get() else "disabled"
    #     _pe.configure(state=s)
    #     _pb.configure(state=s)
    #     _refresh_status()
    #
    # ctk.CTkCheckBox(dlg, text="Include Patch Report matching",
    #                 variable=include_patch_var, command=_toggle_p).pack(anchor="w", padx=16)
    #
    # ctk.CTkLabel(dlg, text="Patch Failure Report  (CSV)",
    #              font=ctk.CTkFont(weight="bold")).pack(anchor="w", **PAD)
    # ff = ctk.CTkFrame(dlg, fg_color="transparent")
    # ff.pack(fill="x", padx=16)
    # _fe = ctk.CTkEntry(ff, textvariable=failure_var, width=380,
    #                    state="normal" if include_failure_var.get() else "disabled")
    # _fe.pack(side="left")
    # _fb = ctk.CTkButton(ff, text="Browse", width=80,
    #                     command=lambda: select_file(failure_var, [("CSV Files", "*.csv")]),
    #                     state="normal" if include_failure_var.get() else "disabled")
    # _fb.pack(side="left", padx=6)
    #
    # def _toggle_f():
    #     s = "normal" if include_failure_var.get() else "disabled"
    #     _fe.configure(state=s)
    #     _fb.configure(state=s)
    #     _refresh_status()
    #
    # ctk.CTkCheckBox(dlg, text="Include Patch Failure analysis",
    #                 variable=include_failure_var, command=_toggle_f).pack(anchor="w", padx=16)
    #
    # ctk.CTkLabel(dlg, text="Browser Audit  (XLSX — from PS browser scan)",
    #              font=ctk.CTkFont(weight="bold")).pack(anchor="w", **PAD)
    # bf = ctk.CTkFrame(dlg, fg_color="transparent")
    # bf.pack(fill="x", padx=16)
    # _be = ctk.CTkEntry(bf, textvariable=browser_audit_var, width=380,
    #                    state="normal" if include_browser_audit_var.get() else "disabled")
    # _be.pack(side="left")
    # _bb = ctk.CTkButton(bf, text="Browse", width=80,
    #                     command=lambda: select_file(browser_audit_var, [("Excel Files", "*.xlsx")]),
    #                     state="normal" if include_browser_audit_var.get() else "disabled")
    # _bb.pack(side="left", padx=6)
    #
    # def _toggle_b():
    #     s = "normal" if include_browser_audit_var.get() else "disabled"
    #     _be.configure(state=s); _bb.configure(state=s)
    #     _refresh_status()
    #
    # ctk.CTkCheckBox(dlg, text="Include Browser Audit (merges into Version Drift)",
    #                 variable=include_browser_audit_var, command=_toggle_b).pack(anchor="w", padx=16)
    #
    # # ── Status ──────────────────────────────────────────────────────────────────
    # _dlg_status_var = tk.StringVar()
    #
    # def _refresh_status(*_):
    #     parts = []
    #     if include_patch_var.get() and patch_var.get():
    #         parts.append(f"Patch: {Path(patch_var.get()).name}")
    #     if include_failure_var.get() and failure_var.get():
    #         parts.append(f"Failure: {Path(failure_var.get()).name}")
    #     if include_browser_audit_var.get() and browser_audit_var.get():
    #         parts.append(f"Browser: {Path(browser_audit_var.get()).name}")
    #     txt = "  |  ".join(parts) if parts else "No patch data selected"
    #     _dlg_status_var.set(txt)
    #     _update_patch_status()
    #
    # patch_var.trace_add("write",        _refresh_status)
    # failure_var.trace_add("write",      _refresh_status)
    # browser_audit_var.trace_add("write",_refresh_status)
    # _refresh_status()
    #
    # ctk.CTkLabel(dlg, textvariable=_dlg_status_var,
    #              text_color=_MUTED_FG, font=ctk.CTkFont(size=11)).pack(pady=(10, 0))
    ctk.CTkButton(dlg, text="Close", width=100,
                  fg_color="gray40", hover_color="gray30",
                  command=dlg.destroy).pack(pady=(12, 16))


# ===========================================================================
# ROOT WINDOW
# ===========================================================================

root = ctk.CTk()
root.title("N-able CVE Dashboard & Triage Tool")
root.geometry("1040x760")
root.resizable(True, True)
root.minsize(860, 640)

# Prefer maximised on Windows, but do not crash on other platforms.
try:
    root.state("zoomed")
except tk.TclError:
    pass

# ── Menu bar (no CTk equivalent — plain tk.Menu works fine on CTk root) ───────
menubar   = tk.Menu(root)
help_menu = tk.Menu(menubar, tearoff=0)
help_menu.add_command(label="Advanced — Patch Report Options…", command=open_advanced_dialog)
help_menu.add_separator()
help_menu.add_command(label="Update CVE Data  (git pull cvelistV5)", command=update_cve_list)
help_menu.add_separator()
help_menu.add_command(label="About", command=show_about)
menubar.add_cascade(label="Help", menu=help_menu)
root.configure(menu=menubar)

root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure(0, weight=1)

# ── Main shell ────────────────────────────────────────────────────────────────
shell = ctk.CTkFrame(root, fg_color="transparent")
shell.grid(row=0, column=0, sticky="nsew", padx=18, pady=18)
shell.grid_columnconfigure(0, weight=1)
shell.grid_rowconfigure(1, weight=1)

# ── Header ───────────────────────────────────────────────────────────────────
header = ctk.CTkFrame(shell, fg_color="transparent")
header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
header.grid_columnconfigure(0, weight=1)

ctk.CTkLabel(
    header,
    text="N-able CVE Dashboard & Triage Tool",
    font=ctk.CTkFont(size=24, weight="bold"),
).grid(row=0, column=0, sticky="w")

ctk.CTkLabel(
    header,
    text="Build a clean Excel dashboard from vulnerability, RMM, patch and trend exports.",
    text_color=_MUTED_FG,
    font=ctk.CTkFont(size=13),
).grid(row=1, column=0, sticky="w", pady=(2, 0))

ctk.CTkButton(
    header,
    text="Advanced Options",
    width=150,
    command=open_advanced_dialog,
).grid(row=0, column=1, rowspan=2, sticky="e", padx=(12, 0))

# ── Scrollable main container ─────────────────────────────────────────────────
_scroll = ctk.CTkScrollableFrame(shell, fg_color="transparent")
_scroll.grid(row=1, column=0, sticky="nsew")
_scroll.grid_columnconfigure(0, weight=1)

# ==========================================================================
# GUI HELPERS
# ==========================================================================

def _card(parent, title, subtitle=None):
    """Create a section card with consistent spacing."""
    frame = ctk.CTkFrame(
        parent,
        fg_color=_CARD_FG,
        border_color=_CARD_BORDER,
        border_width=1,
        corner_radius=14,
    )
    frame.grid(sticky="ew", padx=2, pady=8)
    frame.grid_columnconfigure(0, weight=1)

    ctk.CTkLabel(
        frame,
        text=title,
        font=ctk.CTkFont(size=15, weight="bold"),
    ).grid(row=0, column=0, sticky="w", padx=16, pady=(14, 0))

    row = 1
    if subtitle:
        ctk.CTkLabel(
            frame,
            text=subtitle,
            text_color=_MUTED_FG,
            font=ctk.CTkFont(size=12),
        ).grid(row=1, column=0, sticky="w", padx=16, pady=(2, 8))
        row = 2

    return frame, row


def _file_row(parent, row, variable, browse_command, button_text="Browse", state="normal"):
    frame = ctk.CTkFrame(parent, fg_color="transparent")
    frame.grid(row=row, column=0, sticky="ew", padx=16, pady=(4, 12))
    frame.grid_columnconfigure(0, weight=1)

    entry = ctk.CTkEntry(frame, textvariable=variable, state=state)
    entry.grid(row=0, column=0, sticky="ew")

    button = ctk.CTkButton(frame, text=button_text, width=96, command=browse_command, state=state)
    button.grid(row=0, column=1, padx=(8, 0))
    return entry, button


def _inline_field(parent, row, label, variable, width=120, suffix=None):
    frame = ctk.CTkFrame(parent, fg_color="transparent")
    frame.grid(row=row, column=0, sticky="w", padx=16, pady=(6, 6))
    ctk.CTkLabel(frame, text=label).pack(side="left")
    entry = ctk.CTkEntry(frame, textvariable=variable, width=width)
    entry.pack(side="left", padx=(8, 8))
    if suffix:
        ctk.CTkLabel(frame, text=suffix, text_color=_MUTED_FG, font=ctk.CTkFont(size=11)).pack(side="left")
    return entry


def _filename_or_missing(value, missing="Not selected"):
    return Path(value).name if value else missing


class Tooltip:
    """
    Minimal hover tooltip for any tkinter/customtkinter widget. Shows a
    small dark popup near the cursor after a short delay, matching the
    app's own dark theme rather than the OS default tooltip styling.

    Usage:
        Tooltip(some_widget, "Explanation shown on hover.")
    """
    _DELAY_MS = 450

    def __init__(self, widget, text, wraplength=320):
        self.widget = widget
        self.text = text
        self.wraplength = wraplength
        self._after_id = None
        self._tip_window = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _event=None):
        self._cancel()
        self._after_id = self.widget.after(self._DELAY_MS, self._show)

    def _cancel(self):
        if self._after_id is not None:
            self.widget.after_cancel(self._after_id)
            self._after_id = None

    def _show(self):
        if self._tip_window is not None:
            return
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        except tk.TclError:
            return
        self._tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tw.attributes("-topmost", True)
        label = tk.Label(
            tw, text=self.text, justify="left", wraplength=self.wraplength,
            background="#2b2b2b", foreground="white",
            relief="solid", borderwidth=1,
            padx=8, pady=5, font=("Segoe UI", 9),
        )
        label.pack()

    def _hide(self, _event=None):
        self._cancel()
        if self._tip_window is not None:
            self._tip_window.destroy()
            self._tip_window = None

# ==========================================================================
# VARIABLES
# ==========================================================================

def _default_stale_cutoff() -> str:
    """1st of the previous calendar month (dd/mm/yyyy) — the default stale-device cutoff."""
    first_of_this_month = date.today().replace(day=1)
    last_of_prev_month = first_of_this_month - timedelta(days=1)
    first_of_prev_month = last_of_prev_month.replace(day=1)
    return first_of_prev_month.strftime('%d/%m/%Y')


vuln_var = tk.StringVar()
rmm_var = tk.StringVar()
skip_rmm_var = tk.BooleanVar()
score_var = tk.StringVar(value="1.0")
date_var = tk.StringVar(value=_default_stale_cutoff())
_date_display_var = tk.StringVar(value=date_var.get())   # mirrors date_var for the label
show_all_dates_var = tk.BooleanVar()
report_month_var = tk.StringVar(value=datetime.now().strftime('%B %Y'))
prev_report_var = tk.StringVar()
include_trend_var = tk.BooleanVar(value=True)
sync_baselines_var = tk.BooleanVar()
patch_var = tk.StringVar()
failure_var = tk.StringVar()
browser_audit_var = tk.StringVar()
patch_check_var = tk.StringVar()
include_patch_var = tk.BooleanVar()
include_failure_var = tk.BooleanVar()
include_browser_audit_var = tk.BooleanVar()
include_patch_check_var = tk.BooleanVar()
include_health_score_var = tk.BooleanVar()
advanced_summary_var = tk.BooleanVar()
patch_status_var = tk.StringVar(value="Patch evidence: not configured")
status_var = tk.StringVar(value="Ready")

# ==========================================================================
# REQUIRED INPUTS CARD
# ==========================================================================

inputs_card, row = _card(
    _scroll,
    "1. Required reports",
    "Select the current vulnerability export and the matching RMM/device inventory export.",
)

ctk.CTkLabel(inputs_card, text="Vulnerability / CVE Report  (CSV or XLSX)").grid(
    row=row, column=0, sticky="w", padx=16, pady=(4, 0)
)
row += 1
vuln_entry, vuln_browse_btn = _file_row(
    inputs_card,
    row,
    vuln_var,
    lambda: select_file(vuln_var),
)
Tooltip(vuln_entry, "The vulnerability detections export from N-able (CSV or XLSX).\n"
                    "This should be this month's export — the tool only ever sees "
                    "currently active detections, so a patched CVE simply disappears "
                    "from next month's file rather than being marked resolved in it.")
row += 1

ctk.CTkLabel(inputs_card, text="Device Inventory / RMM Report  (CSV or XLSX)").grid(
    row=row, column=0, sticky="w", padx=16, pady=(2, 0)
)
row += 1
rmm_entry, rmm_browse_btn = _file_row(
    inputs_card,
    row,
    rmm_var,
    lambda: select_file(rmm_var),
)
Tooltip(rmm_entry, "The device inventory / RMM export from N-able (CSV or XLSX). "
                   "Used to fill in device type, username, last check-in, and to "
                   "identify devices that are stale or missing from RMM entirely.")
row += 1

# Skip RMM checkbox — disabled for now (commented out). skip_rmm_var stays
# declared (defaults to False) so nothing else in the form has to change;
# RMM data is always required until this is re-enabled.
# ctk.CTkCheckBox(
#     inputs_card,
#     text="Skip RMM — CVE export already includes device information",
#     variable=skip_rmm_var,
#     command=toggle_rmm_state,
# ).grid(row=row, column=0, sticky="w", padx=16, pady=(0, 14))

# ==========================================================================
# FILTERS CARD
# ==========================================================================

filters_card, row = _card(
    _scroll,
    "2. Filters and reporting scope",
    "Set the CVSS threshold and stale-device handling before generating the workbook.",
)

score_entry = _inline_field(filters_card, row, "Minimum CVE Score:", score_var, width=80, suffix="Example: 9.0")
Tooltip(score_entry, "CVSS floor for what gets included in the report. Lower values "
                     "(e.g. 1.0) pull in far more data — every severity, not just "
                     "critical — and take longer to generate. Use 9.0 for a "
                     "critical-only report, 7.0 for critical + high.")
row += 1

# Stale date row
_date_frame = ctk.CTkFrame(filters_card, fg_color="transparent")
_date_frame.grid(row=row, column=0, sticky="w", padx=16, pady=(4, 6))
ctk.CTkLabel(_date_frame, text="Exclude stale devices last seen before:").pack(side="left")
# Read-only date display — updated by the calendar picker
_date_display_label = ctk.CTkLabel(
    _date_frame,
    textvariable=_date_display_var,
    text_color="white",
    font=ctk.CTkFont(size=13, weight="bold"),
    width=100,
    anchor="w",
)
_date_display_label.pack(side="left", padx=(8, 4))
date_entry = ctk.CTkButton(
    _date_frame,
    text="📅 Pick date",
    width=110,
    command=open_date_picker,
    fg_color=_BLUE,
    hover_color=_BLUE_HOVER,
)
date_entry.pack(side="left", padx=(0, 8))
Tooltip(date_entry, "Devices whose last check-in is before this date are excluded "
                    "from the Active scope and moved to the Stale Excluded Devices "
                    "tab instead of being silently dropped. Defaults to the 1st of "
                    "last month — adjust it if you need a different cutoff.")
_show_all_dates_cb = ctk.CTkCheckBox(
    _date_frame,
    text="Show all dates",
    variable=show_all_dates_var,
    command=toggle_date_state,
)
_show_all_dates_cb.pack(side="left", padx=(8, 0))
Tooltip(_show_all_dates_cb, "Ignore the stale-date cutoff entirely and include every "
                            "device regardless of how long it's been offline.")
row += 1

row += 1

report_month_entry = _inline_field(filters_card, row, "Report Month:", report_month_var, width=160)
Tooltip(report_month_entry, "The label shown on the report (e.g. \"April 2026\"). "
                            "Defaults to the current month — override this when "
                            "generating a report retroactively.")
row += 1

# "Refresh product baselines before run" and "Show Patching Health Score on
# Summary sheet" moved to Advanced Options — see open_advanced_dialog().

# ==========================================================================
# OPTIONAL DATA CARD
# ==========================================================================

optional_card, row = _card(
    _scroll,
    "3. Optional evidence and trend data",
    "Use these when you want patch evidence, failure analysis, or month-over-month comparison.",
)

ctk.CTkLabel(optional_card, text="Previous Dashboard  (optional — month-over-month trends)").grid(
    row=row, column=0, sticky="w", padx=16, pady=(4, 0)
)
row += 1
prev_report_entry, prev_report_browse_btn = _file_row(
    optional_card,
    row,
    prev_report_var,
    lambda: select_file(prev_report_var, [("Previous Report", "*.xlsx *.csv"), ("Excel Files", "*.xlsx"), ("CSV RawData", "*.csv")]),
    state="disabled",
)
Tooltip(prev_report_entry, "Last month's GENERATED dashboard output (not a raw N-able "
                           "export). Enables the Month-over-Month Remediation Summary "
                           "and trend detail sheets on this run.")
row += 1
_trend_cb = ctk.CTkCheckBox(
    optional_card,
    text="Include month-over-month trend analysis",
    variable=include_trend_var,
    command=toggle_trend_state,
)
_trend_cb.grid(row=row, column=0, sticky="w", padx=16, pady=(0, 8))
Tooltip(_trend_cb, "On by default. Untick if you don't have a previous dashboard "
                   "to compare against, or don't want trend sections on this run.")
row += 1

patch_status_label = ctk.CTkLabel(
    optional_card,
    textvariable=patch_status_var,
    text_color=_MUTED_FG,
    font=ctk.CTkFont(size=12),
)
patch_status_label.grid(row=row, column=0, sticky="w", padx=16, pady=(0, 14))

# ==========================================================================
# RUN CARD
# ==========================================================================

run_card, row = _card(
    _scroll,
    "4. Generate dashboard",
    "You will be prompted where to save the Excel workbook after clicking generate.",
)

generate_btn = ctk.CTkButton(
    run_card,
    text="GENERATE COMPLETE DASHBOARD",
    command=process_reports,
    fg_color=_GREEN,
    hover_color=_GREEN_HOVER,
    font=ctk.CTkFont(size=15, weight="bold"),
    height=48,
    corner_radius=16,
)
generate_btn.grid(row=row, column=0, sticky="ew", padx=16, pady=(4, 10))
Tooltip(generate_btn, "Builds the workbook and prompts you for where to save it. "
                      "Generation time depends on your threshold and fleet size — "
                      "a low threshold across a large fleet can take a minute or more.")
row += 1

status_line = ctk.CTkLabel(
    run_card,
    textvariable=status_var,
    text_color=_MUTED_FG,
    font=ctk.CTkFont(size=12),
)
status_line.grid(row=row, column=0, sticky="w", padx=16, pady=(0, 8))
row += 1

_prog_frame = ctk.CTkFrame(run_card, fg_color="transparent")
progress_bar = ctk.CTkProgressBar(_prog_frame, mode="indeterminate")
progress_bar.grid(row=0, column=0, sticky="ew")
_prog_frame.grid_columnconfigure(0, weight=1)


def show_progress():
    _prog_frame.grid(row=row, column=0, sticky="ew", padx=16, pady=(0, 16))
    progress_bar.start()


def hide_progress():
    progress_bar.stop()
    _prog_frame.grid_remove()

hide_progress()

# ==========================================================================
# STATUS / STATE UPDATES
# ==========================================================================

def _update_patch_status(*_):
    parts = []
    if include_patch_var.get():
        parts.append(f"Patch: {_filename_or_missing(patch_var.get())}")
    if include_failure_var.get():
        parts.append(f"Failure: {_filename_or_missing(failure_var.get())}")
    if include_browser_audit_var.get():
        parts.append(f"Browser: {_filename_or_missing(browser_audit_var.get())}")
    patch_status_var.set(
        "Patch evidence: " + "  |  ".join(parts)
        if parts else
        "Patch evidence: not configured"
    )


def _update_ready_hint(*_):
    missing = []
    if not vuln_var.get():
        missing.append("CVE report")
    if not skip_rmm_var.get() and not rmm_var.get():
        missing.append("RMM report")
    if include_trend_var.get() and not prev_report_var.get():
        missing.append("previous dashboard")
    if include_patch_var.get() and not patch_var.get():
        missing.append("patch report")
    if include_failure_var.get() and not failure_var.get():
        missing.append("patch failure report")
    if include_browser_audit_var.get() and not browser_audit_var.get():
        missing.append("browser audit")

    if generate_btn.cget("state") == "disabled":
        return
    status_var.set("Ready" if not missing else "Waiting for: " + ", ".join(missing))


for _var in (
    vuln_var, rmm_var, skip_rmm_var, prev_report_var, include_trend_var,
    patch_var, failure_var, include_patch_var, include_failure_var,
    browser_audit_var, include_browser_audit_var,
):
    _var.trace_add("write", _update_ready_hint)

for _var in (patch_var, failure_var, include_patch_var, include_failure_var,
             browser_audit_var, include_browser_audit_var):
    _var.trace_add("write", _update_patch_status)

# Apply initial state.
toggle_date_state()
toggle_trend_state()
toggle_rmm_state()
_update_patch_status()
_update_ready_hint()

root.mainloop()