"""Patch Report Comparison Tool.

Compares two N-able patch overview exports and writes a highlighted copy of the
second (current) report so you can see at a glance what is Installed, Installing,
Missing, Pending or Reboot Required - and which rows are new since the previous
report.

The main focus is follow-up: every patch that was Failed or Missing in report 1
is tracked through to its status in report 2 and listed on its own sheet, so you
can see what was fixed, what is still outstanding and what dropped off the report.

Rows are matched between reports on Device + Patch name.

Requires: openpyxl (pip install openpyxl)
"""

from __future__ import annotations

import datetime as dt
import html
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover - surfaced to the user in the GUI
    openpyxl = None


APP_TITLE = "Patch Report Comparison"

# ---------------------------------------------------------------------------
# Highlight styles.  Colours mirror Excel's built-in cell styles but are written
# as explicit fills so they render the same regardless of the workbook theme.
# A fill of None means font colour only (Excel's "Warning Text" behaves this way).
# ---------------------------------------------------------------------------


class Hl:
    """A highlight: display label plus fill and font colours."""

    def __init__(self, label, fill, font, bold=False):
        self.label = label
        self.fill_hex = fill
        self.font_hex = font
        self.bold = bold
        self._fill = None
        self._font = None

    @property
    def fill(self):
        if self._fill is None and self.fill_hex:
            self._fill = PatternFill("solid", start_color=self.fill_hex,
                                     end_color=self.fill_hex)
        return self._fill

    @property
    def font(self):
        if self._font is None:
            self._font = Font(color=self.font_hex, bold=self.bold)
        return self._font

    def apply(self, cell):
        if self.fill is not None:
            cell.fill = self.fill
        cell.font = self.font

    def swatch_colours(self):
        """(background, foreground) for the tkinter preview swatch."""
        return "#" + (self.fill_hex or "FFFFFF"), "#" + self.font_hex


def build_styles():
    return {
        "Installed": Hl("Good - green", "C6EFCE", "006100"),
        "Installing": Hl("Note - yellow", "FFEB9C", "9C6500"),
        "Missing": Hl("Bad - red", "FFC7CE", "9C0006"),
        "Reboot Required": Hl("Accent1 - light blue", "BDD7EE", "1F4E79"),
        "Failed": Hl("Accent2 - orange", "F8CBAD", "833C0C"),
        "Pending": Hl("Warning Text - red text", None, "C00000"),
        "Ignored": Hl("Accent5 - teal", "4BACC6", "FFFFFF"),
        "__new__": Hl("Accent3 - green", "A9D08E", "375623"),
    }


# Statuses the user can choose to highlight, in display order.
HIGHLIGHT_STATUSES = ["Installed", "Installing", "Missing", "Reboot Required",
                      "Failed", "Pending", "Ignored"]

# Statuses in report 1 that are tracked through to report 2.
FOCUS_STATUSES = ("Failed", "Missing")

# Change classifications written to the "Change" column.
CHG_NEW_DEVICE = "New device"
CHG_NEW_PATCH = "New patch"
CHG_CHANGED = "Status changed"
CHG_UNCHANGED = "Unchanged"

# Follow-up outcomes for patches that were Failed or Missing in report 1.
OUT_RESOLVED = "Resolved"
OUT_IN_PROGRESS = "In progress"
OUT_OUTSTANDING = "Outstanding"
OUT_DROPPED = "No longer reported"
OUTCOMES = (OUT_OUTSTANDING, OUT_IN_PROGRESS, OUT_RESOLVED, OUT_DROPPED)

OUTCOME_STYLE = {
    OUT_RESOLVED: "Installed",       # green
    OUT_IN_PROGRESS: "Installing",   # yellow
    OUT_OUTSTANDING: "Missing",      # red
    OUT_DROPPED: "Reboot Required",  # light blue
}

# Report-2 statuses that count as still-broken / in-flight / fixed.
STILL_BROKEN = {"failed", "missing"}
IN_FLIGHT = {"pending", "installing", "reboot required"}

REQUIRED_COLUMNS = ["Client", "Site", "Device", "Status", "Patch"]
DATE_COLUMN = "Discovered / Install Date"

FILTER_ALL = "all"
FILTER_CHANGED = "changed"
FILTER_FOCUS = "focus"


# ---------------------------------------------------------------------------
# Data loading / comparison
# ---------------------------------------------------------------------------


def clean(value):
    """Normalise a cell value for comparison (HTML entities, case, whitespace)."""
    if value is None:
        return ""
    text = html.unescape(str(value))
    return re.sub(r"\s+", " ", text).strip().lower()


def display(value):
    """Tidy a cell value for output (report 1 exports are HTML-escaped)."""
    if isinstance(value, str):
        return html.unescape(value).strip()
    return value


def find_header_row(ws, max_scan=10):
    """Return (row_index, {lowercase_name: column_index}) for the header row."""
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), start=1):
        if row is None:
            continue
        lookup = {}
        for col_idx, cell in enumerate(row):
            if cell is None:
                continue
            lookup[re.sub(r"\s+", " ", str(cell)).strip().lower()] = col_idx
        if all(name.lower() in lookup for name in REQUIRED_COLUMNS):
            return row_idx, lookup
    return None, None


def load_report(path, log):
    """Read a patch report into (rows, column_lookup)."""
    log(f"Reading {os.path.basename(path)} ...")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        header_row, lookup = find_header_row(ws)
        if lookup is None:
            raise ValueError(
                f"{os.path.basename(path)}: could not find a header row containing "
                + ", ".join(REQUIRED_COLUMNS)
            )
        rows = []
        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            if row is None or all(cell is None for cell in row):
                continue
            rows.append(row)
        title = ws.title
    finally:
        wb.close()
    log(f"  {len(rows):,} data rows, sheet '{title}'")
    return rows, lookup


def cell_at(row, lookup, name):
    idx = lookup.get(name.lower())
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def row_key(row, lookup):
    return (clean(cell_at(row, lookup, "Device")), clean(cell_at(row, lookup, "Patch")))


def index_previous(rows, lookup):
    """Build the lookups needed to compare report 2 against report 1.

    Returns (statuses_by_key, devices, focus_by_key) where focus_by_key holds the
    report-1 rows whose status was Failed or Missing.
    """
    statuses = {}
    devices = set()
    focus = {}
    for row in rows:
        key = row_key(row, lookup)
        status = str(cell_at(row, lookup, "Status") or "").strip()
        devices.add(key[0])
        statuses.setdefault(key, set()).add(status)
        if status in FOCUS_STATUSES:
            focus.setdefault(key, {"row": row, "statuses": set()})["statuses"].add(status)
    return statuses, devices, focus


def classify(status, key, prev_statuses, prev_devices):
    """Return (change, previous_status_text) for a current-report row."""
    previous = prev_statuses.get(key)
    if previous is None:
        if key[0] not in prev_devices:
            return CHG_NEW_DEVICE, ""
        return CHG_NEW_PATCH, ""
    prev_text = " / ".join(sorted(s for s in previous if s))
    if status in previous:
        return CHG_UNCHANGED, prev_text
    return CHG_CHANGED, prev_text


def outcome_for(status):
    """Follow-up outcome for a patch that was Failed or Missing in report 1."""
    normalised = status.strip().lower()
    if normalised in STILL_BROKEN:
        return OUT_OUTSTANDING
    if normalised in IN_FLIGHT:
        return OUT_IN_PROGRESS
    return OUT_RESOLVED


# ---------------------------------------------------------------------------
# Output workbook
# ---------------------------------------------------------------------------

def _style_header(ws, columns, widths):
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", start_color="44546A", end_color="44546A")
    for col_idx, name in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center")
    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.freeze_panes = "A2"


def _finish_sheet(ws, columns, last_row):
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(last_row, 1)}"


def write_output(path, rows, lookup, prev_statuses, prev_devices, prev_focus,
                 prev_lookup, options, log, progress):
    styles = build_styles()
    new_style = styles["__new__"]
    selected = {s for s in HIGHLIGHT_STATUSES if options["statuses"].get(s)}

    columns = ["Client", "Site", "Device", "Status", "Patch", DATE_COLUMN,
               "Previous Status", "Change", "Follow-up"]
    widths = [26, 20, 24, 16, 62, 20, 18, 16, 20]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Patch Report"
    _style_header(ws, columns, widths)

    status_counts = {}
    change_counts = {}
    transitions = {}
    outcome_counts = {}
    followup_rows = []   # rows for the follow-up sheet
    seen_focus = set()
    written = 0
    total = len(rows)
    out_row = 2

    for i, row in enumerate(rows):
        if i % 500 == 0:
            progress(i, total)

        key = row_key(row, lookup)
        status = str(cell_at(row, lookup, "Status") or "").strip()
        change, prev_text = classify(status, key, prev_statuses, prev_devices)
        is_new = change in (CHG_NEW_DEVICE, CHG_NEW_PATCH)

        status_counts[status or "(blank)"] = status_counts.get(status or "(blank)", 0) + 1
        change_counts[change] = change_counts.get(change, 0) + 1
        if change == CHG_CHANGED:
            label = f"{prev_text} -> {status}"
            transitions[label] = transitions.get(label, 0) + 1

        focus = prev_focus.get(key)
        outcome = ""
        if focus is not None:
            seen_focus.add(key)
            outcome = outcome_for(status)
            outcome_counts[outcome] = outcome_counts.get(outcome, 0) + 1
            was = " / ".join(sorted(focus["statuses"]))
            followup_rows.append([
                display(cell_at(row, lookup, "Client")),
                display(cell_at(row, lookup, "Site")),
                display(cell_at(row, lookup, "Device")),
                display(cell_at(row, lookup, "Patch")),
                was,
                status,
                outcome,
                cell_at(row, lookup, DATE_COLUMN),
            ])

        if options["filter_mode"] == FILTER_CHANGED and change == CHG_UNCHANGED:
            continue
        if options["filter_mode"] == FILTER_FOCUS and focus is None:
            continue

        values = [
            display(cell_at(row, lookup, "Client")),
            display(cell_at(row, lookup, "Site")),
            display(cell_at(row, lookup, "Device")),
            status,
            display(cell_at(row, lookup, "Patch")),
            cell_at(row, lookup, DATE_COLUMN),
            prev_text,
            change,
            outcome,
        ]

        status_hl = styles.get(status) if status in selected else None
        row_hl = None
        if is_new and options["highlight_new"] and options["new_fills_row"]:
            row_hl = new_style
        elif status_hl is not None and options["whole_row"]:
            row_hl = status_hl

        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=out_row, column=col_idx, value=value)
            if isinstance(value, (dt.datetime, dt.date)):
                cell.number_format = "dd-mmm-yyyy"
            if row_hl is not None:
                row_hl.apply(cell)

        # When the row is not filled wholesale, the Status cell still carries
        # its own colour.
        if status_hl is not None and row_hl is None:
            status_hl.apply(ws.cell(row=out_row, column=4))

        # New rows are flagged in the Change column even when the row is filled
        # by status, so nothing new can hide behind a status colour.
        if is_new and options["highlight_new"] and not options["new_fills_row"]:
            new_style.apply(ws.cell(row=out_row, column=8))

        # Follow-up outcome always gets its own colour - this is the focus.
        if outcome:
            styles[OUTCOME_STYLE[outcome]].apply(ws.cell(row=out_row, column=9))

        out_row += 1
        written += 1

    progress(total, total)
    _finish_sheet(ws, columns, out_row - 1)

    # Anything that was Failed or Missing in report 1 and is absent from report 2.
    for key, focus in prev_focus.items():
        if key in seen_focus:
            continue
        row = focus["row"]
        outcome_counts[OUT_DROPPED] = outcome_counts.get(OUT_DROPPED, 0) + 1
        followup_rows.append([
            display(cell_at(row, prev_lookup, "Client")),
            display(cell_at(row, prev_lookup, "Site")),
            display(cell_at(row, prev_lookup, "Device")),
            display(cell_at(row, prev_lookup, "Patch")),
            " / ".join(sorted(focus["statuses"])),
            "",
            OUT_DROPPED,
            None,
        ])

    _write_followup(wb, followup_rows, styles)
    _write_summary(wb, status_counts, change_counts, transitions, outcome_counts,
                   selected, options, styles, total, written, len(prev_focus))

    wb.save(path)
    log(f"Saved {written:,} rows to {path}")
    return {
        "status_counts": status_counts,
        "change_counts": change_counts,
        "transitions": transitions,
        "outcome_counts": outcome_counts,
        "focus_total": len(prev_focus),
        "written": written,
    }


def _write_followup(wb, followup_rows, styles):
    ws = wb.create_sheet("Failed & Missing Follow-up")
    columns = ["Client", "Site", "Device", "Patch", "Report 1 Status",
               "Report 2 Status", "Outcome", "Discovered / Install Date"]
    _style_header(ws, columns, [26, 20, 24, 62, 16, 18, 20, 22])

    order = {name: i for i, name in enumerate(OUTCOMES)}
    followup_rows.sort(key=lambda r: (order.get(r[6], 99), str(r[0]), str(r[2]), str(r[3])))

    for out_row, values in enumerate(followup_rows, start=2):
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=out_row, column=col_idx, value=value)
            if isinstance(value, (dt.datetime, dt.date)):
                cell.number_format = "dd-mmm-yyyy"
        style = styles.get(values[5]) if values[5] else None
        if style is not None:
            style.apply(ws.cell(row=out_row, column=6))
        styles[OUTCOME_STYLE[values[6]]].apply(ws.cell(row=out_row, column=7))

    _finish_sheet(ws, columns, len(followup_rows) + 1)


def _write_summary(wb, status_counts, change_counts, transitions, outcome_counts,
                   selected, options, styles, total, written, focus_total):
    ws = wb.create_sheet("Summary")
    bold = Font(bold=True)
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 12

    def heading(row, text):
        cell = ws.cell(row=row, column=1, value=text)
        cell.font = Font(bold=True, size=12)
        return row + 1

    r = 1
    r = heading(r, "Failed / Missing in report 1 - follow-up")
    ws.cell(row=r, column=1, value="Tracked patches").font = bold
    ws.cell(row=r, column=2, value=focus_total).font = bold
    r += 1
    for outcome in OUTCOMES:
        cell = ws.cell(row=r, column=1, value=outcome)
        styles[OUTCOME_STYLE[outcome]].apply(cell)
        ws.cell(row=r, column=2, value=outcome_counts.get(outcome, 0))
        r += 1
    ws.cell(row=r, column=1,
            value="See the 'Failed & Missing Follow-up' sheet for the detail.")
    r += 2

    r = heading(r, "Legend")
    for status in HIGHLIGHT_STATUSES:
        style = styles[status]
        style.apply(ws.cell(row=r, column=1, value=status))
        note = style.label if status in selected else f"{style.label} - not highlighted"
        ws.cell(row=r, column=2, value=note)
        r += 1
    new_style = styles["__new__"]
    new_style.apply(ws.cell(row=r, column=1, value="New since previous report"))
    ws.cell(row=r, column=2,
            value=new_style.label if options["highlight_new"] else "Accent3 - not highlighted")
    r += 2

    r = heading(r, "Current status totals")
    ws.cell(row=r, column=1, value="Status").font = bold
    ws.cell(row=r, column=2, value="Rows").font = bold
    r += 1
    for status, count in sorted(status_counts.items(), key=lambda kv: -kv[1]):
        ws.cell(row=r, column=1, value=status)
        ws.cell(row=r, column=2, value=count)
        r += 1
    r += 1

    r = heading(r, "Change against previous report")
    for change in (CHG_UNCHANGED, CHG_CHANGED, CHG_NEW_PATCH, CHG_NEW_DEVICE):
        ws.cell(row=r, column=1, value=change)
        ws.cell(row=r, column=2, value=change_counts.get(change, 0))
        r += 1
    r += 1

    if transitions:
        r = heading(r, "Status transitions")
        ws.cell(row=r, column=1, value="From -> To").font = bold
        ws.cell(row=r, column=2, value="Rows").font = bold
        r += 1
        for label, count in sorted(transitions.items(), key=lambda kv: -kv[1]):
            ws.cell(row=r, column=1, value=label)
            ws.cell(row=r, column=2, value=count)
            r += 1
        r += 1

    filter_labels = {
        FILTER_ALL: "All rows",
        FILTER_CHANGED: "New or changed only",
        FILTER_FOCUS: "Failed / Missing in report 1 only",
    }
    r = heading(r, "Run details")
    for label, value in [
        ("Rows in current report", total),
        ("Rows written", written),
        ("Rows filtered", filter_labels[options["filter_mode"]]),
        ("Highlight scope", "Entire row" if options["whole_row"] else "Status cell only"),
        ("New rows", "Fill entire row" if options["new_fills_row"] else "Mark Change column"),
        ("Generated", dt.datetime.now().strftime("%Y-%m-%d %H:%M")),
    ]:
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=value)
        r += 1


def run_comparison(previous_path, current_path, output_path, options, log, progress):
    prev_rows, prev_lookup = load_report(previous_path, log)
    curr_rows, curr_lookup = load_report(current_path, log)

    log("Indexing previous report ...")
    prev_statuses, prev_devices, prev_focus = index_previous(prev_rows, prev_lookup)
    log(f"  {len(prev_statuses):,} device/patch pairs across {len(prev_devices):,} devices")
    log(f"  {len(prev_focus):,} patches were Failed or Missing")

    log("Comparing and writing output ...")
    return write_output(output_path, curr_rows, curr_lookup, prev_statuses, prev_devices,
                        prev_focus, prev_lookup, options, log, progress)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------


class PatchReportApp(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=12)
        self.master = master
        self.grid(row=0, column=0, sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.previous_path = tk.StringVar()
        self.current_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status_vars = {s: tk.BooleanVar(value=True) for s in HIGHLIGHT_STATUSES}
        self.highlight_new = tk.BooleanVar(value=True)
        self.new_fills_row = tk.BooleanVar(value=False)
        self.whole_row = tk.BooleanVar(value=True)
        self.filter_mode = tk.StringVar(value=FILTER_ALL)
        self.open_when_done = tk.BooleanVar(value=True)
        self.status_text = tk.StringVar(value="Select both reports to begin.")
        self.last_dir = os.path.expanduser("~")
        self.queue = queue.Queue()
        self.worker = None
        self.styles = build_styles() if openpyxl is not None else {}

        self._build_files()
        self._build_options()
        self._build_log()
        self._build_actions()
        self.after(100, self._drain_queue)

    # -- layout ------------------------------------------------------------
    def _build_files(self):
        frame = ttk.LabelFrame(self, text="Reports", padding=10)
        frame.grid(row=0, column=0, sticky="ew")
        frame.columnconfigure(1, weight=1)

        rows = [
            ("Report 1 (previous):", self.previous_path, self._browse_previous),
            ("Report 2 (current):", self.current_path, self._browse_current),
            ("Output file:", self.output_path, self._browse_output),
        ]
        for r, (label, var, command) in enumerate(rows):
            ttk.Label(frame, text=label).grid(row=r, column=0, sticky="w", pady=3)
            ttk.Entry(frame, textvariable=var).grid(row=r, column=1, sticky="ew", padx=6, pady=3)
            ttk.Button(frame, text="Browse...", command=command).grid(row=r, column=2, pady=3)

        ttk.Label(
            frame,
            text=("Report 2 is the one that gets highlighted. Rows match on Device + Patch. "
                  "Patches that were Failed or Missing in report 1 get their own sheet."),
            foreground="#555555", wraplength=740, justify="left",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(6, 0))

    def _swatch(self, parent, style):
        bg, fg = style.swatch_colours()
        return tk.Label(parent, text=" Aa ", bg=bg, fg=fg, relief="solid",
                        borderwidth=1, font=("Segoe UI", 8))

    def _build_options(self):
        frame = ttk.LabelFrame(self, text="Highlighting", padding=10)
        frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        frame.columnconfigure(0, weight=1)

        grid = ttk.Frame(frame)
        grid.grid(row=0, column=0, sticky="w")
        entries = [(s, self.styles[s].label if self.styles else "") for s in HIGHLIGHT_STATUSES]
        entries.append(("__new__", self.styles["__new__"].label if self.styles else ""))
        for i, (status, label) in enumerate(entries):
            r, c = i // 2, (i % 2) * 3
            text = "New since report 1" if status == "__new__" else status
            var = self.highlight_new if status == "__new__" else self.status_vars[status]
            if self.styles:
                self._swatch(grid, self.styles[status]).grid(row=r, column=c, padx=(0, 6), pady=2)
            ttk.Checkbutton(grid, text=text, variable=var).grid(
                row=r, column=c + 1, sticky="w")
            ttk.Label(grid, text=label, foreground="#777777").grid(
                row=r, column=c + 2, sticky="w", padx=(6, 28))

        ttk.Separator(frame, orient="horizontal").grid(row=1, column=0, sticky="ew", pady=8)

        opts = ttk.Frame(frame)
        opts.grid(row=2, column=0, sticky="w")
        ttk.Checkbutton(opts, text="Colour the entire row (otherwise the Status cell only)",
                        variable=self.whole_row).grid(row=0, column=0, sticky="w", pady=2)
        ttk.Checkbutton(opts, text="Fill new rows entirely with green (otherwise flag the Change column)",
                        variable=self.new_fills_row).grid(row=1, column=0, sticky="w", pady=2)
        ttk.Checkbutton(opts, text="Open the output file when finished",
                        variable=self.open_when_done).grid(row=2, column=0, sticky="w", pady=2)

        rows_frame = ttk.Frame(frame)
        rows_frame.grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Label(rows_frame, text="Rows to output:").grid(row=0, column=0, sticky="w", padx=(0, 10))
        for i, (mode, text) in enumerate([
            (FILTER_ALL, "All"),
            (FILTER_CHANGED, "New or changed only"),
            (FILTER_FOCUS, "Failed / Missing in report 1 only"),
        ]):
            ttk.Radiobutton(rows_frame, text=text, value=mode,
                            variable=self.filter_mode).grid(row=0, column=i + 1, sticky="w", padx=(0, 16))

    def _build_log(self):
        frame = ttk.LabelFrame(self, text="Progress", padding=10)
        frame.grid(row=2, column=0, sticky="nsew", pady=(10, 0))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        self.rowconfigure(2, weight=1)

        self.progress = ttk.Progressbar(frame, mode="determinate", maximum=100)
        self.progress.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        self.log_widget = tk.Text(frame, height=12, wrap="none", state="disabled",
                                  font=("Consolas", 9))
        self.log_widget.grid(row=1, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(frame, orient="vertical", command=self.log_widget.yview)
        scroll.grid(row=1, column=1, sticky="ns")
        self.log_widget.configure(yscrollcommand=scroll.set)

    def _build_actions(self):
        frame = ttk.Frame(self)
        frame.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        frame.columnconfigure(0, weight=1)
        ttk.Label(frame, textvariable=self.status_text, foreground="#555555").grid(
            row=0, column=0, sticky="w")
        self.run_button = ttk.Button(frame, text="Compare && Highlight", command=self.start)
        self.run_button.grid(row=0, column=1, padx=(8, 0))
        ttk.Button(frame, text="Close", command=self.master.destroy).grid(row=0, column=2, padx=(8, 0))

    # -- file pickers ------------------------------------------------------
    def _ask_open(self, title):
        path = filedialog.askopenfilename(
            title=title, initialdir=self.last_dir,
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if path:
            self.last_dir = os.path.dirname(path)
        return path

    def _browse_previous(self):
        path = self._ask_open("Select report 1 (previous)")
        if path:
            self.previous_path.set(path)
            self._suggest_output()

    def _browse_current(self):
        path = self._ask_open("Select report 2 (current)")
        if path:
            self.current_path.set(path)
            self._suggest_output()

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save highlighted report as", initialdir=self.last_dir,
            defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            self.output_path.set(path)

    def _suggest_output(self):
        current = self.current_path.get()
        if current and not self.output_path.get():
            base, _ = os.path.splitext(current)
            self.output_path.set(f"{base}_highlighted.xlsx")

    # -- worker plumbing ---------------------------------------------------
    def log(self, message):
        self.queue.put(("log", message))

    def _drain_queue(self):
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "log":
                    self.log_widget.configure(state="normal")
                    self.log_widget.insert("end", payload + "\n")
                    self.log_widget.see("end")
                    self.log_widget.configure(state="disabled")
                elif kind == "progress":
                    done, total = payload
                    self.progress["value"] = (done / total * 100) if total else 0
                elif kind == "status":
                    self.status_text.set(payload)
                elif kind == "done":
                    self._finish(payload)
        except queue.Empty:
            pass
        self.after(100, self._drain_queue)

    def start(self):
        if openpyxl is None:
            messagebox.showerror(APP_TITLE, "openpyxl is not installed.\n\nRun: pip install openpyxl")
            return
        previous, current, output = (self.previous_path.get().strip(),
                                     self.current_path.get().strip(),
                                     self.output_path.get().strip())
        for label, path in (("Report 1", previous), ("Report 2", current)):
            if not path:
                messagebox.showwarning(APP_TITLE, f"Please select {label}.")
                return
            if not os.path.isfile(path):
                messagebox.showerror(APP_TITLE, f"{label} not found:\n{path}")
                return
        if not output:
            messagebox.showwarning(APP_TITLE, "Please choose an output file.")
            return
        if os.path.abspath(output) in (os.path.abspath(previous), os.path.abspath(current)):
            messagebox.showerror(APP_TITLE, "The output file must not overwrite a source report.")
            return

        options = {
            "statuses": {s: v.get() for s, v in self.status_vars.items()},
            "highlight_new": self.highlight_new.get(),
            "new_fills_row": self.new_fills_row.get(),
            "whole_row": self.whole_row.get(),
            "filter_mode": self.filter_mode.get(),
        }

        self.run_button.configure(state="disabled")
        self.progress["value"] = 0
        self.log_widget.configure(state="normal")
        self.log_widget.delete("1.0", "end")
        self.log_widget.configure(state="disabled")
        self.queue.put(("status", "Working..."))

        self.worker = threading.Thread(
            target=self._work, args=(previous, current, output, options), daemon=True)
        self.worker.start()

    def _work(self, previous, current, output, options):
        try:
            result = run_comparison(
                previous, current, output, options, self.log,
                lambda done, total: self.queue.put(("progress", (done, total))))
            self.queue.put(("done", ("ok", output, result)))
        except PermissionError:
            self.queue.put(("done", ("error", output,
                                     "Could not write the output file - it may be open in Excel.")))
        except Exception as exc:  # surfaced in the GUI
            self.queue.put(("done", ("error", output, str(exc))))

    def _finish(self, payload):
        outcome, output, result = payload
        self.run_button.configure(state="normal")
        if outcome == "error":
            self.progress["value"] = 0
            self.status_text.set("Failed.")
            self.log(f"ERROR: {result}")
            messagebox.showerror(APP_TITLE, result)
            return

        self.progress["value"] = 100
        focus_total = result["focus_total"]
        outcome_counts = result["outcome_counts"]
        self.log("")
        self.log(f"Failed / Missing in report 1: {focus_total:,} patches")
        for name in OUTCOMES:
            self.log(f"  {name:<20}{outcome_counts.get(name, 0):>8,}")
        self.log("")
        self.log("Current status totals:")
        for status, count in sorted(result["status_counts"].items(), key=lambda kv: -kv[1]):
            self.log(f"  {status:<20}{count:>8,}")
        self.log("")
        self.log("Against the previous report:")
        for change in (CHG_UNCHANGED, CHG_CHANGED, CHG_NEW_PATCH, CHG_NEW_DEVICE):
            self.log(f"  {change:<20}{result['change_counts'].get(change, 0):>8,}")
        transitions = result["transitions"]
        if transitions:
            self.log("")
            self.log("Top status transitions:")
            for label, count in sorted(transitions.items(), key=lambda kv: -kv[1])[:10]:
                self.log(f"  {label:<44}{count:>6,}")
        self.log("")
        self.log(f"Done - {result['written']:,} rows written.")
        self.status_text.set(
            f"Done - {result['written']:,} rows written, "
            f"{outcome_counts.get(OUT_OUTSTANDING, 0):,} still outstanding.")

        if self.open_when_done.get():
            try:
                if sys.platform.startswith("win"):
                    os.startfile(output)  # noqa: S606
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", output])
                else:
                    subprocess.Popen(["xdg-open", output])
            except Exception as exc:
                self.log(f"Could not open the output file: {exc}")


def main():
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("880x800")
    root.minsize(760, 660)
    try:
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    PatchReportApp(root)
    if openpyxl is None:
        messagebox.showerror(APP_TITLE, "openpyxl is not installed.\n\nRun: pip install openpyxl")
    root.mainloop()


if __name__ == "__main__":
    main()
