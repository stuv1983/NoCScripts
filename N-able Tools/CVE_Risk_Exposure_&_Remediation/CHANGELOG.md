# N-able CVE Dashboard & Triage Tool — Full Technical Changelog

---

## Architecture Overview

The tool is organised into three tiers: pipeline control, data & enrichment, and Excel sheet builders. Nineteen modules in total.

**Pipeline & entrypoints**

| Module | Role |
|---|---|
| `main.py` | CustomTkinter GUI — collects inputs, spawns background thread, streams progress |
| `run_dashboard.py` | Command-line entrypoint — same pipeline as the GUI without launching Tk |
| `orchestrator.py` | Pipeline coordinator — receives a `DashboardRequest`, runs the pipeline, writes the workbook |
| `config.py` | Loads `config.json` (product map, version rules, cvss cache) and exposes shared constants |

**Data & enrichment (no GUI, no xlsxwriter)**

| Module | Role |
|---|---|
| `data_pipeline.py` | Data engine — load, merge, patch-match, trend arithmetic |
| `resolution.py` | Single source of truth for "is this CVE/device row resolved?" — shared by every consumer |
| `diagnostics.py` | Root-cause classification — patch lag, version drift, detection mismatches |
| `snapshot.py` | History — lightweight JSON snapshots for month-over-month trend tracking |
| `cve_lookup.py` | CVE enrichment — NVD / CVE.org / OSV / cvelistV5 lookups |
| `version_sync.py` | Baseline sync — fetches rolling stable versions from vendor APIs |

**Excel sheet builders (no pandas loading, no GUI — DataFrames in, sheets out)**

| Module | Role |
|---|---|
| `excel_builder.py` | Workbook writer — top-level `build_report()`, coordinates all sheet builders |
| `summary_sheet.py` | Summary sheet — Patching Health Score, Key Metrics, Resolution Status, Device Breakdown, Top At-Risk Devices, Month-over-Month Remediation |
| `product_sheets.py` | Per-product ☑/☐ triage tabs plus the "Patch Confirmed" variant |
| `trend_sheets.py` | Trend Summary, New This Month, Persisting CVEs, Resolved Since Previous Report |
| `device_sheets.py` | All Detections, Stale Excluded Devices, CVEs on Stale Devices, Device Inventory, Raw Data |
| `patch_sheets.py` | Patch Evidence Notes, Patch Lag, Version Drift, Patch Failures, Products Not in Patch Scope |
| `formatting.py` | Named colour palette and shared xlsxwriter format factories |
| `sheet_helpers.py` | Small xlsxwriter writing helpers (CVE/NVD links, health-score subtotal block) shared across sheet builders |
| `sheet_names.py` | Reserved sheet names — imported by both `orchestrator.py` and `data_pipeline.py` to keep name collisions impossible |

---

## Release History

---

### v0.1 — Initial Build
**Files:** all

The tool existed as a single monolithic Python script (`N-able_CVE_Dashboard_7.py`). All logic — loading, merging, trend math, Excel writing — lived in one file. No GUI. Run via command line with hardcoded paths.

---

### v0.2 — Architecture Refactor: Modular Split

**Files:** `orchestrator.py`, `data_pipeline.py`, `excel_builder.py`, `main.py`

Split the monolith into the current module architecture. Tkinter GUI added. Each module given strict responsibilities with zero cross-contamination (no Excel logic in data_pipeline, no data processing in main.py).

---

### v0.3 — Core Pipeline Fixes

#### Triage Scope Bug (Critical)
**File:** `orchestrator.py`

`filtered_df` included RESOLVED rows, causing product sheets and "New This Month" to treat resolved Office detections as active. 

**Fix:** Introduced a clean two-scope system:
- `filtered_df` — all rows above score threshold (evidence/history scope)
- `triage_df` — UNRESOLVED only, not-in-RMM excluded (active triage scope)

Product sheets, trend math, and exposure counts all use `triage_df`. Raw Data and Resolved sheets use `filtered_df`.

#### Score Threshold Default (Critical)
**Files:** `orchestrator.py`, `main.py`, `run_dashboard.py`

Default threshold was `9.0` but was being compared against the wrong column in some code paths. Hardcoded to `9.0` consistently across all entry points with a `float()` cast, with the CVSS `Vulnerability Score` column explicitly targeted.

**Fix:** `run_dashboard.py` CLI default aligned to `9.0` (was `1.0`). All three entry points now use the same default and target the same column.

#### RMM Merge — LEFT vs INNER Join
**File:** `orchestrator.py`

`exclude_missing_rmm=True` caused an INNER join that silently dropped devices not in the RMM inventory before any filtering ran. 

**Fix:** Default changed to `exclude_missing_rmm=False`. Unmatched devices are marked `'Not Found in RMM'` in the `Last Response` column and excluded from triage sheets but retained in Raw Data.

---

### v0.4 — Patch Matching Engine

#### Status Column Collision (Critical Bug)
**File:** `data_pipeline.py` — `process_patch_match()`

Both the CVE export and the patch report have a `Status` column. After the pandas merge with `suffixes=('', '_p')`, the CVE status (`RESOLVED`/`UNRESOLVED`) stayed as `Status` while the patch install status (`Installed`/`Pending`) became `Status_p`. Three classifier functions — `_classify_version_check`, `_classify_resolution`, `_classify_baseline_compliance` — all read `row.get('Status', '')`, so they silently consumed the CVE threat status instead of the patch install status. Every single row returned `'Patch not yet installed'` regardless of actual patch state.

**Fix:** Renamed patch `Status` to `_patch_status` before the merge, so classifiers read `row.get('_patch_status', '')`. After classification, `_patch_status` is renamed back to `Status_p` for display. This fixed 763 rows that should have shown `Patch confirmed` but were showing `Unresolved`.

#### `_RULES` Mapping for Pending/Installing/Missing/Failed States
**File:** `diagnostics.py`

`Matched - installing`, `Matched - pending`, `Matched - missing`, `Matched - failed` all mapped to `cause = None`, silently dropping those rows from `root_cause_df`. The `DISPLAY_MAP` and `_PENALTIES` entries for `patch_installing`, `patch_pending`, `patch_missing` already existed but were never reached.

**Fix:** These four statuses now map to their correct internal cause codes so they appear in the evidence summary and contribute to health scores.

#### Microsoft Office 365 Product Key Bug
**File:** `excel_builder.py`

`get_base_product()` stripped `365` as a version suffix, so `'Microsoft Office 365'` resolved to `'office'` instead of `'office365'`, causing sheet lookup failures and Office CVEs not being marked resolved even when the raw data showed `RESOLVED`.

**Fix:** `_sheet_pk` now derived from the raw `Affected Products` column values in each group, not from the base product name.

#### Resolved Sheet Consolidation
**Files:** `excel_builder.py`, `orchestrator.py`

`'Patch Confirmed'` and `'Resolved (Patch Confirmed)'` were separate sheets with overlapping content. The orchestrator now pre-merges all raw `Status=RESOLVED` rows with patch-confirmed rows before writing. Single sheet: `'Resolved (Patch Confirmed)'`.

---

### v0.5 — Trend Engine: Ghost Ticket Fix

#### Raw Data as Absolute Source of Truth
**File:** `data_pipeline.py` — `compute_trends()`

Previous behaviour: `compute_trends` implicitly trusted manual `_Checkbox_Resolved` checkboxes from the previous month's report. If a CVE was ticked resolved last month, it was hidden from the Persisting set even if the raw scanner still marked it UNRESOLVED.

**Fix:** `compute_trends` no longer filters based on previous checkboxes. If a CVE is UNRESOLVED in raw data, it persists. The checkbox data is read separately and used only for re-detection tracking — never to exclude rows from arithmetic.

#### `load_previous_report` Tuple Return
**File:** `data_pipeline.py`

`load_previous_report` previously attached `_Checkbox_Resolved` as a column on the returned DataFrame, which allowed it to flow into `_active_trend_scope` and corrupt trend arithmetic.

**Fix:** The function now returns `(df, resolved_pairs)` — a clean DataFrame with no checkbox contamination, plus a standalone `set` of `(device, cve_id)` tuples. `compute_trends` receives `prev_resolved_pairs` as a keyword parameter and uses it only for re-detection flagging.

#### `_active_trend_scope` — Stale and Not-in-RMM Exclusion
**File:** `data_pipeline.py`

Stale devices were not being fully purged from trend math in all code paths. Not-in-RMM devices could also bleed into trend arithmetic when `skip_rmm=True` or no RMM was provided.

**Fix:** `_active_trend_scope()` now:
1. Filters `Not Found in RMM` rows explicitly (guards on column presence, logs debug note if absent for older report format)
2. Applies `inventory_devices` filter
3. Applies `stale_devices` filter

All three exclusions are now unconditional and applied in the same function so every caller uses identical logic.

#### Snapshot Month Key Bug
**File:** `snapshot.py`

`load_history()` year calculation used `now.year - (now.month - 1 - i < 0)` — boolean subtraction that only ever subtracts one year, producing wrong keys for any window spanning more than one year boundary.

**Fix:** `yr = now.year + (now.month - 1 - i) // 12` — Python's floor division correctly handles any number of months backward across arbitrary year boundaries.

#### `report_month` Added to Snapshot Records
**File:** `snapshot.py`

Snapshots were keyed by OS execution month, so generating "April" in May would be stored under `2026-05`. The user-supplied `report_month` label is now parsed and used as the aggregate file key.

---

### v0.6 — Data Transparency: Waterfall Reconciliation

#### `build_client_summary_sheet` Signature Change (Critical)
**File:** `excel_builder.py`

Old signature: `(workbook, filtered_df, trend_data=None, ...)`. The orchestrator was calling with `(wb, filtered_df, triage_df, threshold, trend_data=trend_data, ...)`. Python bound `triage_df` to `trend_data` positionally, then hit `trend_data=` as a keyword — "got multiple values for argument 'trend_data'".

**Fix:** New signature: `(workbook, filtered_df, triage_df, threshold, trend_data=None, ...)`. All four positional args explicit; `trend_data` is keyword-only.

#### Data Filtering Reconciliation Waterfall Table
**File:** `excel_builder.py` — `build_client_summary_sheet()`

Added a new table to the Client Summary sheet that makes the filtering math fully transparent:

```
[+]  Total raw detections (all devices, CVSS ≥ threshold)
[-]  Excluded: stale devices (Last Response before <cutoff>)
[-]  Excluded: device not found in RMM
[=]  Active tracked scope (Key Metrics above)
```

Answers the question "where did 2,000 rows go?" with exact numbers per exclusion reason. Note below the table explains why unique Device and CVE Type counts don't subtract as cleanly as row counts (overlap between excluded and active groups).

#### Key Metrics Now Source from `triage_df` Only
**File:** `excel_builder.py`

All Key Metrics (total rows, unique CVEs, unique devices, server count, exploit count, etc.) now read exclusively from `triage_df` (active scope). Previously they read from `filtered_df` which included not-in-RMM rows, inflating every metric.

---

### v0.7 — Not-in-RMM Audit Tracking

#### Not-in-RMM Devices Added to Stale Sheets
**Files:** `orchestrator.py`, `excel_builder.py`

Devices "Not Found in RMM" were previously excluded silently — counted in the waterfall but not listed anywhere for investigation.

**Fix:** Both stale sheet builders now receive `not_in_rmm_df` and render two clearly labelled sections:

**`build_stale_excluded_sheet`** — two sections:
- `⏱ Date-Stale Devices` — pale amber rows (`#FFFDE7`), amber header (`#FFF2CC`)
- `🚫 Not Found in RMM` — pale red rows (`#FFEBEE`), dark red header (`#C00000`)

**`build_stale_cves_sheet`** — same two-section layout with CVE/NVD hyperlinks colour-matched to their section. Both sections show UNRESOLVED CVEs only, pulled from raw data.

Explanatory note: "N-able is still reporting vulnerabilities for a device absent from the RMM inventory — verify decommission status (shadow IT / orphaned agent)."

#### `not_in_rmm_df` Extracted in Orchestrator
**File:** `orchestrator.py`

Changed from `not_in_rmm = int(...)` (count only) to `not_in_rmm_df = filtered_df[...].copy()` (full DataFrame). The stale sheet trigger changed from `if not stale_excluded.empty` to `if not stale_excluded.empty or not not_in_rmm_df.empty` so either exclusion type independently triggers the sheets.

---

### v0.8 — Username Column

#### Username Propagated from RMM Throughout Pipeline
**File:** `data_pipeline.py`

The RMM Device Inventory report includes a `Username` (logged-on user) column which was never brought into the merged dataset.

**Fix:**
- `load_rmm_data()` detects `Username` under multiple aliases: `username`, `user name`, `logged-on user`, `logged on user`, `current user`, `user` (case-insensitive). Any match is standardised to `'Username'`. If no match exists, the column is created empty.
- `merge_data()` conditionally appends `'Username'` to `rmm_pull` when present in `df_rmm`.
- After merge: `merged['Username'] = merged['Username'].fillna('')` so not-in-RMM rows get blank rather than NaN.
- Username flows automatically into `filtered_df`, `triage_df`, `stale_excluded`, and all downstream sheets. Both stale sheet builders include `'Username'` in `cols_to_keep`.

---

### v0.9 — GUI: Fullscreen, Menu Restructure, Patch Options to Advanced Dialog

**File:** `main.py`

#### Fullscreen Support
- `resizable(True, True)` — both axes freely resizable
- `root.state("zoomed")` — opens maximised on Windows
- `root.minsize(520, 600)` — prevents unusable shrinking

#### Restructured Help Menu
Previous: no menu at all. New structure:
```
Help
  ├─ Advanced — Patch Report Options…
  ├─ ─────────────────────────────
  ├─ Update CVE Data  (git pull cvelistV5)
  ├─ ─────────────────────────────
  └─ About
```

#### Patch Options Behind Advanced Dialog
Patch Report and Patch Failure Report widgets removed from the main window. `Help > Advanced` opens a modal `Toplevel` dialog containing those widgets. StringVars persist across opens so file paths are remembered. A one-line status indicator on the main window (`"No patch data (Help ▸ Advanced to configure)"` / `"Patch: filename.csv"`) shows current state at a glance without opening the dialog.

#### Input Validation Added
- `score_var`: `float()` conversion now wrapped in `try/except ValueError` with a user-friendly error message
- `date_var`: `datetime.strptime(..., "%d/%m/%Y")` validation before dispatch — previously a malformed date silently fell back to `1900-01-01`, excluding all devices as stale

#### Progress Bar Stabilised
Progress bar moved into a fixed-height `tk.Frame` with `pack_propagate(False)`. `grid_remove()` hides it between runs without destroying the widget, eliminating layout jitter on subsequent runs.

#### Update CVE Data
`git pull` on the `cvelistV5` repo runs in a daemon background thread. Searches for the repo at the hardcoded default path, then falls back to the script's parent directory. Handles: `git` not on PATH, 120-second timeout, arbitrary exceptions — all surfaced via `root.after()` on the main thread.

---

### v0.10 — CLI Parity Fix

**File:** `run_dashboard.py`

- `--threshold` default changed from `1.0` to `9.0` to match GUI default
- `--report-month` argument added — allows retroactive labelling of runs (`"April 2026"` generated in May). Passed through to `DashboardRequest.report_month`
- `--since` date format updated to `dd/mm/yyyy` to match `dayfirst=True` parsing throughout pipeline

---

### v0.11 — Garbage Cleanup

**File:** `data_pipeline.py`

Removed redundant duplicate `_sr` assignment in `process_patch_match()` that ran before the `_patch_status` rename (result was immediately overwritten by the post-rename assignment).

---

### v0.12 — Memory Footprint Reductions

**Files:** `data_pipeline.py`, `orchestrator.py`, `requirements.txt`

**Problem:** Processing large N-able/RMM exports through pandas was consuming ~1.7 GB of RAM due to two compounding issues: pandas defaulting all text columns to the `object` dtype (a pointer-per-row to heap-allocated Python strings), and a chain of defensive `.copy()` calls in the orchestrator that duplicated the merged DataFrame multiple times in memory.

**Changes:**

`data_pipeline.py` — Added `_downcast_low_cardinality(df, cols)` helper that casts low-cardinality string columns to `category` dtype. Category stores each unique string value once and uses integers for all rows, typically cutting per-column RAM by ~90%. Called from:
- `load_vulnerability_data` — casts `Vulnerability Severity`, `Threat Status`, `Has Known Exploit`, `CISA KEV`
- `load_rmm_data` — casts `Device Type`
- `merge_data` — re-downcasts the merged frame after all conditional `.loc` writes are complete

`data_pipeline.py` — Switched `pd.read_csv` in `load_vulnerability_data` to `dtype_backend='pyarrow'` with a silent fallback to the default backend. PyArrow-backed strings share memory more aggressively than Python-object strings. Requires `pyarrow>=14`.

`data_pipeline.py` — Added explicit decategorise step immediately after `pd.merge()` inside `merge_data`. Categorical columns reject `.loc[mask, col] = value` writes for values not in their category list and raise `TypeError`. `Device Type` and `Last Response` are cast back to `object` right after the merge (before any conditional writes), then re-downcasted to `category` at the end of the function.

`orchestrator.py` — Removed `.copy()` from `filtered_df`, `triage_df`, and `not_in_rmm_df`. Confirmed by grep + audit that `excel_builder.py` only reads from these frames via `.loc[mask, col]` returning `.nunique()` counts — no assignments. Kept `.copy()` on `raw_df` (required, as `merged_df` is mutated by the date filter below it) and `stale_excluded` (detaches from `merged_df` before the filter rebind).

`requirements.txt` — Added `pyarrow>=14` as an optional-but-recommended dependency.

**What was deliberately NOT changed:**

`xlsxwriter` `constant_memory=True` — `build_overview_sheet` writes column 0 down rows r0..r0+3, then jumps back to row r0 to start column 4. `constant_memory` mode flushes rows as the write cursor advances and raises on writes to already-closed rows. Unsafe without refactoring the overview sheet to write strictly top-to-bottom.

`usecols` on `pd.read_csv` — CVE export column names are not fixed; `load_vulnerability_data` handles aliasing across `cvss score` / `cvss v3.1 base score` / `base score` / etc. A static `usecols` list would silently drop columns on exports with different headers. Categorical dtype gives the same memory benefit without the brittleness.

---

### v0.13 — Vectorised Merge: 41s → 1s

**Files:** `data_pipeline.py`, `excel_builder.py`

**Problem:** Total run time was ~40 seconds. Profiling via log timestamps isolated the entire delay to `merge_data` — loading was instant, Excel writing was ~7s. The root cause was two row-by-row `.apply()` loops both calling `parse_last_response()`, which internally calls `pd.to_datetime()` per-row inside a Python `try/except`:

```python
# Before — called twice across 11,161 rows each
merged['_Sort_Time'] = merged['Last Response'].apply(parse_last_response)          # loop 1
merged['Days Since Last Response'] = merged['Last Response'].apply(_calc_days_from_lr)  # loop 2 — also called parse_last_response internally
```

On 11k rows: 22,322 individual Python-level datetime parse attempts. Measured at ~41 seconds.

**Fix — `data_pipeline.py`:** Replaced both `.apply()` loops with a single vectorised pass:

```python
# After — one bulk pd.to_datetime call covers both columns
_sort_time = pd.to_datetime(_lr_str.where(~_sentinel_mask, other=pd.NaT),
                            errors='coerce', format='mixed', dayfirst=False)
merged['_Sort_Time'] = _sort_time.fillna(_epoch)

# Days reuses the already-parsed series — no second parse
_days_num = (_now - merged['_Sort_Time']).dt.days.clip(lower=0).astype(object)
_days_num[_no_data] = '—'
merged['Days Since Last Response'] = _days_num
```

- `format='mixed'` explicitly handles the mix of date formats in N-able exports and silences the `UserWarning: Could not infer format` that appeared when pandas fell back to element-by-element `dateutil` parsing
- Sentinel rows (`Not Found in RMM`, `N/A`, empty) masked out before parsing, set to `'—'` after
- `parse_last_response()` itself is unchanged — still used by other callers outside `merge_data`

**Fix — `excel_builder.py`:** Added `observed=True` to `filtered_df.groupby('Device Type')` call. Silences `FutureWarning: The default of observed=False is deprecated` that pandas emits when grouping on a categorical column without explicitly stating observed behaviour. `observed=True` is correct here — only device types actually present in the filtered data should appear in the count.

**Result:** Merge step 41s → <1s. Total run time 41s → 7s on 11,161 CVE rows / 276 devices with trend comparison enabled.

---

### v0.14 — Resilience & Maintainability Pass

**Files:** `data_pipeline.py`, `diagnostics.py`, `orchestrator.py`

Based on a technical review of the core engine identifying four categories of improvement: row iteration performance, exception specificity, path handling consistency, and global state safety.

#### iterrows() → itertuples() / bulk assignment

`iterrows()` boxes every row into a full pandas Series object — heap allocation, dtype coercion, attribute lookup overhead per row. `itertuples()` returns a lightweight C-level namedtuple with direct attribute access, typically 10–50× faster.

Converted the following loops:

`data_pipeline.py` — `_apply_cascade_resolution` `has_ver` build loop: `iterrows()` → `itertuples(index=False)`. Column names with spaces (`Matched Patch`, `Matched Patch Version`, `Patch Install Date`) accessed via `getattr(row, 'Matched_Patch', '')` (itertuples replaces spaces with underscores).

`data_pipeline.py` — `_apply_cascade_resolution` write-back loop: kept as `iterrows()` because it needs the integer index `idx` to collect into `resolve_indices`. Refactored from `df.at[idx] = value` inside the loop to collecting all indices first, then a single `df.loc[resolve_indices, 'Patch Evidence Status'] = 'Patch confirmed - pending rescan'` bulk assignment after the loop — reduces write overhead and avoids repeated copy-on-write triggers.

`data_pipeline.py` — `load_previous_report` checkbox loop: `iterrows()` → `itertuples(index=False)`. Accesses `row.Name` and `row.Vulnerability_Name`.

`data_pipeline.py` — `compute_patch_diagnostics` (stub): both `lag_rows` and `mismatch_rows` loops converted to `itertuples`. Note: the active version of this function lives in `diagnostics.py`; the stub in `data_pipeline.py` is unused by the orchestrator but kept consistent.

`diagnostics.py` — `build_recommended_actions` loop: `iterrows()` → `itertuples(index=False)`. Accesses `Patch_Evidence_Notes`, `Product`, `Device`.

`diagnostics.py` — `compute_patch_diagnostics` root cause loop: `iterrows()` → `itertuples(index=False)`. All `row.get('Column Name', '')` calls replaced with `getattr(row, 'Column_Name', '')`.

`diagnostics.py` — `compute_patch_diagnostics` lag loop: `iterrows()` → `itertuples(index=False)`.

`orchestrator.py` — `patch_gap_pairs` build loop: `iterrows()` → `itertuples(index=False)`. Accesses `row._nk`, `row._ck`, `row._root_cause` (all underscore-prefixed, so no space substitution needed).

#### Exception specificity

`data_pipeline.py` — `parse_last_response`: replaced three `except Exception: pass` clauses with `except (ValueError, TypeError): pass` (and `except (ValueError, TypeError, AttributeError)` for the digits/timedelta branch). Bare `except Exception` swallows `KeyboardInterrupt`, `MemoryError`, and `SystemExit` — exceptions that should propagate. The new clauses catch only what `pd.to_datetime` and `int()` actually raise on bad input.

The sheet-level `except Exception: continue` in `load_previous_report` is intentionally left broad — it guards against arbitrary `xlsxwriter` parse failures from malformed or encrypted product sheets, where the correct action is always to skip and continue regardless of failure mode.

#### Path handling

`data_pipeline.py` — `load_previous_report`: replaced four `os.path.basename(file_path)` calls with `Path(file_path).name`. Added `from pathlib import Path` to imports. `pathlib` is already used throughout `orchestrator.py`; this brings `data_pipeline.py` into alignment.

#### Global state (noted, not yet refactored)

The review correctly identified `FIXED_VERSION_RULES.clear() / .update()` in `orchestrator._try_sync_baselines` as a global mutation risk under concurrent runs. Full remediation requires threading `rules` as an explicit parameter through `process_patch_match`, `_apply_cascade_resolution`, and `cve_lookup.enrich_from_detections` — a cross-cutting change deferred to a dedicated refactor. The current tool is single-threaded from the GUI, so this is low risk today. Noted here for the next architecture pass.

#### Tkinter thread safety (confirmed, no change needed)

Reviewed `main.py` — all cross-thread UI updates already use `root.after(0, callback)`. The `git pull` background thread writes only to local variables and queues results through `root.after`. No direct widget access from background threads found. Current implementation is correct.

---

### v0.15 — GUI Modernisation: CustomTkinter

**Files:** `main.py`

Replaced the standard `tkinter` widget set with `customtkinter` (CTk 5.2.2). Dark mode enabled by default via `ctk.set_appearance_mode("dark")`.

Every visual widget converted: `ctk.CTkLabel`, `ctk.CTkEntry`, `ctk.CTkButton`, `ctk.CTkCheckBox`, `ctk.CTkFrame`, `ctk.CTkToplevel`, `ctk.CTkProgressBar`, `ctk.CTkScrollableFrame`. Three widget types intentionally kept as plain tkinter: `tk.StringVar` / `tk.BooleanVar` (CTk uses these unchanged), `tk.Menu` (no CTk equivalent — native menu bar is correct), `filedialog` and `messagebox` (system dialogs, no CTk replacement).

Main window content wrapped in `CTkScrollableFrame` — on small screens the old layout clipped the bottom widgets; scroll frame ensures everything is reachable at any window height.

All toggle helpers updated from `tk.NORMAL` / `tk.DISABLED` constants to `"normal"` / `"disabled"` strings (CTk requirement).

`requirements.txt` — added `customtkinter>=5.2`.

---

### v0.16 — Username Column: RMM → All Worksheets

**Files:** `data_pipeline.py`, `excel_builder.py`

`data_pipeline.py` — `load_rmm_data`: added Username column detection across seven aliases (`username`, `user name`, `logged-on user`, `logged on user`, `current user`, `user`, positional header). Standardised to `'Username'`. Missing column filled with empty string so downstream code never guards for absence.

`data_pipeline.py` — `merge_data`: `'Username'` added to `rmm_pull` conditionally (only when present in `df_rmm` and absent from `df_vuln`). Post-merge guard fills any remaining NaN values.

`excel_builder.py` — `Username` column added immediately after `Device Name` in: All Detections, product sheets (`cols_order`), trend detail sheets (`detail_cols`), Resolved (Patch Confirmed), Raw Data, and all stale / not-in-RMM sheets.

---

### v0.17 — Top At-Risk Devices Table

**Files:** `excel_builder.py`

Added a **Top At-Risk Devices** table to the Client Summary sheet, positioned between the Data Filtering waterfall and the CVSS Score Split.

Priority rules (in order): every Server with at least one unresolved CVE always appears regardless of count; every device with a CVE flagged `Has Known Exploit = Yes` always appears; remaining slots filled to a maximum of 10 by highest unresolved CVE count.

Columns: 💻 Device Name, 👤 Username, ⚠️ Unresolved CVEs, 💣 Has Exploit, 🖥️ Device Type.

Row highlighting: amber = server, red-tint = exploit device (exploit takes priority if both apply).

Only unresolved CVEs count — resolved detections excluded from aggregation. Status column detected from either `Threat Status` or `Status` column names.

Aggregation uses `itertuples` on the aggregated frame (no spaces in column names at that stage) — safe from the itertuples space-in-column-name bug.

---

### v0.18 — Worksheet Restructure: 4 Dedicated Stale / Not-in-RMM Sheets

**Files:** `excel_builder.py`, `orchestrator.py`

Replaced the old two-section layout (stale + not-in-RMM rows merged into one sheet with section headers) with four dedicated single-table filterable sheets:

- 🕑 Stale Excluded Devices — devices excluded because Last Response predates the cutoff
- 🚫 Devices Not in RMM — devices absent from the RMM inventory (audit record)
- 🕑 CVEs on Stale Devices — CVE-level detail for stale devices with NVD links
- 🚫 CVEs on Devices Not in RMM — CVE-level detail for not-in-RMM devices with NVD links

Each sheet is a flat table with autofilter on all columns. Not-in-RMM rows highlighted red.

Added two shared helper functions (`_write_device_table`, `_write_cve_table`) and two prep functions (`_device_prep`, `_cve_prep`) that use `iterrows` for the write loop — avoiding the `itertuples` space-in-column-name bug where `'Device Name'` (space) would be silently renamed to `_0` by `itertuples`, causing all device names to write as blank.

`excel_builder.py` — `build_stale_excluded_sheet` and `build_stale_cves_sheet` signatures unchanged so orchestrator required no call-site changes.

`orchestrator.py` — updated sheet name references for trend sheets renamed with emoji prefixes (🆕 New This Month, ⏳ Persisting CVEs).

Added `'Device Name'` header rename and `Username` column to all sheets. Added emoji prefixes to Client Summary section headers (📊 Key Metrics, ✅ Resolution Status, 🔍 Data Filtering, 📈 CVSS Score Split, 📅 Month-over-Month, 🚨 Top At-Risk Devices).

---

### v0.19 — Resolution Status: Consistent Metric Across Both Sheets

**Files:** `excel_builder.py`, `orchestrator.py`

**Problem:** The April Detections overview sheet "Resolution Status" tile showed `Unresolved: −1,682` and reported a fundamentally different metric from the Client Summary, making the two figures incomparable.

**Root cause 1 — negative unresolved:** `patch_confirmed_count` was computed as the size of the intersection of `patch_resolved_pairs` (3-tuples: `device, cve, product`) with `triage_keys` (also 3-tuples). The same `(device, cve)` pair matching multiple products (e.g. Chrome AND Edge both resolving CVE-2024-X) counted as two separate entries, inflating the resolved count above `n_total` (which was a 2-tuple unique pair count) and producing a negative unresolved value.

Fix in `orchestrator.py`: after taking the 3-tuple intersection, deduplicate to `(device, cve)` 2-tuples before counting:
```python
_confirmed_3tuples = patch_resolved_pairs & triage_keys
patch_confirmed_count = len({(d, v) for d, v, _ in _confirmed_3tuples})
```

**Root cause 2 — metric mismatch:** The overview tile used `patch_confirmed_count` (patch tool confirmation) while Client Summary used N-able's `Status` column. These measure different things and will never produce comparable numbers.

Fix in `excel_builder.py`: the overview tile now uses the same source as Client Summary — N-able's `Status` column — counting unique `(device, CVE)` pairs per status. Resolved + Unresolved can exceed Total when the same CVE is resolved on some devices and unresolved on others; an overlap count is shown with an explanatory note.

**Resolution Status footnote** also updated: changed `"296 CVE types resolved vs 108 unresolved"` to `"{rows} detection rows resolved vs {rows} unresolved ({pct}% remediated). {n} CVE type(s) still present on at least one active device."` — the old wording implied full remediation of 296 CVE types when it actually meant 296 appeared in at least one resolved detection row.

---

### v0.20 — CVE Enrichment: Local cvelistV5 Repo Integration

**Files:** `cve_lookup.py`, `orchestrator.py`

#### Local Repo as Source 0

The CVE lookup source order was extended with a new "Source 0" that reads directly from a local clone of the MITRE `cvelistV5` git repository, bypassing all network calls for already-cloned CVEs.

`cve_lookup.py` — `_cve_local_path()`: constructs the expected filesystem path for a CVE JSON 5.0 file using the repo's directory layout (`cves/{year}/{prefix}xxx/CVE-{year}-{num}.json`). `_load_cve_org_local()`: opens the file and returns the parsed dict — identical structure to the CVE.org API response, so `_parse_cve_org()` handles both paths without branching.

Full lookup order in `lookup_fixed_version()`:
1. **Local cvelistV5 repo** — file read, no network round trip
2. **CVE.org API** — only hit when local file is missing or has no version data
3. **NVD API 2.0** — fallback when CVE.org returns nothing useful
4. **OSV.dev** — last resort, especially for open-source packages

Both `enrich_config()` and `enrich_from_detections()` accept a `cve_repo_path` parameter and pass it through to `lookup_fixed_version`. If the path is absent or the file doesn't exist, the function falls back silently — no error, no changed behaviour.

#### CVSS Score Pre-pass

`cve_lookup.py` — `enrich_from_detections()`: added a bulk CVSS pre-pass at the end of each enrichment run. After enriching version data, the function iterates all CVEs in the detection frame that are not yet in `cvss_score_cache`, reads each local JSON file (if the repo is present), and calls `_extract_cvss_score()` to populate the cache. Because each file read is ~1 ms, this pre-pass covers hundreds of CVEs in seconds without any network calls. Results are written back to `config.json` in a single write.

`data_pipeline.py` — `load_vulnerability_data()`: Step 2 of score derivation (lines 500–514) applies the CVSS cache to the loaded DataFrame. When the pre-pass has already populated scores for CVEs that were previously derived from severity labels (CRITICAL/IMPORTANT/MODERATE/LOW), those approximate scores are upgraded to real CVSS floats automatically on the next run.

#### Auto-repo Discovery in Orchestrator

`orchestrator.py` — `_find_cve_repo()`: searches three candidate paths in order:
1. Hardcoded project path (`C:\NoCScripts\...\cvelistV5`)
2. Script's parent directory (`Path(__file__).parent / 'cvelistV5'`)
3. Parent of script's parent (`Path(__file__).parent.parent / 'cvelistV5'`)

Returns the first path that `exists()`, or `None`. `_pull_cve_repo()` runs `git pull --ff-only` on the found repo before each enrichment pass, with a 30-second timeout. Handles: `git` not on PATH, timeout, arbitrary exceptions — all non-fatal, logged at `WARNING` or `DEBUG`.

The repo path is passed to `enrich_from_detections(df_vuln, cve_repo_path=_cve_repo)` so every run automatically uses local data where available.

---

### v0.21 — CVE Enrichment: Edge Chromium Supplement & Auto Product Map

**Files:** `cve_lookup.py`

#### Edge Chromium Version Resolution

Chromium CVEs (filed against Google Chrome by the CNA) previously stored a Chrome fixed version under `chrome` but had no corresponding Edge entry. Microsoft Edge is a Chromium derivative on a separate version train — the Chrome version is incorrect for Edge compliance checks.

`cve_lookup.py` — when `lookup_fixed_version` matches a CVE to `'chrome'` and `'edge'` is present in `product_map`, a supplementary Edge lookup runs:

1. **CVE listed in Microsoft Edge release notes** — queries `edgeupdates.microsoft.com/api/products`, finds the Stable channel, walks releases sorted oldest-first, returns the first release that explicitly lists the CVE ID. Most authoritative source.
2. **Chromium milestone fallback** — extracts the Chrome major version from the Chrome fixed version, finds the earliest Edge Stable release on that Chromium milestone whose release date is on or after the CVE publish date (date guard prevents attributing a pre-publication Edge release as the fix).
3. **No result stored** — if neither method resolves a version, the `edge` key is omitted. No guessing.

The Edge API response is cached module-level (`_EDGE_CACHE`) so it is fetched at most once per run regardless of how many Chromium CVEs are processed.

Chrome and Edge canonical keys always receive independently determined versions — the Chrome version is never copied to `edge`.

#### Auto Product Map

`cve_lookup.py` — `_auto_update_product_map()`: when `auto_add_products=True` (the default), vendor/product pairs from CVE data that have no matching `product_map` entry are automatically registered. Two helper functions drive this:

- `_derive_canonical()`: produces a stable lowercase alphanum+underscore key from the MITRE vendor/product string. Handles known divergences (e.g. `'haxx curl'` → `'curl'`, `'adobe acrobat'` → `'acrobat'`) via a `_KNOWN_CANONICAL` lookup before falling back to generic slug derivation.
- `_derive_search_key()`: produces the substring that will match both the CVE export's `Affected Products` column and the patch report's `Patch` column — strips version numbers, architecture tags, and noise words so the key stays stable across releases.

New entries are inserted before the first `'windows'` catch-all in `product_map` so specificity order is preserved. The canonical key is also seeded into `fixed_version_rules` so subsequent runs can store version data against it.

#### Per-thread Session Pooling

`cve_lookup.py` — replaced the module-level `_SESSION` global with `threading.local()` storage. Each worker thread in the `ThreadPoolExecutor` inside `enrich_config` now owns its own `requests.Session` and connection pool, eliminating pool exhaustion and state leakage under concurrent lookups.

---

### v0.22 — Config Health Check

**File:** `orchestrator.py`

Added `_config_health_check(cfg)` — runs at the start of every `run()` call and appends issues to `DashboardResult.warnings` without blocking execution.

Checks performed:
- **Duplicate `product_map` keys** — the same search string appearing more than once causes non-deterministic product matching (first-match wins, duplicates are dead entries).
- **Orphaned `fixed_version_rules` entries** — a product key in `fixed_version_rules` with no corresponding `product_map` entry means its version rules can never be applied; the pipeline will never route a CVE to that key.
- **Unparseable `_baseline` versions** — a `_baseline` value that fails the `^\d+(?:\.\d+){1,5}$` regex will cause `_version_gte` to return `None` for every comparison, silently classifying all devices as `'Version comparison failed'` instead of compliant or non-compliant.
- **Chrome/Edge version collision** — if the same CVE ID has an identical version string under both `chrome` and `edge` in `fixed_version_rules`, it means the Chrome version was incorrectly copied to Edge. Chrome and Edge are on different version trains and must always differ.

Issues are logged at `WARNING` level and surfaced to the GUI via `DashboardResult.warnings`.

---

### v0.23 — Dual-pass Stale Filter

**File:** `orchestrator.py`

The previous stale filter used a single date cutoff: devices last seen before `cutoff_date` were excluded. This correctly caught devices inactive since before the reporting period but missed devices that had gone offline during the period and were now ≥ 30 days stale by the time the report was generated.

**Fix:** The orchestrator now runs two stale passes before building `stale_excluded`:

- **Pass 1 — date-stale:** last `_Sort_Time` < `cutoff_date` AND `Last Response != 'Not Found in RMM'`. Unchanged from previous behaviour.
- **Pass 2 — days-stale:** `Days Since Last Response >= 30` AND `Last Response != 'Not Found in RMM'` AND device not already caught by Pass 1. Uses `pd.to_numeric(..., errors='coerce')` so sentinel `'—'` values don't raise.

Both passes are concatenated into `stale_excluded`. The union of their device names is removed from `merged_df` in a single filter. The log line reports both counts separately: `%d stale excluded (%d by date-filter, %d by %d-day rule)`.

The stale Excluded Devices sheet `Reason` column distinguishes the two causes: `⏱ Date-Stale` vs `⏱ Days-Stale`.

---

### v0.24 — Raw Scanner Override for Checkbox Integrity

**File:** `orchestrator.py`

**Problem:** The patch tool's `patch_resolved_pairs` set could contain `(device, cve, product)` 3-tuples for CVEs that the N-able scanner still marks UNRESOLVED. When these pairs survived into `build_product_sheets`, the row rendered as a blue ☑ (resolved) even though the scanner disagreed. A stale cache entry, a product-name formatting difference, or a mismatched patch record could all produce false positives.

**Fix:** A "bulletproof scanner override" step runs in the orchestrator immediately after `patch_resolved_pairs` is built, reading directly from `raw_df` (the pre-filter, pre-join source of truth):

- **Step 1 — strip false positives:** builds `_unresolved_pairs_2d` as a set of `(device, cve)` 2-tuples from all rows where the scanner status is `UNRESOLVED`. Any 3-tuple in `patch_resolved_pairs` whose `(device, cve)` appears in `_unresolved_pairs_2d` is removed. Product-string formatting differences can never cause a miss because the comparison is 2-tuple, not 3-tuple.
- **Step 2 — inject raw RESOLVED pairs:** builds `_raw_inject_pairs` from all rows where scanner status is `RESOLVED` and adds them to `patch_resolved_pairs`, skipping any pair where the scanner also shows `UNRESOLVED` on the same device+CVE. This ensures CVEs marked RESOLVED by N-able render as ☑ even without patch tool data.

Rule: **UNRESOLVED always wins.** If the scanner shows UNRESOLVED for a (device, cve), no patch evidence — regardless of source — can override it to blue.

---

### v0.25 — Device Inventory Sheet

**File:** `excel_builder.py`

Added `build_device_report_sheet()` — writes all RMM devices to a `Device Inventory` sheet, one row per device, sorted by `Days Since Last Response` ascending (most recently seen first).

Columns: Device Name, Device Type, OS, Last Response, Days Since Last Response, Username, Site, Serial Number, Manufacturer, Model. Columns are only written when present in `df_rmm` — the function does not raise on missing columns.

Colour coding by recency:
- **Green** (`#E8F5E9`) — seen within 7 days
- **Amber** (`#FFF9C4`) — 8–14 days
- **Red** (`#FFEBEE`) — 15–29 days
- **Dark red** (`#FFCDD2`) — 30+ days (stale threshold)

Called from orchestrator unconditionally when `df_rmm is not None`.

---

### v0.26 — Patch Failure Report Sheet

**Files:** `data_pipeline.py`, `excel_builder.py`, `orchestrator.py`

Added support for an optional N-able Patch Failure Report as a third input file, surfacing devices where patch delivery is actively failing alongside the CVE exposure data.

`data_pipeline.py` — `load_patch_failure_report()`: reads the failure report, normalises device names, and returns a cleaned DataFrame. `build_patch_failure_lookup()`: aggregates by device, counts failures, and identifies the most common failure description per device.

`excel_builder.py` — `build_patch_failure_sheet()`: writes a `Patch Failures` sheet showing all failure records with device, patch name, failure reason, and failure count. Devices that also appear in `triage_df` (i.e. have active CVEs) are highlighted to surface the intersection of "patch failing" and "CVE exposed".

`orchestrator.py` — `DashboardRequest` gains `failure_report_path` and `include_failure_report` fields. When a failure report is loaded, the top 3 devices by failure count are appended to `DashboardResult.warnings`. If any failure devices also carry unresolved CVEs, a second warning is appended with the CVE and device counts and a pointer to the Patch Failures sheet.

`main.py` — Patch Failure Report file picker added to the Advanced dialog alongside the existing Patch Report picker.

---

### v0.27 — Browser Audit Integration

**Files:** `data_pipeline.py`, `orchestrator.py`

Added optional ingestion of a browser audit export (per-device installed browser version inventory) to enrich the version drift diagnostics sheet.

`data_pipeline.py` — `load_browser_audit()`: reads the audit file, auto-detects browser sheets by name (Chrome, Edge, Firefox), normalises column names, and returns a unified DataFrame with a `Browser` column. `merge_browser_audit_into_drift()`: merges per-device audit data into the existing `version_drift_df` from `compute_patch_diagnostics`, adding an `Audit Note` column that flags per-user/AppData installs and 32-bit installs on 64-bit OS. These are common causes of version drift that patch tools miss because they scan system-level installs only.

`orchestrator.py` — `DashboardRequest` gains `browser_audit_path` and `include_browser_audit` fields. When included, `merge_browser_audit_into_drift` is called after `compute_patch_diagnostics` and the result replaces `diagnostics['version_drift_df']`. Load errors are caught and appended to `warnings` rather than failing the run.

---

## Known Architecture Decisions

- **`not_in_rmm_df` excluded from trend arithmetic** — Not-in-RMM rows are excluded from `_active_trend_scope` via the `Last Response != 'Not Found in RMM'` filter regardless of whether an RMM inventory was provided. This is intentional: devices with no confirmed identity in the managed estate must not generate phantom New/Resolved/Persisting signals.

- **`_Checkbox_Resolved` removed from DataFrame** — Manual checkbox state from previous reports is kept as a standalone `set` returned from `load_previous_report` and passed explicitly to `compute_trends`. It is used only for re-detection tracking (identifying CVEs ticked resolved last month that re-appeared). It never touches any DataFrame that feeds `_active_trend_scope`.

- **`Status_p` naming preserved for display** — The patch report's `Status` column is renamed to `_patch_status` before the merge to avoid collision with the CVE `Status` column, then renamed back to `Status_p` before the internal column drop so the Excel output column name users see is unchanged.

- **CVSS cache can overwrite source-file scores** — `load_vulnerability_data` applies the `cvss_score_cache` from `config.json` after loading, upgrading severity-band estimates to real CVSS floats. If the source file already contains real numeric scores, the cache still runs and may overwrite them with cached values from a prior enrichment run. This is intentional for files using severity labels; it is worth noting for files that already carry real scores.

- **`approaching_stale_names` disabled** — The approaching-stale warning (amber highlight for devices within `stale_warning_days` of the staleness threshold) is implemented and wired through to `build_product_sheets` and `build_client_summary_sheet` but is currently commented out in the orchestrator pending further testing. The `stale_warning_days` field on `DashboardRequest` defaults to 14 and is accepted by all downstream functions.

- **Resolved count off-by-one in trend comparison** — A known discrepancy of 1 in the "CVE types resolved" trend metric can occur when a CVE's cached CVSS score changes between the previous and current run. If a CVE was UNRESOLVED in the previous report at one score and is RESOLVED in the current report at a different score, the trend comparison logic may not correctly attribute it as resolved. The counts for New and Persisting CVEs are unaffected. Investigation is deferred; the root cause is in `compute_trends` score re-evaluation against the previous report's raw data.

---


### v0.28 — Per-Sheet Health Subtotals (live score at any workbook size)
**Files:** `product_sheets.py`, `summary_sheet.py`, `sheet_helpers.py`

**Problem:** The Summary sheet's live Patching Health Score embedded one (or two) cross-sheet `COUNTIFS` per product sheet directly into every visible formula. With many product sheets the longest formula reached 23,000+ characters — far past Excel's 8,192-character stored-formula limit — so the score permanently fell back to static values and never updated when ☑/☐ were toggled.

**Fix — total once per sheet, sum cell references:**
- **Per-sheet subtotal block** — every product sheet (full triage *and* Patch Confirmed) now carries six hidden cells at a fixed address (labels col Q, values col R, rows 1–6): resolved ☑, unresolved ☐, resolved/unresolved CVSS ≥ 9, resolved/unresolved known-exploit. Each holds a short *local* `COUNTIF`/`COUNTIFS` over that sheet's own columns, with the generation-time count written as the cached value. Shared writer: `sheet_helpers.write_hs_subtotals` / `hs_subtotal_ref`.
- **Summary fleet-total helper cells** — the six fleet totals (`'Sheet'!$R$1 + …`, one reference per sheet) are written into hidden cells `E5:E10` next to the existing `E4` score helper. Every visible formula (score, grade, coverage rates, points) references `$E$5..$E$10`, so it is constant-size regardless of sheet count. With 70 sheets the longest formula dropped from 23,472 to ~2,600 characters; the 8,192 guard is retained but would now only trip past roughly 200 max-length sheet names, and still degrades gracefully to static values.
- **Resolution Status table converted too** — it built the same giant `COUNTIF` chains and had *no* length guard (a latent corrupt-workbook risk at ~180+ sheets). Its ☑/☐ totals now sum the same per-sheet subtotal cells via two hidden helper cells in col G beside the table, with an 8,192 guard and static fallback. Because this table is live even when the health score is disabled, the subtotal block is written unconditionally.
- **Bonus correctness fix** — the old Summary-side formulas hard-coded `G:G` (Vulnerability Score) and `I:I` (Has Known Exploit) for every sheet, but Patch Confirmed sheets have no Score Lift column, so their columns sit one to the left — the live critical/exploit components were silently testing Risk Severity Index and CISA KEV on those sheets. Subtotal formulas are now built from each sheet's own column order, and sheet names containing apostrophes are escaped in references.

Verified: 88/88 tests pass; a 70-product workbook recalculated in LibreOffice reproduces the Python-computed score and all six fleet totals exactly, and toggling ☐→☑ updates the resolved/unresolved totals live.

---


### v0.29 — Live Grade #NAME? Fix + Score-Box Hardening
**Files:** `summary_sheet.py`

**Problem:** With live health-score formulas now actually being written (v0.28), the grade box showed `#NAME?` in Excel. Root cause: the grade formula used `IFS()`, an Excel 2019+ "future function" — xlsxwriter writes the name verbatim, but Excel stores such functions internally as `_xlfn.IFS`, so a bare `IFS` fails name resolution. The bug predates v0.28; the permanent static-value fallback had masked it.

**Fixes:**
- **`IFS()` → nested `IF()`** — equivalent logic, compatible with every Excel and LibreOffice version.
- **Grade box colour never updated** — the score/grade box conditional-format rules were numeric (`cell between 90,100`), which can never match the grade cell's *text* value ("A".."F"); the grade box stayed neutral dark blue forever. Rules are now formula-based on the live `$E$4` score, so both boxes recolour together.
- **KEV grade caps missing from the live score** — `compute_patching_health_score` caps the score at 74 (KEV ≥ 3, or KEV > 0 with no patch report) / 89 (any KEV), but the live formula only subtracted penalties, so toggling ☑ could push the displayed live score/grade above the documented ceiling. The generation-time cap is now baked into the formula as `MIN(74|89, …)` — consistent with penalties, which are also fixed at generation.
- **Merged formula cells carried cached 0** — `merge_range` cannot store a cached formula result, so readers that don't recalculate on open displayed 0 in the score/grade boxes. Now merges a blank and overwrites the top-left cell via `write_formula` with the correct cached result (the documented xlsxwriter pattern; verified to produce a single cell entry in the sheet XML). The old comment claiming this corrupts the workbook was wrong for this write order.
- **Static-fallback boxes lost their grade colour** — the box format hard-coded neutral blue even in static mode (where no conditional formatting exists to recolour it). Static mode now uses the generation-time grade colour, as the adjacent comment always claimed.

Verified: 88/88 tests pass; a 70-sheet workbook shows identical score/grade values whether read from cached results or fully recalculated in LibreOffice (score 54 / grade D both paths), and a KEV-bearing no-patch-report workbook writes `=MIN(74,…)`.

---


### v0.30 — Live KEV Grade-Cap & Penalty Lift
**Files:** `summary_sheet.py`, `product_sheets.py`, `sheet_helpers.py`

**Problem:** A workbook generated with unresolved CISA KEV CVEs and no patch report scored C (74/100) — correct at generation. But after the reader marked **every** row ☑, the live score still showed 74/C: the v0.29 KEV grade cap was baked in as a constant `MIN(74, …)`, and the −1/CVE KEV penalty was static, so neither could ever lift no matter what the reader resolved.

**Fix — make the cap and KEV penalty conditional on a live unresolved-KEV count:**
- **Seventh per-sheet subtotal** — each product sheet's hidden block (col R) gains `R7`: unresolved rows with CISA KEV = Yes, again using that sheet's own column position for the KEV column.
- **Fleet cell `$E$11`** on Summary sums those refs — a live count of unresolved KEV rows across the workbook.
- **Score formula** — the penalty term becomes `persist_pen + IF($E$11>0, kev_pen, 0)` and the cap becomes `MIN(IF($E$11>0, 74|89, 100), …)`. The cap *tier* (74 vs 89) stays a generation-time constant, but *whether* it applies is live.
- **Exactness note** — the Python penalty counts unique KEV CVE *types*, which no in-sheet formula can deduplicate; but the boundary is exact: unresolved KEV rows = 0 ⇔ unresolved KEV types = 0. So the "everything is ☑" case lifts precisely; intermediate states keep the generation-time penalty (still documented as fixed). If the sheets lack a CISA KEV column, the live cell can't be trusted and the fixed cap/penalty is retained.
- **KEV penalty display row is live too** — shows `=IF($E$11>0, −pts, 0)` so the −1 visibly clears alongside the score, with the label updated to "lifts when all KEV rows are ☑". The footnote now states the cap/penalty lift behaviour instead of claiming all penalties are fixed.

Verified: 88/88 tests pass. A KEV-bearing, no-patch-report workbook recalculates to score 58 / grade D with cap and −2 penalty active while KEV rows remain ☐, and to score 100 / grade A with the penalty at 0 once every row is marked ☑.

---

## Files Changed Per Release

| Release | `main.py` | `orchestrator.py` | `data_pipeline.py` | `excel_builder.py` | `diagnostics.py` | `snapshot.py` | `run_dashboard.py` | `cve_lookup.py` |
|---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| v0.2 | ✓ | ✓ | ✓ | ✓ | — | — | — | — |
| v0.3 | ✓ | ✓ | ✓ | — | — | — | ✓ | — |
| v0.4 | — | ✓ | ✓ | ✓ | ✓ | — | — | — |
| v0.5 | — | ✓ | ✓ | — | — | ✓ | — | — |
| v0.6 | — | ✓ | — | ✓ | — | — | — | — |
| v0.7 | — | ✓ | — | ✓ | — | — | — | — |
| v0.8 | — | — | ✓ | ✓ | — | — | — | — |
| v0.9 | ✓ | — | — | — | — | — | — | — |
| v0.10 | — | — | — | — | — | — | ✓ | — |
| v0.11 | — | — | ✓ | — | — | — | — | — |
| v0.12 | — | ✓ | ✓ | — | — | — | — | — |
| v0.13 | — | — | ✓ | ✓ | — | — | — | — |
| v0.14 | — | ✓ | ✓ | — | ✓ | — | — | — |
| v0.15 | ✓ | — | — | — | — | — | — | — |
| v0.16 | — | — | ✓ | ✓ | — | — | — | — |
| v0.17 | — | — | — | ✓ | — | — | — | — |
| v0.18 | — | ✓ | — | ✓ | — | — | — | — |
| v0.19 | — | ✓ | — | ✓ | — | — | — | — |
| v0.20 | — | ✓ | ✓ | — | — | — | — | ✓ |
| v0.21 | — | — | — | — | — | — | — | ✓ |
| v0.22 | — | ✓ | — | — | — | — | — | — |
| v0.23 | — | ✓ | — | — | — | — | — | — |
| v0.24 | — | ✓ | — | — | — | — | — | — |
| v0.25 | — | ✓ | — | ✓ | — | — | — | — |
| v0.26 | ✓ | ✓ | ✓ | ✓ | — | — | — | — |
| v0.27 | — | ✓ | ✓ | — | — | — | — | — |