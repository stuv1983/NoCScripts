"""
Missing Device Checker - RMMcheck2.py
(Compares an Intune/Entra export against an RMM export to identify corporate
 devices present in Intune but missing from RMM.)

Author: Stuart Villanti <svillanti@kenstra.com>
Date: 2025-08-06

Key Decisions:
- We define a device as "corporate" if it is either:
    • joinType contains "JOINED" (and not "REGISTERED"), or
    • Ownership == "CORPORATE"
- We build lookup sets from RMM by both serial number and a normalized name key,
  then flag any Intune corporate device that matches neither as missing.
- Serial‐number matches take precedence; devices without serials fall back to name.
- Sign-in timestamps are parsed in three passes (precise → ISO → relative) for
  freshness filtering.
"""

import pandas as pd
import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox
import re

# ------------------------------------------------------------------------------
# Helper Functions
# ------------------------------------------------------------------------------

def parse_last_checkin(age_str: str) -> pd.Timestamp:
    """
    Relative "HH:MM.SS" → absolute Timestamp; NaT on error.

    Fix: original code did float('0.' + secs) * 60 which raises ValueError
    when secs contains a decimal point (e.g. "53.8" → "0.53.8").
    Correct approach: treat the whole token as float seconds.
    """
    try:
        hrs, rest = age_str.split(':', 1)
        mins, secs = (rest.split('.', 1) + ['0'])[:2]
        delta = pd.Timedelta(
            hours=int(hrs),
            minutes=int(mins),
            seconds=float(secs)          # ← was: float('0.' + secs) * 60
        )
        return pd.Timestamp.now() - delta
    except (ValueError, AttributeError):
        return pd.NaT

def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Trim headers and map common variants to canonical names:
      - Device/DisplayName      → Device name
      - serial…number           → Serial number
      - Last check-in           → Last check-in
      - any joinType/trustType  → joinType
    """
    df.rename(columns=lambda c: c.strip(), inplace=True)
    mapping = {}
    for col in df.columns:
        low = col.lower()
        if low in ('device', 'displayname'):
            mapping[col] = 'Device name'
        elif 'serial' in low and 'number' in low:
            mapping[col] = 'Serial number'
        elif low == 'last check-in':
            mapping[col] = 'Last check-in'
        elif 'jointype' in low or 'trusttype' in low:
            mapping[col] = 'joinType'
    return df.rename(columns=mapping)

def normalize(df: pd.DataFrame) -> pd.DataFrame:
    """Uppercase & trim key fields; replace whole-cell 'NAN' strings with empty."""
    df = standardize_columns(df)
    for f in ('Device name', 'Serial number', 'Ownership', 'operatingSystem', 'joinType'):
        if f in df.columns:
            df[f] = (
                df[f].astype(str)
                     .str.strip()
                     .str.upper()
                     .replace({'NAN': ''})
            )
    return df

def normalize_name(name: str) -> str:
    """Make a simple key: uppercase, strip LAPTOP-/DESKTOP-, drop non-alphanum."""
    nm = str(name).upper()
    for p in ('LAPTOP-', 'DESKTOP-'):
        if nm.startswith(p):
            nm = nm[len(p):]
    return re.sub(r'[^A-Z0-9]', '', nm).strip()

def read_file(path: str) -> pd.DataFrame:
    """
    Load a CSV or Excel file into a DataFrame.
    Accepts .csv, .xlsx, and .xls; raises ValueError for anything else.
    For multi-sheet workbooks the first sheet is used.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == '.csv':
        return pd.read_csv(path)
    elif ext in ('.xlsx', '.xls'):
        return pd.read_excel(path, sheet_name=0)
    else:
        raise ValueError(f"Unsupported file type '{ext}'. Please select a .csv or .xlsx file.")

def style_sheet(ws, header_color: str = "1F4E79"):
    """Apply consistent table styling to an openpyxl worksheet."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    thin = Side(style='thin', color='CCCCCC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        for cell in row:
            cell.font = Font(name='Arial', size=10)
            cell.border = border
            if row_idx == 1:
                cell.font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', start_color=header_color)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(vertical='center')

    # Auto-fit column widths (cap at 60)
    for col_idx, col in enumerate(ws.columns, start=1):
        max_len = max((len(str(cell.value or '')) for cell in col), default=8)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 60)

    ws.freeze_panes = 'A2'

# ------------------------------------------------------------------------------
# Core Logic
# ------------------------------------------------------------------------------

def find_missing_devices(rmm_file: str, intune_file: str, output_file: str):
    """
    1.  Load & normalize RMM export.
    2.  Extract serial numbers from Description if blank.
    3.  Load & normalize Intune/Entra export.
    4.  Filter to corporate devices (joinType='JOINED' OR Ownership='CORPORATE').
    5.  Parse LastSignInDT (precise → ISO → relative).
    6.  Drop stale (<2025-01-01) and non-Windows devices.
    7.  De-duplicate by Device name.
    8.  Split into missing / matched sets.
    9.  Write four-sheet Excel workbook:
          • Missing from RMM
          • Matched in RMM & Intune
          • Raw Intune data
          • Raw RMM data
    """
    try:
        from openpyxl import load_workbook

        # ── Step 1: RMM import & clean ─────────────────────────────────────
        rmm_raw = read_file(rmm_file)
        rmm = normalize(rmm_raw.copy())
        rmm['Device name'] = rmm['Device name'].str.strip()

        if 'Serial number' not in rmm.columns:
            rmm['Serial number'] = ''
        if 'Description' in rmm.columns:
            blank = rmm['Serial number'].str.strip() == ''
            extracted = (
                rmm.loc[blank, 'Description']
                   .str.extract(r'(?:S/N|SN)[:\s]*([A-Za-z0-9\-]+)', expand=False)
                   .fillna('').str.upper().str.strip()
            )
            rmm.loc[blank, 'Serial number'] = extracted

        serials = {s for s in rmm['Serial number'].dropna().str.strip().str.upper() if s}
        rmm['NormalizedName'] = rmm['Device name'].apply(normalize_name)
        names = set(rmm['NormalizedName'])

        # ── Step 2: Intune/Entra import & clean ───────────────────────────
        intune_raw = read_file(intune_file)
        df = normalize(intune_raw.copy())

        if 'Serial number' not in df.columns:
            df['Serial number'] = ''
        if 'Ownership' not in df.columns:
            df['Ownership'] = 'UNKNOWN'

        # ── Step 3: Corporate filter ───────────────────────────────────────
        joined_mask = False
        if 'joinType' in df.columns:
            joined_mask = (
                df['joinType'].str.contains('JOINED', na=False) &
                ~df['joinType'].str.contains('REGISTERED', na=False)
            )
        corp = df[(joined_mask) | (df['Ownership'] == 'CORPORATE')].copy()

        # ── Step 4: Parse LastSignInDT ─────────────────────────────────────
        if 'approximateLastSignInDateTime' in corp.columns:
            corp['LastSignInDT'] = pd.to_datetime(
                corp['approximateLastSignInDateTime'], errors='coerce'
            )

        if 'Last check-in' in corp.columns:
            iso = pd.to_datetime(corp['Last check-in'], errors='coerce')
            corp['LastSignInDT'] = iso
            m = corp['LastSignInDT'].isna() & corp['Last check-in'].notna()
            corp.loc[m, 'LastSignInDT'] = corp.loc[m, 'Last check-in'].apply(parse_last_checkin)

        if 'LastSignInDT' not in corp.columns:
            corp['LastSignInDT'] = pd.NaT

        # ── Step 5: Drop stale & non-Windows ──────────────────────────────
        cutoff = pd.Timestamp(2025, 1, 1)
        corp = corp[corp['LastSignInDT'] >= cutoff]
        if 'operatingSystem' in corp.columns:
            corp = corp[corp['operatingSystem'].str.contains('WINDOWS', na=False)]

        # ── Step 6: De-duplicate ───────────────────────────────────────────
        if 'Device name' in corp.columns:
            corp.drop_duplicates(subset='Device name', keep='last', inplace=True)

        corp['NormalizedName'] = corp['Device name'].apply(normalize_name)

        # ── Step 7: Split missing / matched ───────────────────────────────
        present_by_sn = corp['Serial number'].str.strip().str.upper().isin(serials)
        present_by_nm = corp['NormalizedName'].isin(names)
        present_mask  = present_by_sn | present_by_nm

        missing = corp[~present_mask].copy()
        matched = corp[present_mask].copy()

        # Columns to expose in result sheets
        result_cols = [
            'Device ID', 'Device name', 'Primary user UPN', 'operatingSystem',
            'Last check-in', 'approximateLastSignInDateTime', 'LastSignInDT',
            'Serial number', 'Ownership', 'joinType'
        ]

        def trim(frame: pd.DataFrame) -> pd.DataFrame:
            cols = [c for c in result_cols if c in frame.columns]
            return frame[cols]

        # ── Step 8: Write Excel workbook ───────────────────────────────────
        os.makedirs(os.path.dirname(output_file) or '.', exist_ok=True)

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            trim(missing).to_excel(writer, sheet_name='Missing from RMM',          index=False)
            trim(matched).to_excel(writer, sheet_name='Matched in RMM & Intune',   index=False)
            intune_raw.to_excel(   writer, sheet_name='Raw Intune data',            index=False)
            rmm_raw.to_excel(      writer, sheet_name='Raw RMM data',               index=False)

        # Apply styling post-write
        wb = load_workbook(output_file)
        header_colors = {
            'Missing from RMM':        'C00000',   # red
            'Matched in RMM & Intune': '375623',   # green
            'Raw Intune data':         '1F4E79',   # navy
            'Raw RMM data':            '1F4E79',   # navy
        }
        for sheet_name, color in header_colors.items():
            if sheet_name in wb.sheetnames:
                style_sheet(wb[sheet_name], header_color=color)
        wb.save(output_file)

        msg = (
            f'{len(missing)} device(s) missing from RMM.\n'
            f'{len(matched)} device(s) matched.\n\n'
            f'Output: {output_file}'
        ) if not missing.empty else (
            f'All corporate Intune devices are present in RMM.\n'
            f'{len(matched)} matched device(s).\n\n'
            f'Output: {output_file}'
        )
        messagebox.showinfo('Check Complete', msg)

    except Exception as e:
        traceback.print_exc()
        messagebox.showerror('Error', f'{type(e).__name__}: {e}')

# ------------------------------------------------------------------------------
# GUI Launcher
# ------------------------------------------------------------------------------

def launch_gui():
    """
    Simple Tk form:
      - Select RMM CSV
      - Select Intune/Entra CSV
      - Select output .xlsx
      - Run comparison
    """
    root = tk.Tk()
    root.title('Intune → RMM Missing Devices')

    rmm_var    = tk.StringVar()
    intune_var = tk.StringVar()
    out_var    = tk.StringVar()

    # RMM file (CSV or Excel)
    tk.Label(root, text='RMM File:').grid(row=0, column=0, sticky='e', padx=6, pady=4)
    tk.Entry(root, textvariable=rmm_var, width=55).grid(row=0, column=1, pady=4)
    tk.Button(root, text='Browse…',
              command=lambda: rmm_var.set(
                  filedialog.askopenfilename(
                      filetypes=[("CSV or Excel", "*.csv *.xlsx *.xls"),
                                 ("CSV", "*.csv"),
                                 ("Excel", "*.xlsx *.xls")]
                  )
              )).grid(row=0, column=2, padx=4)

    # Intune file (CSV or Excel)
    tk.Label(root, text='Intune/Entra File:').grid(row=1, column=0, sticky='e', padx=6, pady=4)
    tk.Entry(root, textvariable=intune_var, width=55).grid(row=1, column=1, pady=4)
    tk.Button(root, text='Browse…',
              command=lambda: intune_var.set(
                  filedialog.askopenfilename(
                      filetypes=[("CSV or Excel", "*.csv *.xlsx *.xls"),
                                 ("CSV", "*.csv"),
                                 ("Excel", "*.xlsx *.xls")]
                  )
              )).grid(row=1, column=2, padx=4)

    # Output XLSX  ← was CSV
    tk.Label(root, text='Output Excel:').grid(row=2, column=0, sticky='e', padx=6, pady=4)
    tk.Entry(root, textvariable=out_var, width=55).grid(row=2, column=1, pady=4)
    tk.Button(root, text='Save As…',
              command=lambda: out_var.set(
                  filedialog.asksaveasfilename(
                      defaultextension='.xlsx',
                      filetypes=[("Excel Workbook", "*.xlsx")]
                  )
              )).grid(row=2, column=2, padx=4)

    # Run
    tk.Button(root, text='Run Check',
              command=lambda: find_missing_devices(
                  rmm_var.get(), intune_var.get(), out_var.get()
              )).grid(row=3, column=1, pady=12)

    root.mainloop()


if __name__ == '__main__':
    launch_gui()