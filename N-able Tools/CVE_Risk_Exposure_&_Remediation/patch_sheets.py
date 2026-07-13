"""
patch_sheets.py — sheets derived from the patch report / diagnostics
pipeline: Patch Evidence Notes, Patch Lag, Version Drift, Resolved (Patch
Confirmed), Products Not in Patch Scope, Patch Failures, CVEs on Failing
Devices, and the raw Patch Match Overview/Full/Report sheets.

Split out of excel_builder.py as part of breaking that file into one
module per sheet category.

Note: build_patch_resolved_sheet is currently unused — orchestrator.py
imports it but the call sites are commented out ("large sheet, slow to
write"). Moved as-is rather than dropped, since removing it wasn't part
of this refactor's scope.

Author : Stu Villanti <s.villanti@kenstra.com>
"""
import pandas as pd
from typing import Optional

def build_diagnostics_sheets(writer, diagnostics: dict) -> None:
    wb = writer.book
    red  = wb.add_format({'bg_color': '#FCE4D6'})
    amb  = wb.add_format({'bg_color': '#FFF2CC'})
    grn  = wb.add_format({'bg_color': '#E2EFDA'})
    note = wb.add_format({'italic': True, 'font_color': '#595959'})

    _LABEL_COLOUR = {
        'Patch required':                '#FCE4D6',
        'Installed but still detected':  '#FCE4D6',
        'No patch evidence':             '#FFF2CC',
        'Product not tracked':           '#FFF2CC',
        'No patch baseline defined':     '#FFF2CC',
        'Installed but version unknown': '#BDD7EE',
    }

    rc_df = diagnostics.get('root_cause_df', pd.DataFrame())
    if not rc_df.empty:
        _SHOW_COLS = ['Device', 'Product', 'CVE', 'Patch Match Result',
                      'Resolved', 'Patch Evidence Notes', 'Baseline Compliance',
                      'Recommended Steps']
        out = rc_df[[c for c in _SHOW_COLS if c in rc_df.columns]].copy()
        out.to_excel(writer, sheet_name='Patch Evidence Notes', index=False)
        ws = writer.sheets['Patch Evidence Notes']
        ws.autofilter(0, 0, len(out), len(out.columns) - 1)
        ws.set_column('A:A', 28)
        ws.set_column('B:B', 30)
        ws.set_column('C:C', 20)
        ws.set_column('D:D', 35)
        ws.set_column('E:E', 12)
        ws.set_column('F:F', 32)
        ws.set_column('G:G', 22)
        ws.set_column('H:H', 55)
        for i, label in enumerate(out.get('Patch Evidence Notes', []), start=1):
            colour = _LABEL_COLOUR.get(str(label), '#FFFFFF')
            ws.set_row(i, 30, wb.add_format({'bg_color': colour, 'text_wrap': True, 'valign': 'top'}))
        ws.write(len(out) + 2, 0,
                 'Patch Evidence Notes indicate likely follow-up areas based on CVE and '
                 'patch report correlation — not confirmed root cause.', note)

    lag_df = diagnostics.get('patch_lag_df', pd.DataFrame())
    if not lag_df.empty:
        lag_df.to_excel(writer, sheet_name='Patch Lag', index=False)
        ws = writer.sheets['Patch Lag']
        ws.autofilter(0, 0, len(lag_df), len(lag_df.columns) - 1)
        ws.set_column('A:A', 28); ws.set_column('B:B', 18); ws.set_column('C:C', 32)
        ws.set_column('F:F', 12)
        # Find the Lag column index for conditional_format
        _lag_cols = list(lag_df.columns)
        _lag_idx  = _lag_cols.index('Lag (days)') if 'Lag (days)' in _lag_cols else None
        _n = len(lag_df)
        if _lag_idx is not None:
            _lag_letter = chr(ord('A') + _lag_idx)
            # Green: 0–14 days — patched promptly
            ws.conditional_format(1, 0, _n, len(_lag_cols) - 1, {
                'type': 'formula', 'criteria': f'=AND(${_lag_letter}2>=0,${_lag_letter}2<=14)',
                'format': grn,
            })
            # Amber: 15–60 days — acceptable but slow
            ws.conditional_format(1, 0, _n, len(_lag_cols) - 1, {
                'type': 'formula', 'criteria': f'=AND(${_lag_letter}2>14,${_lag_letter}2<=60)',
                'format': amb,
            })
            # Red: >60 days or negative (patch pre-dates detection — data anomaly)
            ws.conditional_format(1, 0, _n, len(_lag_cols) - 1, {
                'type': 'formula', 'criteria': f'=OR(${_lag_letter}2>60,${_lag_letter}2<0)',
                'format': red,
            })
        ws.write(_n + 2, 0,
                 'Negative lag = patch installed before CVE was first detected.', note)

    drift_df = diagnostics.get('version_drift_df', pd.DataFrame())
    if not drift_df.empty:
        drift_df.to_excel(writer, sheet_name='Version Drift', index=False)
        ws = writer.sheets['Version Drift']
        ws.autofilter(0, 0, len(drift_df), len(drift_df.columns) - 1)
        ws.set_column('A:A', 36); ws.set_column('C:C', 60)
        # Audit Note column (may or may not be present)
        _cols = list(drift_df.columns)
        if 'Audit Note' in _cols:
            _an_idx = _cols.index('Audit Note')
            ws.set_column(_an_idx, _an_idx, 50)
        # Colour by Distinct Versions count using conditional_format
        _dv_cols = list(drift_df.columns)
        _dv_idx  = _dv_cols.index('Distinct Versions') if 'Distinct Versions' in _dv_cols else None
        _dn = len(drift_df)
        if _dv_idx is not None:
            _dv_letter = chr(ord('A') + _dv_idx)
            ws.conditional_format(1, 0, _dn, len(_dv_cols) - 1, {
                'type': 'formula', 'criteria': f'=${_dv_letter}2=1',
                'format': grn,
            })
            ws.conditional_format(1, 0, _dn, len(_dv_cols) - 1, {
                'type': 'formula', 'criteria': f'=AND(${_dv_letter}2>=2,${_dv_letter}2<4)',
                'format': amb,
            })
            ws.conditional_format(1, 0, _dn, len(_dv_cols) - 1, {
                'type': 'formula', 'criteria': f'=${_dv_letter}2>=4',
                'format': red,
            })
        ws.write(len(drift_df) + 2, 0,
                 'High distinct-version count = inconsistent update cadence across fleet. '
                 'Audit Note: per-user/AppData installs bypass GPO — remove and replace with system-scope. '
                 '32-bit installs on 64-bit OS should be replaced.',
                 wb.add_format({'italic': True, 'font_color': '#595959', 'text_wrap': True}))
        ws.set_row(len(drift_df) + 2, 36)
        no_data = diagnostics.get('version_drift_no_data', [])
        if no_data:
            ws.write(len(drift_df) + 4, 0,
                     f'ℹ  No version data for: {", ".join(no_data)} — '
                     f'these products are not returning version numbers from the patch tool. '
                     f'Version drift cannot be assessed until they are tracked.',
                     wb.add_format({'italic': True, 'font_color': '#7F6000', 'text_wrap': True}))
    else:
        ws = writer.book.add_worksheet('Version Drift')
        no_data = diagnostics.get('version_drift_no_data', [])
        if no_data:
            ws.write(0, 0,
                     f'No version data available for: {", ".join(no_data)}. '
                     f'These products are detected by N-able but the patch tool is not '
                     f'returning installed version numbers — they may not be in your patch '
                     f'policy scope. Version drift cannot be assessed.',
                     wb.add_format({'italic': True, 'font_color': '#7F6000', 'text_wrap': True}))
            ws.set_column('A:A', 80)
            ws.set_row(0, 50)


def build_patch_resolved_sheet(writer, patch_full_df: 'pd.DataFrame') -> None:
    import pandas as pd

    resolved = patch_full_df[
        patch_full_df['Patch Evidence Status'] == 'Patch confirmed - pending rescan'
    ].copy()

    if resolved.empty:
        return

    wb       = writer.book
    grn      = wb.add_format({'bg_color': '#E2EFDA'})
    hdr      = wb.add_format({'bold': True, 'bg_color': '#375623',
                               'font_color': 'white', 'border': 1})
    note_fmt = wb.add_format({'italic': True, 'font_color': '#595959'})

    if 'Patch Install Date' in resolved.columns and 'First detected' in resolved.columns:
        idt = pd.to_datetime(resolved['Patch Install Date'], errors='coerce')
        fdt = pd.to_datetime(resolved['First detected'],    errors='coerce')
        resolved['Lag (days)'] = (idt - fdt).dt.days
    else:
        resolved['Lag (days)'] = ''

    cols = [c for c in [
        'Name', 'Vulnerability Name', 'Affected Products',
        'Vulnerability Score', 'Matched Patch Version',
        'Patch Install Date', 'First detected', 'Lag (days)',
        'Product Baseline', 'Baseline Compliance',
    ] if c in resolved.columns]

    out = (resolved[cols]
           .drop_duplicates(subset=['Name', 'Vulnerability Name'])
           .sort_values(['Affected Products', 'Vulnerability Score'],
                        ascending=[True, False])
           .reset_index(drop=True))

    out.to_excel(writer, sheet_name='Resolved (Patch Confirmed)', index=False)
    ws = writer.sheets['Resolved (Patch Confirmed)']
    ws.autofilter(0, 0, len(out), len(out.columns) - 1)
    ws.set_row(0, None, hdr)

    ws.set_column('A:A', 28)
    ws.set_column('B:B', 22)
    ws.set_column('C:C', 32)
    ws.set_column('D:D', 10)
    ws.set_column('E:E', 22)
    ws.set_column('F:G', 20)
    ws.set_column('H:H', 12)
    ws.set_column('I:I', 20)
    ws.set_column('J:J', 22)

    # Replace per-row set_row loop with a single conditional_format rule.
    # set_row() called N times = N xlsxwriter calls; one conditional_format
    # rule = one XML element regardless of row count — O(1) vs O(n).
    if len(out):
        ws.conditional_format(1, 0, len(out), len(out.columns) - 1, {
            'type':     'formula',
            'criteria': '=$A2<>""',
            'format':   grn,
        })

    note_row = len(out) + 2
    unique_cves     = out['Vulnerability Name'].nunique()
    unique_devices  = out['Name'].nunique()
    ws.write(note_row, 0,
             f'{unique_cves} CVE type(s) resolved across {unique_devices} device(s) '
             f'via patch report. Install date confirmed after first detection date.',
             note_fmt)
    ws.merge_range(note_row, 0, note_row, len(out.columns) - 1,
                   f'{unique_cves} CVE type(s) confirmed patched across {unique_devices} '
                   f'device(s) via patch report. Install date confirmed after first detection date.',
                   note_fmt)


def build_products_not_tracked_sheet(writer,
                                      patch_full_df: 'pd.DataFrame') -> None:
    import re
    from data_pipeline import get_base_product, _detect_product, _norm_text

    wb       = writer.book
    red      = wb.add_format({'bg_color': '#FCE4D6'})
    amb      = wb.add_format({'bg_color': '#FFF2CC'})
    hdr      = wb.add_format({'bold': True, 'bg_color': '#1F4E79',
                               'font_color': 'white', 'border': 1})
    code_fmt = wb.add_format({'font_name': 'Courier New', 'font_size': 9,
                               'bg_color': '#F2F2F2'})
    note_fmt = wb.add_format({'italic': True, 'font_color': '#595959',
                               'text_wrap': True})

    unmanaged = patch_full_df[
        patch_full_df['Patch Match Result'] == 'Device in patch report - product not found'
    ].copy()

    if unmanaged.empty:
        return

    unmanaged['_bp'] = unmanaged['Affected Products'].apply(get_base_product)
    unmanaged['_pk'] = unmanaged['Affected Products'].apply(
        lambda v: _detect_product(_norm_text(str(v))))

    agg = (unmanaged.groupby(['_bp', '_pk'])
           .agg(
               devices       = ('Name',               'nunique'),
               cves          = ('Vulnerability Name', 'nunique'),
               sample_names  = ('Name', lambda x: ', '.join(sorted(x.unique())[:3])
                                         + (' ...' if x.nunique() > 3 else '')),
           )
           .reset_index()
           .sort_values('devices', ascending=False)
           .reset_index(drop=True))

    def _suggest_entry(bp, pk):
        bp_clean = re.sub(r'\s+\d[\d.]+\s*$', '', str(bp).lower().strip())
        key      = pk if pk else bp_clean.replace(' ', '_')
        return f'["{bp_clean}", "{key}"]'

    agg['In product_map'] = agg['_pk'].apply(lambda v: '✓' if v else '✗')
    agg['Suggested config.json entry'] = agg.apply(
        lambda r: _suggest_entry(r['_bp'], r['_pk']), axis=1)

    out = agg.rename(columns={
        '_bp':          'Product (as detected by N-able)',
        '_pk':          'Internal Key',
        'devices':      'Devices Affected',
        'cves':         'CVE Count',
        'sample_names': 'Sample Devices',
    })[['Product (as detected by N-able)', 'Devices Affected', 'CVE Count',
        'Sample Devices', 'In product_map', 'Suggested config.json entry']]

    out.to_excel(writer, sheet_name='Products Not in Patch Scope', index=False)
    ws = writer.sheets['Products Not in Patch Scope']
    ws.autofilter(0, 0, len(out), len(out.columns) - 1)
    ws.set_row(0, None, hdr)

    ws.set_column('A:A', 40)
    ws.set_column('B:B', 16)
    ws.set_column('C:C', 11)
    ws.set_column('D:D', 45)
    ws.set_column('E:E', 14)
    ws.set_column('F:F', 45)

    for i, row in enumerate(out.itertuples(), start=1):
        n = row._2
        ws.set_row(i, None, red if n >= 10 else amb)
        ws.write(i, 5, row._6, code_fmt)

    note_row = len(out) + 2
    ws.merge_range(note_row, 0, note_row, 5,
                   'These products are detected by N-able on devices that ARE in the patch report, '
                   'but this specific product is not included in the RMM patch policy for those devices. '
                   'To fix: add the product to your RMM patch policy scope. '
                   'If the product is also missing from config.json (✗ in "In product_map"), '
                   'add the suggested entry to config.json product_map as well.',
                   note_fmt)
    ws.set_row(note_row, 50)


def build_patch_failure_sheet(writer, failure_df: 'pd.DataFrame',
                              failure_lookup: dict,
                              cve_device_overlap: 'pd.DataFrame',
                              inventory_devices: 'set | None' = None) -> None:
    import pandas as pd
    wb  = writer.book
    red = wb.add_format({'bg_color': '#FCE4D6'})
    amb = wb.add_format({'bg_color': '#FFF2CC'})
    grn = wb.add_format({'bg_color': '#E2EFDA'})
    hdr = wb.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1})
    hdr_red  = wb.add_format({'bold': True, 'bg_color': '#C00000', 'font_color': 'white', 'border': 1})
    note_fmt = wb.add_format({'italic': True, 'font_color': '#595959'})
    title_fmt= wb.add_format({'bold': True, 'font_size': 12, 'bg_color': '#1F4E79',
                               'font_color': 'white', 'border': 1})

    active_lookup = failure_lookup
    excluded_count = 0
    if inventory_devices:
        active_lookup  = {d: info for d, info in failure_lookup.items()
                          if d in inventory_devices}
        excluded_count = len(failure_lookup) - len(active_lookup)

    rows = []
    for device, info in sorted(active_lookup.items(),
                               key=lambda x: -x[1]['failure_count']):
        rows.append({
            'Device':               device,
            'Total Failures':       info['failure_count'],
            'Unique KBs Failing':   info['unique_kbs'],
            'Primary Failure Type': info['top_category'].replace('_', ' ').title(),
            'Description':          info['top_description'],
            'All Categories':       ', '.join(f"{k.replace('_',' ').title()}: {v}"
                                             for k, v in info['categories'].items()),
        })
    if not rows:
        return

    summary_df = pd.DataFrame(rows)

    # ── Summary stats at top before the table ───────────────────────────────
    ws = wb.add_worksheet('Patch Failures')
    stat_fmt  = wb.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1})
    stat_val  = wb.add_format({'border': 1, 'align': 'right'})

    total_failures  = sum(r['Total Failures']     for r in rows)
    total_devices   = len(rows)
    total_kbs       = sum(r['Unique KBs Failing'] for r in rows)
    active_fail_df  = failure_df[failure_df['_device_norm'].isin(set(active_lookup.keys()))]
    cat_totals      = active_fail_df['_failure_cat'].value_counts()
    top_cat_label   = cat_totals.index[0].replace('_', ' ').title() if not cat_totals.empty else '—'
    top_cat_count   = int(cat_totals.iloc[0]) if not cat_totals.empty else 0

    ws.merge_range(0, 0, 0, 5, 'Patch Failure Analysis', title_fmt)
    stats = [
        ('Devices with failures', total_devices),
        ('Total failure events',  total_failures),
        ('Distinct KBs failing',  total_kbs),
        ('Most common failure',   f'{top_cat_label} ({top_cat_count} events)'),
    ]
    if excluded_count:
        stats.append(('Excluded (not in inventory)', excluded_count))
    for si, (label, val) in enumerate(stats):
        ws.write(1 + si, 0, label, stat_fmt)
        ws.write(1 + si, 1, val,   stat_val)
    ws.set_column('A:A', 26); ws.set_column('D:D', 30)
    ws.set_column('E:E', 55); ws.set_column('F:F', 55)
    ws.set_column('B:C', 18)

    # ── Device failure table ─────────────────────────────────────────────────
    tbl_start = len(stats) + 3
    for ci, col in enumerate(summary_df.columns):
        ws.write(tbl_start, ci, col, hdr)
    for i, row in enumerate(rows, start=tbl_start + 1):
        fc = row['Total Failures']
        ws.set_row(i, None, red if fc >= 20 else amb if fc >= 5 else grn)
        for ci, col in enumerate(summary_df.columns):
            ws.write(i, ci, row[col])
    ws.autofilter(tbl_start, 0, tbl_start + len(rows), len(summary_df.columns) - 1)

    # ── Category totals below table ──────────────────────────────────────────
    note_start = tbl_start + len(rows) + 2
    ws.write(note_start, 0, 'Failure category totals (active devices):', hdr)
    for i, (cat, count) in enumerate(cat_totals.items()):
        ws.write(note_start + 1 + i, 0, f'  {cat.replace("_"," ").title()}')
        ws.write(note_start + 1 + i, 1, count)

    # ── CVEs on Failing Devices — enriched ──────────────────────────────────
    if not cve_device_overlap.empty:
        # Add primary failure type to each row from lookup
        _fail_info = {d: info for d, info in active_lookup.items()}
        _norm_name = cve_device_overlap['Name'].astype(str).apply(
            lambda n: n.strip().upper().split('\\')[-1].split('.')[0])

        cve_out = cve_device_overlap.copy()
        cve_out['_nk'] = _norm_name
        cve_out['Primary Failure Type'] = cve_out['_nk'].map(
            lambda nk: _fail_info[nk]['top_category'].replace('_', ' ').title()
                       if nk in _fail_info else '—'
        )
        cve_out['Total Device Failures'] = cve_out['_nk'].map(
            lambda nk: _fail_info[nk]['failure_count'] if nk in _fail_info else 0
        )
        cve_out['Failure Description'] = cve_out['_nk'].map(
            lambda nk: _fail_info[nk]['top_description'] if nk in _fail_info else '—'
        )
        cve_out = cve_out.drop(columns=['_nk'], errors='ignore')

        out_cols = [c for c in [
            'Name', 'Vulnerability Name', 'Vulnerability Score', 'Affected Products',
            'Has Known Exploit', 'Primary Failure Type', 'Total Device Failures',
            'Failure Description'
        ] if c in cve_out.columns]

        overlap = (cve_out[out_cols]
                   .drop_duplicates(subset=['Name', 'Vulnerability Name'])
                   .sort_values(['Total Device Failures', 'Vulnerability Score'],
                                ascending=[False, False])
                   .reset_index(drop=True))

        overlap.to_excel(writer, sheet_name='CVEs on Failing Devices', index=False)
        ws2 = writer.sheets['CVEs on Failing Devices']
        ws2.autofilter(0, 0, len(overlap), len(overlap.columns) - 1)
        ws2.set_column('A:A', 26); ws2.set_column('B:B', 22)
        ws2.set_column('D:D', 32); ws2.set_column('F:F', 24)
        ws2.set_column('G:G', 20); ws2.set_column('H:H', 55)
        ws2.set_row(0, None, hdr_red)

        # Colour rows by failure severity
        for i, row in enumerate(overlap.itertuples(index=False), start=1):
            fc = getattr(row, 'Total_Device_Failures', 0) or 0
            ws2.set_row(i, None, red if fc >= 20 else amb if fc >= 5 else grn)

        ws2.write(len(overlap) + 2, 0,
                  f'⚠  {len(overlap)} CVE detection(s) on {overlap["Name"].nunique()} device(s) '
                  f'where patches are actively failing. '
                  f'Resolving the delivery failure (Primary Failure Type) will unblock patching. '
                  f'See Patch Failures sheet for per-device remediation steps.', note_fmt)
        ws2.set_row(len(overlap) + 2, 50)

def build_patch_sheets(writer, overview_df, full_df, patch_df):
    for df, name in ((overview_df, 'Patch Match Overview'),
                     (full_df,     'Patch Match Full Data'),
                     (patch_df,    'Patch Report (Full)')):
        df.to_excel(writer, sheet_name=name, index=False)
        ws = writer.sheets[name]
        ws.autofilter(0, 0, len(df), len(df.columns) - 1)


def build_patch_check_failure_sheet(writer, check_df: 'pd.DataFrame',
                                    check_lookup: dict,
                                    cve_device_overlap: Optional['pd.DataFrame'] = None,
                                    inventory_devices: 'set | None' = None) -> None:
    """
    Builds the 'Patch Check Failures' sheet from an RMM monitoring-check
    export (e.g. N-able's 'Failing Checks' report) — distinct from the
    per-KB 'Patch Failures' sheet above.

    This tracks devices where the RMM agent's own automated Patch Status
    Check is failing to report at all — i.e. RMM can't confirm whether the
    device is patched, which is a leading indicator of a device silently
    falling out of patch management (broken agent, WMI/service issue,
    permissions, etc.) rather than a specific patch failing to install.

    inventory_devices, if provided (normalised device names, e.g. from
    df_rmm['Device_Join']), splits the device list into Active (currently
    in the RMM/CVE dataset) vs Not currently tracked, so the reader can
    tell "this device is live and its patch status is a blind spot" apart
    from "this device may no longer exist."
    """
    import pandas as pd
    wb  = writer.book
    red = wb.add_format({'bg_color': '#FCE4D6'})
    amb = wb.add_format({'bg_color': '#FFF2CC'})
    hdr = wb.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1})
    hdr_red  = wb.add_format({'bold': True, 'bg_color': '#C00000', 'font_color': 'white', 'border': 1})
    note_fmt = wb.add_format({'italic': True, 'font_color': '#595959'})
    title_fmt= wb.add_format({'bold': True, 'font_size': 12, 'bg_color': '#1F4E79',
                               'font_color': 'white', 'border': 1})
    stat_fmt = wb.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1})
    stat_val = wb.add_format({'border': 1, 'align': 'right'})
    date_fmt = wb.add_format({'border': 1, 'num_format': 'yyyy-mm-dd hh:mm'})
    td       = wb.add_format({'border': 1})
    td_r     = wb.add_format({'border': 1, 'align': 'right', 'num_format': '#,##0'})

    if not check_lookup:
        return

    active_devices = set(inventory_devices) if inventory_devices else None

    rows = []
    for device, info in sorted(check_lookup.items(),
                               key=lambda x: (x[1]['days_since'] if x[1]['days_since'] is not None else -1),
                               reverse=True):
        is_active = (device in active_devices) if active_devices is not None else None
        rows.append({
            'Device':             info.get('asset_name', device),
            'Asset Type':         info.get('asset_type', ''),
            'Customer':           info.get('customer', ''),
            'Site':               info.get('site', ''),
            'Check Description':  info.get('check_description', ''),
            'Check Frequency':    info.get('check_frequency', ''),
            'Last Failure':       info.get('last_failure'),
            'Days Since':         info.get('days_since'),
            'Failure Events':     info.get('failure_count', 1),
            'Active':             ('Yes' if is_active else 'No') if is_active is not None else 'Unknown',
        })

    ws = wb.add_worksheet('Patch Check Failures')

    total_devices  = len(rows)
    active_count   = sum(1 for r in rows if r['Active'] == 'Yes')
    unknown_count  = sum(1 for r in rows if r['Active'] == 'Unknown')
    long_silent    = sum(1 for r in rows if isinstance(r['Days Since'], int) and r['Days Since'] >= 30)

    ws.merge_range(0, 0, 0, 5, 'Patch Status Check Failures', title_fmt)
    stats = [
        ('Devices with failing patch status check', total_devices),
        ('Of which currently active (in RMM/CVE data)', active_count) if active_devices is not None
            else ('Devices (active status unknown \u2014 no RMM data supplied)', unknown_count),
        ('Silent \u2265 30 days since last failure recorded', long_silent),
    ]
    for si, (label, val) in enumerate(stats):
        ws.write(1 + si, 0, label, stat_fmt)
        ws.write(1 + si, 1, val,   stat_val)

    ws.set_column('A:A', 26); ws.set_column('B:B', 14)
    ws.set_column('C:C', 20); ws.set_column('D:D', 26)
    ws.set_column('E:E', 40); ws.set_column('F:F', 20)
    ws.set_column('G:G', 20); ws.set_column('H:H', 12)
    ws.set_column('I:I', 16); ws.set_column('J:J', 10)

    tbl_start = len(stats) + 3
    cols = ['Device', 'Asset Type', 'Customer', 'Site', 'Check Description',
            'Check Frequency', 'Last Failure', 'Days Since', 'Failure Events', 'Active']
    for ci, col in enumerate(cols):
        ws.write(tbl_start, ci, col, hdr)
    for i, row in enumerate(rows, start=tbl_start + 1):
        days = row['Days Since']
        row_fmt = red if (isinstance(days, int) and days >= 30) else (amb if row['Active'] == 'Yes' else None)
        if row_fmt is not None:
            ws.set_row(i, None, row_fmt)
        for ci, col in enumerate(cols):
            val = row[col]
            if col == 'Last Failure' and pd.notna(val):
                ws.write_datetime(i, ci, val, date_fmt)
            elif col in ('Days Since', 'Failure Events'):
                ws.write(i, ci, val if val is not None else '', td_r)
            else:
                ws.write(i, ci, str(val) if val is not None else '', td)
    ws.autofilter(tbl_start, 0, tbl_start + len(rows), len(cols) - 1)

    note_row = tbl_start + len(rows) + 2
    ws.merge_range(note_row, 0, note_row, 9,
        '\u2139  \U0001f7e5 Red = failing \u2265 30 days.  \U0001f7e8 Amber = active device, failing < 30 days.  '
        'A device here may still have unresolved CVEs that never get confirmed as patched simply because '
        'RMM cannot report a fresh patch status for it \u2014 check its CVE detections directly rather than '
        'trusting "unresolved" at face value. Fixing the underlying check (agent restart, WMI repair, '
        'permissions) is usually the real remediation, not the individual CVE.',
        note_fmt)
    ws.set_row(note_row, 40)

    # ── CVEs on devices where patch status can't be confirmed ────────────────
    if cve_device_overlap is not None and not cve_device_overlap.empty:
        _fail_info = check_lookup
        _norm_name = cve_device_overlap['Name'].astype(str).apply(
            lambda n: n.strip().upper().split('\\')[-1].split('.')[0])

        cve_out = cve_device_overlap.copy()
        cve_out['_nk'] = _norm_name
        cve_out['Days Since Check Failed'] = cve_out['_nk'].map(
            lambda nk: _fail_info[nk]['days_since'] if nk in _fail_info else None
        )
        cve_out = cve_out.drop(columns=['_nk'], errors='ignore')

        out_cols = [c for c in [
            'Name', 'Vulnerability Name', 'Vulnerability Score', 'Affected Products',
            'Has Known Exploit', 'Days Since Check Failed'
        ] if c in cve_out.columns]

        overlap = (cve_out[out_cols]
                   .drop_duplicates(subset=['Name', 'Vulnerability Name'])
                   .sort_values('Vulnerability Score', ascending=False)
                   .reset_index(drop=True))

        overlap.to_excel(writer, sheet_name='CVEs on Check-Failing Devices', index=False)
        ws2 = writer.sheets['CVEs on Check-Failing Devices']
        ws2.autofilter(0, 0, len(overlap), len(overlap.columns) - 1)
        ws2.set_column('A:A', 26); ws2.set_column('B:B', 22)
        ws2.set_column('D:D', 32); ws2.set_column('F:F', 22)
        ws2.set_row(0, None, hdr_red)
        ws2.write(len(overlap) + 2, 0,
                  f'\u26a0  {len(overlap)} CVE detection(s) on {overlap["Name"].nunique()} device(s) where the '
                  f'patch status check itself is failing \u2014 these CVEs may be stuck showing unresolved '
                  f'purely because RMM cannot confirm a fresh patch status, not because they are genuinely '
                  f'still unpatched.', note_fmt)
        ws2.set_row(len(overlap) + 2, 50)