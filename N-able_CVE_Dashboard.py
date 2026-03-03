# ==========================================
# N-ABLE CVE REPORT MERGER & DASHBOARD UTILITY
# ==========================================

import pandas as pd                  
import tkinter as tk
from tkinter import ttk              
from tkinter import filedialog, messagebox 
from tkcalendar import DateEntry      # <-- NEW IMPORT
import re                            

# ==========================================
# --- GENERAL HELPER FUNCTIONS ---
# ==========================================

def select_file(label_var):
    """Opens a Windows file dialog to select a CSV or Excel file."""
    file_path = filedialog.askopenfilename(
        filetypes=[("Data Files", "*.csv *.xlsx *.xls"), ("CSV Files", "*.csv"), ("Excel Files", "*.xlsx *.xls")]
    )
    if file_path:
        label_var.set(file_path)

def load_data(file_path):
    """Loads data cleanly into pandas based on the file extension."""
    if file_path.lower().endswith(('.xlsx', '.xls')):
        return pd.read_excel(file_path)
    else:
        return pd.read_csv(file_path)

def normalize_device_name(name):
    name = str(name).strip().upper()
    if '\\' in name: name = name.split('\\')[-1] 
    if '.' in name: name = name.split('.')[0]   
    return name

def get_base_product(prod_name):
    p = str(prod_name).strip()
    p = re.sub(r'\bx64\b', '', p, flags=re.IGNORECASE)
    p = re.sub(r'\bx86\b', '', p, flags=re.IGNORECASE)
    p = re.sub(r'\b32-bit\b', '', p, flags=re.IGNORECASE)
    p = re.sub(r'\b64-bit\b', '', p, flags=re.IGNORECASE)
    p = re.sub(r'\s+v?\d+(\.\d+)*\s*$', '', p)
    return p.strip()

def clean_sheet_name(name, used_names):
    if pd.isna(name) or str(name).strip() == "": name = "Unknown Product"
    invalid_chars = r'[\[\]\:\*\?\/\\\'\000]'
    clean_name = re.sub(invalid_chars, '', str(name)).strip()
    clean_name = clean_name[:31].strip() 
    if not clean_name: clean_name = "Unknown Product"
        
    final_name = clean_name
    counter = 1
    while final_name.lower() in [n.lower() for n in used_names]:
        suffix = f"_{counter}"
        final_name = clean_name[:31 - len(suffix)] + suffix
        counter += 1
        
    used_names.add(final_name)
    return final_name

def extract_nvd_link(row):
    cve_pattern = r'(CVE-\d{4}-\d{4,7})'  
    for col in ['CVE', 'CVE ID', 'Vulnerability Name', 'Name']:
        if col in row.index and pd.notna(row[col]):
            cell_value = str(row[col]) if not isinstance(row[col], str) else row[col]
            match = re.search(cve_pattern, cell_value, re.IGNORECASE)
            if match:
                cve_id = match.group(1).upper()
                return f'=HYPERLINK("https://nvd.nist.gov/vuln/detail/{cve_id}", "View")'
    return ""

def make_cve_org_link(val):
    val_str = str(val) if not isinstance(val, str) else val
    if pd.isna(val) or val_str.strip() == "" or val_str.lower() == 'nan': return val
    
    cve_pattern = r'(CVE-\d{4}-\d{4,7})'
    match = re.search(cve_pattern, val_str, re.IGNORECASE)
    if match:
        cve_id = match.group(1).upper()
        display_text = val_str.replace('"', '""')
        if len(display_text) > 250: display_text = display_text[:247] + "..."
        return f'=HYPERLINK("https://www.cve.org/CVERecord?id={cve_id}", "{display_text}")'
    return val

def determine_device_type(os_string):
    val = str(os_string).lower()
    if val == 'nan' or val == 'unknown': return 'Unknown'
    if 'server' in val: return 'Server'
    return 'Workstation'

def parse_last_response(val):
    val = str(val).strip()
    epoch = pd.to_datetime('1900-01-01') 
    if val in ["Not Found in RMM", "N/A", ""]: return epoch
    try: return pd.to_datetime(val)
    except: pass
    
    if val.startswith("overdue_"):
        try:
            clean_val = val.replace("overdue_", "").split(" -")[0]
            return pd.to_datetime(clean_val)
        except: pass
            
    if "days" in val or "hrs" in val:
        try:
            days = 0
            match = re.search(r'(\d+)\s*days', val)
            if match: days = int(match.group(1))
            return pd.Timestamp.now() - pd.Timedelta(days=days)
        except: pass
            
    return epoch

def get_col_letter(col_idx):
    letter = ''
    col_idx += 1
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter

def toggle_rmm_state():
    if skip_rmm_var.get():
        rmm_entry.config(state=tk.DISABLED)
        rmm_button.config(state=tk.DISABLED)
    else:
        rmm_entry.config(state=tk.NORMAL)
        rmm_button.config(state=tk.NORMAL)

def toggle_date_state():
    """Enables or disables the calendar widget based on the checkbox."""
    if show_all_dates_var.get():
        cal.config(state='disabled')
    else:
        cal.config(state='normal')

# ==========================================
# --- DATA PIPELINE FUNCTIONS ---
# ==========================================

def load_vulnerability_data(file_path):
    df_vuln = load_data(file_path)
    
    vuln_rename_dict = {}
    for col in df_vuln.columns:
        c_lower = str(col).strip().lower()
        if c_lower in ['asset name', 'device name', 'endpoint']: vuln_rename_dict[col] = 'Name'
        elif c_lower in ['vulnerability id', 'cve id', 'cve']: vuln_rename_dict[col] = 'Vulnerability Name'
        elif c_lower in ['cvss score', 'cvss v3.1 base score', 'cvss v3 base score', 'base score', 'score']: vuln_rename_dict[col] = 'Vulnerability Score'
        elif c_lower in ['affected products', 'product']: vuln_rename_dict[col] = 'Affected Products'
        elif c_lower in ['severity', 'risk']: vuln_rename_dict[col] = 'Vulnerability Severity'
        elif c_lower in ['threat status', 'status']: vuln_rename_dict[col] = 'Threat Status'
            
    df_vuln.rename(columns=vuln_rename_dict, inplace=True)

    if 'Threat Status' in df_vuln.columns:
        df_vuln = df_vuln[df_vuln['Threat Status'].astype(str).str.strip().str.upper() != 'RESOLVED']
    
    if 'Name' not in df_vuln.columns: df_vuln['Name'] = 'Unknown Device'
    if 'Vulnerability Name' not in df_vuln.columns: df_vuln['Vulnerability Name'] = 'Unknown CVE'
    if 'Affected Products' not in df_vuln.columns: df_vuln['Affected Products'] = 'Unknown Product'
    if 'Vulnerability Score' not in df_vuln.columns: df_vuln['Vulnerability Score'] = 0.0
    if 'Vulnerability Severity' not in df_vuln.columns: df_vuln['Vulnerability Severity'] = 'Unknown'
    
    if 'Has Known Exploit' not in df_vuln.columns: df_vuln['Has Known Exploit'] = 'No'
    if 'CISA KEV' not in df_vuln.columns: df_vuln['CISA KEV'] = 'No'
    if 'Risk Severity Index' not in df_vuln.columns: df_vuln['Risk Severity Index'] = 'Unknown'
    
    df_vuln['Vulnerability Name'] = df_vuln['Vulnerability Name'].fillna('Unknown CVE')
    df_vuln['Name_Join'] = df_vuln['Name'].apply(normalize_device_name)
    df_vuln['Affected Products'] = df_vuln['Affected Products'].fillna('Unknown Product')
    df_vuln['Base Product'] = df_vuln['Affected Products'].apply(get_base_product)
    
    return df_vuln

def load_rmm_data(file_path):
    df_rmm = load_data(file_path)
    col_lower = {c.lower(): c for c in df_rmm.columns}
    
    dev_col, resp_col, os_col = None, None, None
    
    if 'device name' in col_lower: dev_col = col_lower['device name']
    elif 'device' in col_lower: dev_col = col_lower['device']
    elif 'name' in col_lower: dev_col = col_lower['name']
    elif 'asset name' in col_lower: dev_col = col_lower['asset name']
    elif 'hostname' in col_lower: dev_col = col_lower['hostname']
        
    if 'last response (local time)' in col_lower: resp_col = col_lower['last response (local time)']
    elif 'last response (utc)' in col_lower: resp_col = col_lower['last response (utc)']
    elif 'last response' in col_lower: resp_col = col_lower['last response']
    elif 'last check-in' in col_lower: resp_col = col_lower['last check-in']
    
    if 'os version' in col_lower: os_col = col_lower['os version']
    elif 'os' in col_lower: os_col = col_lower['os']
        
    if not dev_col or not resp_col:
        if len(df_rmm.columns) == 9:
            df_rmm.columns = ['Type', 'Client', 'Site', 'Device', 'Description', 'OS', 'Username', 'Last Response', 'Last Boot']
            dev_col = 'Device'
            resp_col = 'Last Response'
            os_col = 'OS'
        else:
            raise ValueError("Could not identify 'Device name' and 'Last response' columns in RMM data.")
    
    df_rmm.rename(columns={dev_col: 'Device', resp_col: 'Last Response'}, inplace=True)
    df_rmm['Device_Join'] = df_rmm['Device'].apply(normalize_device_name)
    df_rmm['Device Type'] = df_rmm[os_col].apply(determine_device_type) if os_col else 'Unknown'
    
    return df_rmm.drop_duplicates(subset=['Device_Join'], keep='first')

def merge_data(df_vuln, df_rmm, skip_rmm):
    if not skip_rmm and df_rmm is not None:
        merged_df = pd.merge(df_vuln, df_rmm[['Device_Join', 'Last Response', 'Device Type']], left_on='Name_Join', right_on='Device_Join', how='left')
        merged_df['Last Response'] = merged_df['Last Response'].fillna("Not Found in RMM")
        merged_df['Device Type'] = merged_df['Device Type'].fillna("Unknown")
    else:
        merged_df = df_vuln.copy()
        merged_df['Last Response'] = "N/A"
        merged_df['Device Type'] = merged_df['Operating System Role'].str.title() if 'Operating System Role' in merged_df.columns else "Unknown"

    merged_df['Vulnerability Score'] = pd.to_numeric(merged_df['Vulnerability Score'], errors='coerce')
    merged_df['_Sort_Time'] = merged_df['Last Response'].apply(parse_last_response)
    return merged_df

# ==========================================
# --- EXCEL GENERATION FUNCTIONS ---
# ==========================================

def build_overview_sheet(workbook, merged_df, filtered_for_sheets_df, threshold, product_to_sheet, header_format, link_format):
    """Builds the main dashboard and metrics view cleanly without charts."""
    overview_sheet = workbook.add_worksheet('Overview')
    
    # --- CALCULATING EXECUTIVE METRICS ---
    is_kev = filtered_for_sheets_df['CISA KEV'].astype(str).str.strip().str.lower().isin(['yes', 'true', '1', 'y'])
    is_exploit = filtered_for_sheets_df['Has Known Exploit'].astype(str).str.strip().str.lower().isin(['yes', 'true', '1', 'y'])

    kev_cves = filtered_for_sheets_df[is_kev]['Vulnerability Name'].nunique()
    kev_devices = filtered_for_sheets_df[is_kev]['Name'].nunique()
    exploit_cves = filtered_for_sheets_df[is_exploit]['Vulnerability Name'].nunique()

    total_detections = len(filtered_for_sheets_df)
    unique_devices = filtered_for_sheets_df['Name'].nunique()
    avg_per_device = round(total_detections / unique_devices, 1) if unique_devices > 0 else 0

    total_servers = merged_df[merged_df['Device Type'] == 'Server']['Name'].nunique()
    servers_affected = filtered_for_sheets_df[filtered_for_sheets_df['Device Type'] == 'Server']['Name'].nunique()
    server_impact_pct = f"{round((servers_affected / total_servers) * 100, 1)}%" if total_servers > 0 else "0%"

    now = pd.Timestamp.now()
    valid_dates_mask = filtered_for_sheets_df['_Sort_Time'] > pd.to_datetime('1950-01-01')
    stale_7 = filtered_for_sheets_df[valid_dates_mask & (filtered_for_sheets_df['_Sort_Time'] < now - pd.Timedelta(days=7))]['Name'].nunique()
    stale_30 = filtered_for_sheets_df[valid_dates_mask & (filtered_for_sheets_df['_Sort_Time'] < now - pd.Timedelta(days=30))]['Name'].nunique()

    # ==========================================
    # 1. EXECUTIVE BANNER (Top Row Horizontal)
    # ==========================================
    overview_sheet.write('A1', 'Exploitability Risk', header_format)
    overview_sheet.write('A2', 'KEV CVEs'); overview_sheet.write('B2', kev_cves)
    overview_sheet.write('A3', 'Devices w/ KEV'); overview_sheet.write('B3', kev_devices)
    overview_sheet.write('A4', 'Known Exploits'); overview_sheet.write('B4', exploit_cves)

    overview_sheet.write('E1', f'Exposure Density (Score {threshold}+)', header_format)
    overview_sheet.write('E2', 'Total Detections'); overview_sheet.write('F2', total_detections)
    overview_sheet.write('E3', 'Unique Devices'); overview_sheet.write('F3', unique_devices)
    overview_sheet.write('E4', 'Avg per Device'); overview_sheet.write('F4', avg_per_device)
    overview_sheet.write('E5', 'Servers Impacted'); overview_sheet.write('F5', f"{servers_affected} ({server_impact_pct})")

    overview_sheet.write('I1', f'Stale Devices (Score {threshold}+)', header_format)
    overview_sheet.write('I2', '> 7 Days Offline'); overview_sheet.write('J2', stale_7)
    overview_sheet.write('I3', '> 30 Days Offline'); overview_sheet.write('J3', stale_30)

    # ==========================================
    # 2. DATA TABLES (Two-Column Layout)
    # ==========================================
    row_tables = 7 
    
    # --- LEFT COLUMN (Col A) ---
    overview_sheet.write(row_tables, 0, 'Unique CVEs by Severity', header_format)
    unique_cves_df = merged_df.drop_duplicates(subset=['Vulnerability Name']).copy()
    sev_counts = unique_cves_df['Vulnerability Severity'].value_counts()
    
    r_sev = row_tables + 1
    for sev, count in sev_counts.items():
        overview_sheet.write(r_sev, 0, str(sev)); overview_sheet.write(r_sev, 1, count)
        r_sev += 1

    row_prod = max(r_sev + 2, 14)
    overview_sheet.write(row_prod, 0, f'Top 10 Products (Score {threshold}+)', header_format)
    prod_counts = filtered_for_sheets_df.groupby('Base Product')['Name'].nunique().sort_values(ascending=False).head(10)
    
    p_idx = row_prod + 1
    for prod, count in prod_counts.items():
        if prod in product_to_sheet:
            target = product_to_sheet[prod]
            overview_sheet.write_url(p_idx, 0, f"internal:'{target}'!A1", string=str(prod), cell_format=link_format)
        else:
            overview_sheet.write(p_idx, 0, str(prod))
        overview_sheet.write(p_idx, 1, count)
        p_idx += 1

    missing_row = p_idx + 2
    overview_sheet.write(missing_row, 0, f"Devices Not Found in RMM (Score {threshold}+)", header_format)
    missing_devices = filtered_for_sheets_df[filtered_for_sheets_df['Last Response'] == "Not Found in RMM"]['Name'].unique()
    
    m_idx = missing_row + 1
    if len(missing_devices) == 0:
        overview_sheet.write(m_idx, 0, "All devices synced")
    else:
        for dev in sorted(missing_devices):
            overview_sheet.write(m_idx, 0, str(dev))
            m_idx += 1


    # --- RIGHT COLUMN (Col E) ---
    overview_sheet.write(row_tables, 4, f'Devices by Type (Score {threshold}+)', header_format)
    dt_counts = filtered_for_sheets_df.groupby('Device Type')['Name'].nunique()
    
    r_dt = row_tables + 1
    for dt, count in dt_counts.items():
        overview_sheet.write(r_dt, 4, str(dt)); overview_sheet.write(r_dt, 5, count)
        r_dt += 1

    row_res = max(r_dt + 2, 14)
    overview_sheet.write(row_res, 4, f'Resolution Status (Score {threshold}+)', header_format)
    
    if product_to_sheet:
        formula_resolved = " + ".join([f"COUNTIF('{s}'!A:A, \"☑\")" for s in product_to_sheet.values()])
        formula_unresolved = " + ".join([f"COUNTIF('{s}'!A:A, \"☐\")" for s in product_to_sheet.values()])
    else:
        formula_resolved, formula_unresolved = "0", "0"

    overview_sheet.write(row_res + 1, 4, "Resolved")
    overview_sheet.write_formula(row_res + 1, 5, f"={formula_resolved}")

    overview_sheet.write(row_res + 2, 4, "Unresolved")
    overview_sheet.write_formula(row_res + 2, 5, f"={formula_unresolved}")

    # ==========================================
    # CLEANUP & FORMATTING
    # ==========================================
    overview_sheet.set_column('A:A', 38)
    overview_sheet.set_column('E:E', 35)
    overview_sheet.set_column('I:I', 20)

def build_all_detections_sheet(writer, merged_df, link_format, missing_row_format):
    merged_df_export = merged_df.copy()
    for col in ['Name_Join', 'Device_Join', 'Base Product']:
        if col in merged_df_export.columns: merged_df_export.drop(columns=[col], inplace=True)
    
    merged_df_export['NVD'] = merged_df_export.apply(extract_nvd_link, axis=1)
    merged_df_export['Vulnerability Name'] = merged_df_export['Vulnerability Name'].apply(make_cve_org_link)
    
    cols = merged_df_export.columns.tolist()
    if 'Device Type' in cols and 'Name' in cols:
        cols.insert(cols.index('Name') + 1, cols.pop(cols.index('Device Type')))
        merged_df_export = merged_df_export[cols]

    merged_df_export = merged_df_export.sort_values(by=['Vulnerability Score', '_Sort_Time', 'Name'], ascending=[False, False, True])
    if '_Sort_Time' in merged_df_export.columns: merged_df_export.drop(columns=['_Sort_Time'], inplace=True)

    merged_df_export.to_excel(writer, sheet_name='All Detections', index=False)
    ws_all = writer.sheets['All Detections']
    ws_all.autofilter(0, 0, len(merged_df_export), len(merged_df_export.columns) - 1)
    
    cols_export_list = merged_df_export.columns.tolist()
    if 'Vulnerability Name' in cols_export_list:
        vn_idx = cols_export_list.index('Vulnerability Name')
        ws_all.set_column(vn_idx, vn_idx, 25, link_format)
    if 'NVD' in cols_export_list:
        nvd_idx = cols_export_list.index('NVD')
        ws_all.set_column(nvd_idx, nvd_idx, 10, link_format)
    if 'Name' in cols_export_list:
        ws_all.set_column(cols_export_list.index('Name'), cols_export_list.index('Name'), 25)
    
    if 'Last Response' in cols_export_list:
        lr_idx = cols_export_list.index('Last Response')
        lr_col_letter = get_col_letter(lr_idx)
        ws_all.conditional_format(1, 0, len(merged_df_export), len(cols_export_list) - 1, {
            'type': 'formula',
            'criteria': f'=${lr_col_letter}2="Not Found in RMM"',
            'format': missing_row_format
        })

def build_product_sheets(writer, filtered_for_sheets_df, product_to_sheet, link_format, missing_row_format):
    cols_order = ['Resolved', 'Vulnerability Name', 'Name', 'Device Type', 'Vulnerability Severity', 'Vulnerability Score', 'Risk Severity Index', 'Has Known Exploit', 'CISA KEV', 'Last Response', 'Affected Products', 'NVD']

    for product, group in filtered_for_sheets_df.groupby('Base Product'):
        sheet_name = product_to_sheet[product]
        group = group.drop_duplicates(subset=['Name', 'Vulnerability Name']).copy()
        group = group.sort_values(by=['Vulnerability Score', '_Sort_Time', 'Name'], ascending=[False, False, True])
        
        group.insert(0, 'Resolved', '☐')
        group['NVD'] = group.apply(extract_nvd_link, axis=1)
        group['Vulnerability Name'] = group['Vulnerability Name'].apply(make_cve_org_link)
        
        final_cols = [c for c in cols_order if c in group.columns]
        group[final_cols].to_excel(writer, sheet_name=sheet_name, index=False)
        
        ws_p = writer.sheets[sheet_name]
        ws_p.autofilter(0, 0, len(group), len(final_cols) - 1)
        
        if 'Resolved' in final_cols:
            res_idx = final_cols.index('Resolved')
            ws_p.data_validation(1, res_idx, len(group), res_idx, {'validate': 'list', 'source': ['☐', '☑']})
            ws_p.set_column(res_idx, res_idx, 10) 

        if 'Vulnerability Name' in final_cols:
            ws_p.set_column(final_cols.index('Vulnerability Name'), final_cols.index('Vulnerability Name'), 25, link_format)
        if 'NVD' in final_cols:
            ws_p.set_column(final_cols.index('NVD'), final_cols.index('NVD'), 10, link_format)
        if 'Name' in final_cols:
            ws_p.set_column(final_cols.index('Name'), final_cols.index('Name'), 25)
        if 'Device Type' in final_cols:
            ws_p.set_column(final_cols.index('Device Type'), final_cols.index('Device Type'), 15)
        
        if 'Last Response' in final_cols:
            lr_idx = final_cols.index('Last Response')
            lr_col_letter = get_col_letter(lr_idx)
            ws_p.conditional_format(1, 0, len(group), len(final_cols) - 1, {
                'type': 'formula',
                'criteria': f'=${lr_col_letter}2="Not Found in RMM"',
                'format': missing_row_format
            })

# ==========================================
# --- ORCHESTRATOR / UI CONTROLLER ---
# ==========================================

def process_reports():
    vuln_path = vuln_var.get()
    rmm_path = rmm_var.get()
    skip_rmm = skip_rmm_var.get()
    
    if not vuln_path:
        messagebox.showerror("Error", "Please select the Vulnerability Report.")
        return
    if not skip_rmm and not rmm_path:
        messagebox.showerror("Error", "Please select the Device Inventory / RMM Report.")
        return

    progress = ttk.Progressbar(root, mode='indeterminate')
    progress.pack(pady=10)
    progress.start()
    root.update()

    try:
        threshold = float(score_var.get())
        
        # 1. Pipeline: Load & Merge Data
        df_vuln = load_vulnerability_data(vuln_path)
        df_rmm = None if skip_rmm else load_rmm_data(rmm_path)
        merged_df = merge_data(df_vuln, df_rmm, skip_rmm)

        # --- NEW: Filter by Calendar Date ---
        if not show_all_dates_var.get():
            try:
                cutoff_date = pd.to_datetime(date_var.get())
                merged_df = merged_df[merged_df['_Sort_Time'] >= cutoff_date]
            except Exception as e:
                progress.stop()
                progress.destroy()
                messagebox.showerror("Error", f"Invalid date format from calendar: {e}")
                return

        if merged_df.empty:
            progress.stop()
            progress.destroy()
            messagebox.showwarning("No Data", "No vulnerability records found matching the specified date filter and thresholds.")
            return

        # 2. Get Output Destination
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not output_file: 
            progress.stop()
            progress.destroy()
            return 

        # 3. Setup Excel Environment
        filtered_for_sheets_df = merged_df[merged_df['Vulnerability Score'] >= threshold].copy()
        used_sheet_names = set(['overview', 'all detections'])
        product_to_sheet = {}
        for product, _ in filtered_for_sheets_df.groupby('Base Product'):
            product_to_sheet[product] = clean_sheet_name(product, used_sheet_names)

        # 4. Write Excel File
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            link_format = workbook.add_format({'font_color': 'blue', 'underline': True})
            header_format = workbook.add_format({'bold': True, 'font_size': 12, 'bg_color': '#D9D9D9', 'border': 1})
            missing_row_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

            build_overview_sheet(workbook, merged_df, filtered_for_sheets_df, threshold, product_to_sheet, header_format, link_format)
            build_all_detections_sheet(writer, merged_df, link_format, missing_row_format)
            build_product_sheets(writer, filtered_for_sheets_df, product_to_sheet, link_format, missing_row_format)

        progress.stop()
        progress.destroy()
        messagebox.showinfo("Success", f"Full Dashboard saved to:\n{output_file}")
        
    except Exception as e:
        progress.stop()
        progress.destroy()
        messagebox.showerror("Error", f"Processing failed: {e}")

# ==========================================
# --- GUI SETUP ---
# ==========================================
root = tk.Tk()
root.title("N-able CVE Dashboard & Triage Tool")
# Increased height from 440 to 520 to fit the calendar layout
root.geometry("540x520")

vuln_var = tk.StringVar()
rmm_var = tk.StringVar()
score_var = tk.StringVar(value="9.0")
skip_rmm_var = tk.BooleanVar(value=False)

tk.Label(root, text="1. Vulnerability Report (CSV or XLSX)", font=('Arial', 10, 'bold')).pack(pady=5)
tk.Entry(root, textvariable=vuln_var, width=65).pack()
tk.Button(root, text="Browse", command=lambda: select_file(vuln_var)).pack()

tk.Label(root, text="2. Device Inventory / RMM Report (CSV or XLSX)", font=('Arial', 10, 'bold')).pack(pady=5)
rmm_entry = tk.Entry(root, textvariable=rmm_var, width=65)
rmm_entry.pack()
rmm_button = tk.Button(root, text="Browse", command=lambda: select_file(rmm_var))
rmm_button.pack()
tk.Checkbutton(root, text="Skip Device Report (Disables 'Last Response')", variable=skip_rmm_var, command=toggle_rmm_state).pack()

tk.Label(root, text="3. Score Threshold (Show in Tabs)", font=('Arial', 10, 'bold')).pack(pady=5)
tk.Entry(root, textvariable=score_var, width=10).pack()

# --- NEW: Calendar Filter Block ---
tk.Label(root, text="4. RMM Check-in Cutoff Date", font=('Arial', 10, 'bold')).pack(pady=5)
date_frame = tk.Frame(root)
date_frame.pack()

date_var = tk.StringVar()
show_all_dates_var = tk.BooleanVar(value=True)

cal = DateEntry(date_frame, selectmode='day', textvariable=date_var, date_pattern='yyyy-mm-dd', width=12)
cal.pack(side=tk.LEFT, padx=5)

tk.Checkbutton(date_frame, text="Show All Dates", variable=show_all_dates_var, command=toggle_date_state).pack(side=tk.LEFT)
toggle_date_state() # Initialize the default disabled state

tk.Button(root, text="GENERATE COMPLETE DASHBOARD", command=process_reports, bg="#0078D7", fg="white", font=('Arial', 10, 'bold'), height=2).pack(pady=20)
root.mainloop()