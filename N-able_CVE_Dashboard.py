# ==========================================
# N-ABLE CVE REPORT MERGER & DASHBOARD UTILITY
# ==========================================

import pandas as pd                  # Used for data manipulation, merging, and grouping
import tkinter as tk                 # Used to build the Graphical User Interface (GUI)
from tkinter import filedialog, messagebox # Used for file selection popups and error alerts
import re                            # Used for regular expressions (cleaning text strings)

# ==========================================
# --- HELPER FUNCTIONS ---
# ==========================================

def select_file(label_var):
    """
    Opens a standard Windows file dialog window to select a CSV file.
    If a file is selected, it updates the corresponding text box in the GUI.
    """
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        label_var.set(file_path)

def clean_sheet_name(name, used_names):
    """
    Sanitizes software product names so they can be used as Excel sheet tabs.
    Excel crashes if sheet names have illegal characters, exceed 31 characters, or are duplicates.
    """
    # 1. Handle missing or completely blank product names
    if pd.isna(name) or str(name).strip() == "":
        name = "Unknown Product"
    
    # 2. Remove characters that break Excel sheet names
    name = str(name)
    invalid_chars = r'[\[\]\:\*\?\/\\\'\000]'
    clean_name = re.sub(invalid_chars, '', name).strip()
    
    # 3. Enforce Excel's strict 31-character limit per tab
    clean_name = clean_name[:31].strip()
    if not clean_name:  # If stripping characters left us with an empty string
        clean_name = "Unknown Product"
        
    # 4. Prevent duplicate sheet names (e.g., two products truncating to the exact same 31 characters)
    final_name = clean_name
    counter = 1
    # Check if the name (lowercased to ignore case) is already in our set of used names
    while final_name.lower() in [n.lower() for n in used_names]:
        suffix = f"_{counter}"
        # Ensure we don't exceed 31 chars when adding the number suffix (e.g., "_1")
        final_name = clean_name[:31 - len(suffix)] + suffix
        counter += 1
        
    # Add the finalized, safe name to our tracking set
    used_names.add(final_name)
    return final_name


# ==========================================
# --- MAIN PROCESSING FUNCTION ---
# ==========================================

def process_reports():
    """
    The main engine. Retrieves file paths from the GUI, loads the CSVs, validates headers,
    merges the data, calculates the dashboard metrics, and builds the final Excel file.
    """
    # Retrieve file paths entered in the GUI
    vuln_path = vuln_var.get()
    rmm_path = rmm_var.get()
    
    # Ensure the user actually selected both files before proceeding
    if not vuln_path or not rmm_path:
        messagebox.showerror("Error", "Please select both files.")
        return

    try:
        # Validate that the user entered a proper decimal number for the CVE score threshold
        try:
            threshold = float(score_var.get())
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid number for the CVE Score (e.g., 9.0).")
            return

        # ---------------------------------------------------------
        # STEP 1: LOAD & VALIDATE DATA
        # ---------------------------------------------------------
        # Attempt to load the Vulnerability Report
        try:
            df_vuln = pd.read_csv(vuln_path)
        except Exception as e:
            messagebox.showerror("File Error", f"Could not read Vulnerability Report:\n{e}")
            return
            
        # Attempt to load the RMM Offline Devices Report
        try:
            df_rmm = pd.read_csv(rmm_path)
        except Exception as e:
            messagebox.showerror("File Error", f"Could not read RMM Report:\n{e}")
            return

        # Fix RMM Headers: N-able sometimes exports without proper headers (e.g., showing 'Column4')
        if len(df_rmm.columns) == 9:
            # If it has exactly 9 columns, overwrite them with the standard N-able layout
            df_rmm.columns = ['Type', 'Client', 'Site', 'Device', 'Description', 'OS', 'Username', 'Last Response', 'Last Boot']
        else:
            # Fallback: Search the headers to dynamically find 'Device' and 'Last Response'
            col_lower = {c.lower(): c for c in df_rmm.columns}
            dev_col = col_lower.get('device') or col_lower.get('name') or col_lower.get('hostname') or col_lower.get('column4')
            resp_col = col_lower.get('last response') or col_lower.get('last check-in') or col_lower.get('column8')
            
            if dev_col and resp_col:
                df_rmm.rename(columns={dev_col: 'Device', resp_col: 'Last Response'}, inplace=True)
            else:
                messagebox.showerror("Format Error", "Could not identify 'Device' and 'Last Response' columns in the RMM report.")
                return

        # Ensure the vulnerability report actually contains the data we need to build the tables
        required_vuln_cols = ['Name', 'Vulnerability Name', 'Vulnerability Score', 'Affected Products', 'Vulnerability Severity']
        missing_cols = [col for col in required_vuln_cols if col not in df_vuln.columns]
        if missing_cols:
            messagebox.showerror("Format Error", f"Vulnerability report is missing required columns:\n{', '.join(missing_cols)}")
            return

        # Prevent 'groupby' functions from crashing if a product name is missing
        df_vuln['Affected Products'] = df_vuln['Affected Products'].fillna('Unknown Product')

        # ---------------------------------------------------------
        # STEP 2: MERGE THE DATA
        # ---------------------------------------------------------
        # Left join the RMM data onto the Vulnerability data matching 'Name' to 'Device'
        merged_df = pd.merge(
            df_vuln, 
            df_rmm[['Device', 'Last Response']], 
            left_on='Name', 
            right_on='Device', 
            how='left'
        )
        
        # Drop the redundant 'Device' column we pulled over
        if 'Device' in merged_df.columns:
            merged_df.drop(columns=['Device'], inplace=True)

        # Force 'Vulnerability Score' to be treated as a number. Changes "N/A" text to NaN.
        merged_df['Vulnerability Score'] = pd.to_numeric(merged_df['Vulnerability Score'], errors='coerce')

        # ---------------------------------------------------------
        # STEP 3: EXPORT TO EXCEL
        # ---------------------------------------------------------
        # Prompt user for save location
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not output_file:
            return 

        # Open the Excel writing engine
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            workbook = writer.book

            # --- PRE-CALCULATE SHEET NAMES FOR HYPERLINKS ---
            # Create a dataset strictly filtered by the user's score threshold
            filtered_for_sheets_df = merged_df[merged_df['Vulnerability Score'] >= threshold].copy()
            
            # Seed used names to protect 'Overview' and 'All Detections'
            used_sheet_names = set(['overview', 'all detections'])
            product_to_sheet = {}
            
            # Map the raw product name to the safe, sanitized Excel tab name
            for product, _ in filtered_for_sheets_df.groupby('Affected Products'):
                sheet_name = clean_sheet_name(product, used_sheet_names)
                product_to_sheet[product] = sheet_name

            # ==========================================
            # BUILD SHEET 1: THE OVERVIEW DASHBOARD
            # ==========================================
            overview_sheet = workbook.add_worksheet('Overview')
            header_format = workbook.add_format({'bold': True, 'font_size': 13})
            link_format = workbook.add_format({'font_color': 'blue', 'underline': True}) # Style for hyperlinks
            
            # --- 1. Severity Summary Table & Pie Chart (Overall Health) ---
            overview_sheet.write('A1', 'Vulnerabilities by Severity (All)', header_format)
            severity_counts = merged_df['Vulnerability Severity'].value_counts()
            
            row = 1
            for sev, count in severity_counts.items():
                overview_sheet.write(row, 0, str(sev))
                overview_sheet.write(row, 1, count)
                row += 1
                
            chart_sev = workbook.add_chart({'type': 'pie'})
            chart_sev.add_series({
                'name': 'Severity',
                'categories': ['Overview', 1, 0, row - 1, 0],
                'values':     ['Overview', 1, 1, row - 1, 1],
            })
            chart_sev.set_title({'name': 'Overall Severity Breakdown'})
            overview_sheet.insert_chart('D1', chart_sev)

            # --- 2. Top Products Summary Table (UNIQUE DEVICES ONLY) ---
            start_prod_row = max(12, row + 2) # Dynamically place below the first table
            overview_sheet.write(start_prod_row, 0, f'Top 10 Affected Products (Unique Devices, Score {threshold}+)', header_format)
            
            # Count the number of UNIQUE device names ('Name') per product, sort highest to lowest
            product_counts = filtered_for_sheets_df.groupby('Affected Products')['Name'].nunique().sort_values(ascending=False).head(10)
            
            p_row = start_prod_row + 1
            for prod, count in product_counts.items():
                # If the product qualified for its own tab, create an internal Excel link to it
                if prod in product_to_sheet:
                    target_sheet = product_to_sheet[prod]
                    overview_sheet.write_url(p_row, 0, f"internal:'{target_sheet}'!A1", string=str(prod), cell_format=link_format)
                else:
                    overview_sheet.write(p_row, 0, str(prod))
                
                overview_sheet.write(p_row, 1, count)
                p_row += 1
                
            chart_prod = workbook.add_chart({'type': 'bar'})
            chart_prod.add_series({
                'name': 'Products',
                'categories': ['Overview', start_prod_row + 1, 0, p_row - 1, 0],
                'values':     ['Overview', start_prod_row + 1, 1, p_row - 1, 1],
            })
            chart_prod.set_title({'name': f'Top 10 Affected Products (Unique Devices, Score {threshold}+)'})
            chart_prod.set_legend({'none': True}) 
            overview_sheet.insert_chart('D14', chart_prod)

            # Adjust column widths so text is readable
            overview_sheet.set_column('A:A', 45)
            overview_sheet.set_column('B:C', 18)
            overview_sheet.set_column('D:E', 25)

            # ==========================================
            # BUILD SHEET 2: ALL RAW DETECTIONS
            # ==========================================
            merged_df.to_excel(writer, sheet_name='All Detections', index=False)
            
            # ==========================================
            # BUILD SHEETS 3+: INDIVIDUAL SOFTWARE TABS
            # ==========================================
            # Define exactly which columns appear on these specific tabs
            cols_order = [
                'Vulnerability Name', 'Name', 'Vulnerability Severity', 
                'Vulnerability Score', 'Risk Severity Index', 'Has Known Exploit', 
                'CISA KEV', 'Last Response'
            ]

            # Custom aggregators used to squash multiple CVE rows into a single device row
            def combine_cves(x): return ', '.join(x.dropna().astype(str).unique())
            def any_yes(x): return 'Yes' if any(str(v).strip().lower() in ['yes', 'true'] for v in x) else 'No'

            # Loop through the filtered data grouped by the software product
            for product, group in filtered_for_sheets_df.groupby('Affected Products'):
                # Retrieve the safe sheet name we calculated earlier
                sheet_name = product_to_sheet[product] 
                
                # Sort so the highest severity threat bubbles to the top of the aggregation
                group = group.sort_values(by='Vulnerability Score', ascending=False)
                
                # Define how to combine the data for devices that appear multiple times
                agg_dict = {}
                if 'Vulnerability Name' in group.columns: agg_dict['Vulnerability Name'] = combine_cves
                if 'Vulnerability Severity' in group.columns: agg_dict['Vulnerability Severity'] = 'first' 
                if 'Vulnerability Score' in group.columns: agg_dict['Vulnerability Score'] = 'max'       
                if 'Risk Severity Index' in group.columns: agg_dict['Risk Severity Index'] = 'first'
                if 'Has Known Exploit' in group.columns: agg_dict['Has Known Exploit'] = any_yes         
                if 'CISA KEV' in group.columns: agg_dict['CISA KEV'] = any_yes                           
                if 'Last Response' in group.columns: agg_dict['Last Response'] = 'first'
                
                # Group by Device Name ('Name') to squash the duplicates
                device_summary = group.groupby('Name', as_index=False).agg(agg_dict)
                
                # Filter down to just our requested columns
                final_cols = [col for col in cols_order if col in device_summary.columns]
                
                # Write to the specific product tab
                device_summary[final_cols].to_excel(writer, sheet_name=sheet_name, index=False)

        # Alert the user that processing is done
        messagebox.showinfo("Success", f"Dashboard and Report saved to:\n{output_file}")

    except Exception as e:
        # Catch any unexpected errors (like having the Excel file already open while trying to overwrite)
        messagebox.showerror("Processing Error", f"An unexpected error occurred:\n{str(e)}")


# ==========================================
# GUI SETUP (TKINTER)
# ==========================================

# Initialize the main window
root = tk.Tk()
root.title("N-able CVE Report Merger & Dashboard")
root.geometry("520x350") # Set dimensions

# Variables to hold the user inputs
vuln_var = tk.StringVar()
rmm_var = tk.StringVar()
score_var = tk.StringVar(value="9.0") # Default score threshold

# GUI Element: File 1 Input
tk.Label(root, text="Step 1: Select Vulnerability Report (Detections)", font=('Arial', 10, 'bold')).pack(pady=(15, 2))
tk.Entry(root, textvariable=vuln_var, width=65).pack()
tk.Button(root, text="Browse", command=lambda: select_file(vuln_var)).pack()

# GUI Element: File 2 Input
tk.Label(root, text="Step 2: Select RMM Device Report (Offline Devices)", font=('Arial', 10, 'bold')).pack(pady=(15, 2))
tk.Entry(root, textvariable=rmm_var, width=65).pack()
tk.Button(root, text="Browse", command=lambda: select_file(rmm_var)).pack()

# GUI Element: Score Threshold Input
tk.Label(root, text="Step 3: Minimum CVE Score for Product Sheets", font=('Arial', 10, 'bold')).pack(pady=(15, 2))
tk.Entry(root, textvariable=score_var, width=15, justify='center').pack()

# GUI Element: Submit Button
tk.Button(root, text="GENERATE DASHBOARD EXCEL", command=process_reports, bg="#0078D7", fg="white", font=('Arial', 10, 'bold'), height=2).pack(pady=20)

# Keeps the GUI window open and waiting for user interaction
root.mainloop()