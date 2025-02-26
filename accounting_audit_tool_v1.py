import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, scrolledtext
import ttkbootstrap as ttk
from PIL import Image, ImageTk
from datetime import datetime
import re
import os
import logging
import sys

# Set up logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

class AuditApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AppGlide Accounting Audit Tool")
        self.root.geometry("1200x900")
        
        # Load logo
        logo_path = r"C:\Users\Admin\Downloads\AuditTool\appglide-logo.png"
        try:
            self.logo_photo = ImageTk.PhotoImage(Image.open(logo_path))
            tk.Label(root, image=self.logo_photo, bg="#4A90E2").pack(pady=10, anchor="center")
        except Exception as e:
            logging.error(f"Error loading logo: {e}")
            tk.Label(root, text="AppGlide", font=("Arial", 24, "bold"), bg="#4A90E2", fg="white").pack(pady=10, anchor="center")
        
        # ttkbootstrap styling
        style = ttk.Style("flatly")
        style.configure("TButton", font=("Arial", 12), background="#50C878", foreground="white")
        style.configure("TAnalyze.TButton", font=("Arial", 12), background="#F5A623", foreground="white")
        style.map("TButton", background=[("active", "#4A90E2")])
        style.map("TAnalyze.TButton", background=[("active", "#F5A623")])
        
        frame = ttk.Frame(root)
        frame.pack(pady=20, padx=20, fill="x")
        
        # --- COA Upload Section ---
        tk.Label(frame, text="Upload Chart of Accounts", font=("Arial", 14, "bold"),
                 bg="#F5F7FA", fg="#333333").grid(row=0, column=0, padx=20, pady=5)
        coa_btn = ttk.Button(frame, text="Upload CoA File", command=self.upload_coa)
        coa_btn.grid(row=1, column=0, padx=20, pady=5)
        self.coa_status = tk.Label(frame, text="COA Status: Not Loaded", font=("Arial", 10),
                                   bg="#F5F7FA", fg="#333333")
        self.coa_status.grid(row=2, column=0, padx=20, pady=5)
        
        # --- Other Sections ---
        sections = [
            ("Trial Balance", "trial_balance"),
            ("Ledger Transactions", "ledger_transactions"),
            ("Profit & Loss", "p_l"),
            ("Balance Sheet", "balance_sheet")
        ]
        self.files = {key: [] for key in [s[1] for s in sections]}
        self.coa_file = []
        self.periods = {"p_l": [], "balance_sheet": [], "trial_balance": []}
        self.chart_of_accounts = None
        
        # Define standard columns; note ledger_transactions does not have a "Periods" key.
        self.STANDARD_COLUMNS = {
            "trial_balance": {
                "Account Name": "Accounts|Account|AcctName|Name|Acct #",
                "Period": "Period|Date",
                "Amount": "Amount|Balance"
            },
            "ledger_transactions": {
                "Accounts": "Accounts|Account|AcctName|Name|Acct #",
                "Date": "Date|Transaction Date",
                "Transaction Type": "Transaction Type|Type|Trans Type",
                "Num": "Num|Invoice Number|Number",
                "Name": "Name|Customer Name|Vendor Name",
                "Memo/Description": "Memo/Description|Description|Notes",
                "Split": "Split|Split Account|Allocation",
                "Amount": "Amount|Transaction Amount"
            },
            "p_l": {
                "Accounts": "Accounts|Account|AcctName|Name|Acct #",
                "Periods": {"formatted": [], "original": []}  # to be filled from file headers
            },
            "balance_sheet": {
                "Accounts": "Accounts|Account|AcctName|Name|Acct #",
                "Periods": {"formatted": [], "original": []}  # to be filled from file headers
            }
        }
        
        # Upload buttons for other sections
        for i, (label, key) in enumerate(sections, start=1):
            tk.Label(frame, text=f"Upload {label}", font=("Arial", 14, "bold"),
                     bg="#F5F7FA", fg="#333333").grid(row=0, column=i, padx=20, pady=5)
            btn = ttk.Button(frame, text=f"Upload {label} File(s)",
                             command=lambda k=key: self.upload_files(k))
            btn.grid(row=1, column=i, padx=20, pady=5)
        
        # Analyze button
        analyze_btn = ttk.Button(frame, text="Analyze", command=self.analyze,
                                 style="TAnalyze.TButton")
        analyze_btn.grid(row=2, column=0, columnspan=5, pady=10, padx=20)
        
        # Results text area
        self.results_text = scrolledtext.ScrolledText(root, width=130, height=40, font=("Arial", 12))
        self.results_text.pack(pady=20, padx=20, fill="both", expand=True)
        self.results_text.configure(bg="#FFFFFF", fg="#333333")
        
        footer = tk.Label(root, text="Powered by AppGlide", font=("Arial", 10, "italic"),
                          bg="#50C878", fg="white")
        footer.pack(side="bottom", fill="x", pady=5)
        
        # Hover animations for buttons
        def on_enter(e):
            if "TAnalyze" in e.widget.winfo_class():
                e.widget.config(background="#F5A623")
            else:
                e.widget.config(background="#4A90E2")
        def on_leave(e):
            if "TAnalyze" in e.widget.winfo_class():
                e.widget.config(background="#F5A623")
            else:
                e.widget.config(background="#50C878")
        for widget in root.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.bind("<Enter>", on_enter)
                widget.bind("<Leave>", on_leave)
        
        self.column_mappings = {}

    # -------------------------------
    # COA Functions
    # -------------------------------
    def upload_coa(self):
        try:
            file_types = [("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")]
            files = filedialog.askopenfilenames(filetypes=file_types, title="Upload Chart of Accounts File(s)")
            if files:
                self.coa_file = files
                self.load_chart_of_accounts()
                self.coa_status.config(text="COA Status: Loaded Successfully", fg="#7ED321")
                messagebox.showinfo("Success", "Chart of Accounts loaded successfully.")
                logging.debug(f"Uploaded {len(files)} Chart of Accounts file(s).")
        except Exception as e:
            self.coa_status.config(text="COA Status: Error Loading", fg="#D0021B")
            logging.error(f"Error in upload_coa: {e}")
            messagebox.showerror("Error", f"Failed to upload Chart of Accounts file: {e}")

    def load_chart_of_accounts(self):
        try:
            if not self.coa_file:
                raise ValueError("No Chart of Accounts file uploaded!")
            file_type = os.path.splitext(self.coa_file[0])[1][1:].lower()
            if file_type == "xlsx":
                coa_df = pd.read_excel(self.coa_file[0], engine='openpyxl')
            else:
                coa_df = pd.read_csv(self.coa_file[0])
            # Clean up column names and fill NA
            coa_df.columns = [col.strip() for col in coa_df.columns]
            coa_df = coa_df.fillna({'Accounts': '', 'Type': '', 'Detail type': '', 'Description': '', 'Total balance': 0})
            required = ['Accounts', 'Type', 'Detail type']
            missing = [c for c in required if c not in coa_df.columns]
            if missing:
                raise ValueError(f"Missing required columns in COA: {', '.join(missing)}")
            type_mapping = {
                'Income': 'Income',
                'Expenses': 'Expenses',
                'Cost of Goods Sold': 'Expenses',
                'Assets': 'Assets',
                'Liabilities': 'Liabilities',
                'Equity': 'Equity',
                'Fixed Assets': 'Assets',
                'Other Current Assets': 'Assets',
                'Other Current Liabilities': 'Liabilities',
                'Long Term Liabilities': 'Liabilities',
                'Other Assets': 'Assets',
                'Other Income': 'Income',
                'Other Expense': 'Expenses'
            }
            coa_df['GAAP_Category'] = coa_df['Type'].map(type_mapping).fillna('Other')
            self.chart_of_accounts = coa_df
            logging.debug(f"Chart of Accounts loaded successfully: {coa_df.columns.tolist()}")
        except Exception as e:
            logging.error(f"Error in load_chart_of_accounts: {e}")
            raise

    # -------------------------------
    # Upload Section Files
    # -------------------------------
    def upload_files(self, section):
        try:
            file_types = [("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")]
            files = filedialog.askopenfilenames(filetypes=file_types, title=f"Upload {section.replace('_', ' ').title()} File(s)")
            if files:
                self.files[section] = files
                # For ledger_transactions, do not extract periods; for P&L/BS, extract period columns.
                if section in ["p_l", "balance_sheet"]:
                    self.extract_periods_from_file(section)
                else:
                    self.map_columns(section)
                logging.debug(f"Uploaded {len(files)} {section.replace('_', ' ').title()} file(s).")
        except Exception as e:
            logging.error(f"Error in upload_files for {section}: {e}")
            messagebox.showerror("Error", f"Failed to upload {section.replace('_', ' ').title()} files: {e}")

    # -------------------------------
    # Extract Periods (for P&L & BS & Trial Balance)
    # -------------------------------
    def extract_periods_from_file(self, section):
        try:
            if not self.files[section]:
                raise ValueError(f"No {section.replace('_', ' ').title()} file uploaded!")
            file_type = os.path.splitext(self.files[section][0])[1][1:].lower()
            if file_type == "xlsx":
                df = pd.read_excel(self.files[section][0], engine='openpyxl', nrows=1)
            else:
                df = pd.read_csv(self.files[section][0], nrows=1)
            df.columns = [col.strip() for col in df.columns.astype(str)]
            headers = df.columns.tolist()
            periods = {"formatted": [], "original": []}
            # For trial_balance, we expect a 'Period' column
            if section == "trial_balance":
                if file_type == "xlsx":
                    full_df = pd.read_excel(self.files[section][0], engine='openpyxl')
                else:
                    full_df = pd.read_csv(self.files[section][0])
                if 'Period' in full_df.columns:
                    unique_periods = full_df['Period'].dropna().unique()
                    for period in unique_periods:
                        try:
                            dt = pd.to_datetime(period, format="%Y-%m-%d", errors="raise")
                        except Exception:
                            dt = pd.to_datetime(period, errors="coerce")
                        if pd.notna(dt):
                            periods["original"].append(dt.strftime("%Y-%m-%d 00:00:00"))
                            periods["formatted"].append(dt.strftime("%b %d, %Y"))
                else:
                    raise ValueError("No 'Period' column found in Trial Balance file to extract periods.")
            else:
                # For P&L and BS, try to parse headers (excluding known account columns)
                exclude = ['Accounts', 'Account Name', 'Period', 'Amount']
                for header in headers:
                    if header not in exclude:
                        try:
                            dt = pd.to_datetime(header, format="%Y-%m-%d %H:%M:%S", errors="raise")
                        except Exception:
                            try:
                                dt = pd.to_datetime(header, format="%m/%d/%Y", errors="raise")
                            except Exception:
                                try:
                                    dt = pd.to_datetime(header, format="%b %d, %Y", errors="raise")
                                except Exception:
                                    try:
                                        dt = pd.to_datetime(header, format="%b %Y", errors="raise")
                                    except Exception:
                                        continue
                        if pd.notna(dt):
                            periods["original"].append(dt.strftime("%Y-%m-%d 00:00:00"))
                            periods["formatted"].append(dt.strftime("%b %d, %Y"))
            if not periods["original"]:
                raise ValueError(f"No valid period columns found in {section.replace('_', ' ').title()} file.")
            # For P&L and BS, update STANDARD_COLUMNS with extracted periods
            if section in ["p_l", "balance_sheet"]:
                self.STANDARD_COLUMNS[section]["Periods"] = periods
            logging.debug(f"Extracted periods for {section}: Original: {periods['original']}, Formatted: {periods['formatted']}")
            # Now call map_columns to update mapping window with period columns (if applicable)
            if section in ["p_l", "balance_sheet", "trial_balance"]:
                self.map_columns(section)
        except Exception as e:
            logging.error(f"Error in extract_periods_from_file for {section}: {e}")
            messagebox.showerror("Error", f"Failed to extract periods for {section.replace('_', ' ').title()}: {e}")

    # -------------------------------
    # Map Columns (for manual mapping)
    # -------------------------------
    def map_columns(self, section):
        try:
            if not self.files[section]:
                messagebox.showerror("Error", f"No {section.replace('_', ' ').title()} file uploaded!")
                return
            file_type = os.path.splitext(self.files[section][0])[1][1:].lower()
            if file_type == "xlsx":
                df = pd.read_excel(self.files[section][0], engine='openpyxl', nrows=1)
            else:
                df = pd.read_csv(self.files[section][0], nrows=1)
            # Normalize column names (strip spaces)
            df.columns = [col.strip() for col in df.columns.astype(str)]
            file_columns = df.columns.tolist()
            if not file_columns or all(col.strip() == "" for col in file_columns):
                raise ValueError(f"No valid column headers found in the {section.replace('_', ' ').title()} file.")
            logging.debug(f"File columns for {section}: {file_columns}")
            
            mapping_window = tk.Toplevel(self.root)
            mapping_window.title(f"Map Columns for {section.replace('_', ' ').title()}")
            mapping_window.geometry("600x400")
            tk.Label(mapping_window, text=f"Map your file columns to standard {section.replace('_', ' ').title()} columns:",
                     font=("Arial", 12, "bold")).pack(pady=10)
            
            # Determine standard columns to map
            if section == "ledger_transactions":
                # Ledger transactions do not have period columns to map
                standard_cols = list(self.STANDARD_COLUMNS[section].keys())
            elif section == "trial_balance":
                standard_cols = list(self.STANDARD_COLUMNS[section].keys())
            else:
                # For P&L and BS, add the period columns from STANDARD_COLUMNS
                periods_list = self.STANDARD_COLUMNS[section]["Periods"]["original"]
                base_cols = [col for col in self.STANDARD_COLUMNS[section].keys() if col != "Periods"]
                standard_cols = base_cols + periods_list
            
            self.column_mappings[section] = {}
            for std_col in standard_cols:
                frame = ttk.Frame(mapping_window)
                frame.pack(pady=5, padx=10, fill="x")
                tk.Label(frame, text=f"Standard: {std_col}", font=("Arial", 10)).pack(side="left")
                var = tk.StringVar()
                options = [""] + file_columns
                default_value = ""
                # For account columns, try a direct case-insensitive match
                if self.chart_of_accounts is not None and std_col in ["Account Name", "Accounts"]:
                    for fc in file_columns:
                        if fc.strip().lower() == "accounts":
                            default_value = fc
                            break
                elif section in ["p_l", "balance_sheet"] and std_col in self.STANDARD_COLUMNS[section].get("Periods", {}).get("original", []):
                    # For period columns, attempt to match by comparing formatted version
                    for fc in file_columns:
                        try:
                            dt = pd.to_datetime(fc, errors="coerce")
                            if pd.notna(dt):
                                if dt.strftime("%b %d, %Y") == std_col or dt.strftime("%Y-%m-%d 00:00:00") == std_col:
                                    default_value = fc
                                    break
                        except Exception:
                            continue
                else:
                    pattern = self.STANDARD_COLUMNS[section].get(std_col, "")
                    if pattern:
                        for fc in file_columns:
                            if re.match(f"^{pattern}$", fc, re.IGNORECASE):
                                default_value = fc
                                break
                var.set(default_value)
                if not default_value:
                    logging.warning(f"No automatic match found for '{std_col}' in {section}. Please map manually.")
                combo = ttk.Combobox(frame, textvariable=var, values=options, state="readonly")
                combo.pack(side="right", fill="x", expand=True)
                self.column_mappings[section][std_col] = var.get()
            
            ttk.Button(mapping_window, text="Save Mappings", command=lambda: self.save_mappings(section, mapping_window)).pack(pady=10)
            logging.debug(f"Column mapping window opened for {section} with standard columns: {standard_cols}")
        except Exception as e:
            logging.error(f"Error in map_columns for {section}: {e}")
            messagebox.showerror("Error", f"Failed to map columns for {section.replace('_', ' ').title()}: {e}")

    def save_mappings(self, section, mapping_window):
        try:
            mappings = {k: v for k, v in self.column_mappings[section].items() if v}
            if not mappings:
                messagebox.showwarning("Warning", f"No columns mapped for {section}! Using defaults.")
                mappings = {}
            # Validate against COA for account columns
            if self.chart_of_accounts is not None and section in ["trial_balance", "ledger_transactions", "p_l", "balance_sheet"]:
                acct_col = mappings.get("Account Name") or mappings.get("Accounts")
                if acct_col:
                    df_temp = self.load_data(section)
                    if not df_temp.empty and acct_col in df_temp.columns:
                        acct_vals = df_temp[acct_col].str.lower().unique()
                        coa_vals = self.chart_of_accounts['Accounts'].str.lower().unique()
                        unmapped = [a for a in acct_vals if a.strip() and a not in coa_vals]
                        if unmapped:
                            logging.warning(f"Unmapped accounts in {section}: {unmapped}")
                            messagebox.showwarning("Warning", f"Some accounts in {section} do not match the COA: {unmapped}.")
            # Validate numeric/date columns
            file_type = os.path.splitext(self.files[section][0])[1][1:].lower()
            if file_type == "xlsx":
                df_check = pd.read_excel(self.files[section][0], engine='openpyxl')
            else:
                df_check = pd.read_csv(self.files[section][0])
            df_check.columns = [col.strip() for col in df_check.columns.astype(str)]
            for std_col, file_col in mappings.items():
                if file_col:
                    if not any(file_col.lower() == col.lower() for col in df_check.columns):
                        logging.warning(f"Column '{file_col}' mapped to '{std_col}' not found in {section} file.")
                        continue
                    if std_col in ["Date", "Period"]:
                        if not pd.to_datetime(df_check[file_col], errors="coerce").notna().any():
                            messagebox.showerror("Error", f"Column '{file_col}' for {std_col} must contain date values!")
                            return
                    elif std_col == "Amount":
                        if not pd.to_numeric(df_check[file_col], errors="coerce").notna().any():
                            messagebox.showerror("Error", f"Column '{file_col}' for {std_col} must contain numeric values!")
                            return
            self.column_mappings[section] = mappings
            mapping_window.destroy()
            messagebox.showinfo("Success", f"Column mappings saved for {section}. Ready to analyze.")
            logging.debug(f"Column mappings saved successfully for {section}.")
        except Exception as e:
            logging.error(f"Error in save_mappings for {section}: {e}")
            messagebox.showerror("Error", f"Failed to save mappings for {section}: {e}")

    # -------------------------------
    # Load Data with Mappings
    # -------------------------------
    def load_data(self, section):
        try:
            if not self.files[section] or not self.column_mappings.get(section):
                logging.warning(f"No {section} file or mappings provided. Returning default DataFrame.")
                return pd.DataFrame(columns=self.STANDARD_COLUMNS[section].keys()).fillna({
                    'Account Name': '', 'Accounts': '', 'Period': pd.NaT, 'Date': pd.NaT,
                    'Transaction Type': '', 'Num': '', 'Name': '', 'Memo/Description': '', 'Split': '',
                    'Category/Product/Service': '', 'Amount': 0
                })
            file_type = os.path.splitext(self.files[section][0])[1][1:].lower()
            if file_type == "xlsx":
                df = pd.read_excel(self.files[section][0], engine='openpyxl')
            else:
                df = pd.read_csv(self.files[section][0])
            df.columns = [col.strip() for col in df.columns.astype(str)]
            df = df.fillna({
                'Accounts': '', 'Account Name': '', 'Period': pd.NaT, 'Date': pd.NaT,
                'Transaction Type': '', 'Num': '', 'Name': '', 'Memo/Description': '', 'Split': '',
                'Category/Product/Service': ''
            })
            for col in df.columns:
                if df[col].dtype in ['float64', 'int64']:
                    df[col] = df[col].fillna(0)
            mapped_df = pd.DataFrame()
            for std_col, file_col in self.column_mappings[section].items():
                # Use a case-insensitive match for file columns
                if file_col:
                    matched = None
                    for col in df.columns:
                        if col.lower() == file_col.lower():
                            matched = col
                            break
                    if matched:
                        mapped_df[std_col] = df[matched]
                    else:
                        logging.warning(f"Skipping missing column '{std_col}' for {section}.")
                        mapped_df[std_col] = '' if std_col in ["Accounts", "Account Name"] else (0 if std_col=="Amount" else pd.NaT)
                else:
                    logging.warning(f"Skipping unmapped column '{std_col}' for {section}.")
                    mapped_df[std_col] = '' if std_col in ["Accounts", "Account Name"] else (0 if std_col=="Amount" else pd.NaT)
            # Section-specific transformations
            # For trial_balance, parse period and amount; then categorize accounts
            if section == "trial_balance":
                if "Period" in mapped_df.columns:
                    mapped_df["Period"] = pd.to_datetime(mapped_df["Period"], errors="coerce")
                if "Amount" in mapped_df.columns:
                    mapped_df["Amount"] = pd.to_numeric(mapped_df["Amount"], errors="coerce").fillna(0)
                acct_col = self.column_mappings.get("trial_balance", {}).get("Account Name") or \
                           self.column_mappings.get("trial_balance", {}).get("Accounts")
                if self.chart_of_accounts is not None and acct_col:
                    mapped_df = self.categorize_accounts(mapped_df, acct_col)
            elif section == "ledger_transactions":
                if "Date" in mapped_df.columns:
                    mapped_df["Date"] = pd.to_datetime(mapped_df["Date"], errors="coerce")
                if "Amount" in mapped_df.columns:
                    mapped_df["Amount"] = pd.to_numeric(mapped_df["Amount"], errors="coerce").fillna(0)
                acct_col = self.column_mappings.get("ledger_transactions", {}).get("Accounts")
                if self.chart_of_accounts is not None and acct_col:
                    mapped_df = self.categorize_accounts(mapped_df, acct_col)
            elif section in ["p_l", "balance_sheet"]:
                # For these, convert period columns to numeric
                period_cols = [c for c in mapped_df.columns if c in self.STANDARD_COLUMNS[section]["Periods"]["original"]]
                for c in period_cols:
                    try:
                        mapped_df[c] = pd.to_numeric(mapped_df[c], errors="coerce").fillna(0)
                    except Exception as e:
                        logging.error(f"Error converting {c} in {section}: {e}")
                        mapped_df[c] = 0
                if section == "p_l" and "Total" in mapped_df.columns:
                    mapped_df.drop(columns=["Total"], inplace=True)
                acct_col = self.column_mappings.get(section, {}).get("Accounts")
                if self.chart_of_accounts is not None and acct_col:
                    mapped_df = self.categorize_accounts(mapped_df, acct_col)
            if mapped_df.empty:
                logging.warning(f"No data loaded for {section}. Returning default DataFrame.")
                return pd.DataFrame(columns=self.STANDARD_COLUMNS[section].keys()).fillna({
                    'Account Name': '', 'Accounts': '', 'Period': pd.NaT, 'Date': pd.NaT,
                    'Transaction Type': '', 'Num': '', 'Name': '', 'Memo/Description': '', 'Split': '',
                    'Category/Product/Service': '', 'Amount': 0
                })
            logging.debug(f"Data loaded for {section}. First 5 rows:\n{mapped_df.head().to_string()}")
            return mapped_df
        except Exception as e:
            logging.error(f"Error in load_data for {section}: {e}")
            messagebox.showerror("Error", f"Failed to load {section}: {e}")
            return pd.DataFrame()

    # -------------------------------
    # Categorize Accounts using COA
    # -------------------------------
    def categorize_accounts(self, df, account_col):
        if self.chart_of_accounts is None:
            logging.warning("COA not loaded. Marking GAAP_Category as 'Uncategorized'.")
            df["GAAP_Category"] = "Uncategorized"
            return df
        if account_col not in df.columns:
            logging.warning(f"Account column '{account_col}' not found. Marking GAAP_Category as 'Uncategorized'.")
            df["GAAP_Category"] = "Uncategorized"
            return df
        df[account_col] = df[account_col].astype(str).str.strip().str.lower()
        self.chart_of_accounts["Accounts"] = self.chart_of_accounts["Accounts"].astype(str).str.strip().str.lower()
        coa_map = dict(zip(self.chart_of_accounts["Accounts"], self.chart_of_accounts["GAAP_Category"]))
        df["GAAP_Category"] = df[account_col].map(coa_map).fillna("Uncategorized")
        logging.debug(f"Categorized accounts using '{account_col}': {df['GAAP_Category'].unique()}")
        return df

    # -------------------------------
    # Detect GAAP Errors
    # -------------------------------
    def detect_gaap_errors(self, dfs):
        try:
            errors = []
            tb_df = dfs.get("trial_balance", pd.DataFrame())
            trans_df = dfs.get("ledger_transactions", pd.DataFrame())
            p_l_df = dfs.get("p_l", pd.DataFrame())
            bs_df = dfs.get("balance_sheet", pd.DataFrame())
            
            tb_df = tb_df.fillna({'Account Name': '', 'Accounts': '', 'Period': pd.NaT, 'Amount': 0})
            trans_df = trans_df.fillna({'Accounts': '', 'Date': pd.NaT, 'Transaction Type': '', 'Num': '',
                                        'Name': '', 'Memo/Description': '', 'Split': '', 'Amount': 0,
                                        'Category/Product/Service': ''})
            p_l_df = p_l_df.fillna({'Accounts': '', **{c: 0 for c in p_l_df.columns if c not in ['Accounts','GAAP_Category']}})
            bs_df = bs_df.fillna({'Accounts': '', **{c: 0 for c in bs_df.columns if c not in ['Accounts','GAAP_Category']}})
            
            # (For brevity, only TB and BS checks are illustrated below. Extend similar logic for revenue, expense, etc.)
            # 1. Trial Balance Integrity Check
            tb_acct_col = self.column_mappings.get("trial_balance", {}).get("Account Name") or \
                          self.column_mappings.get("trial_balance", {}).get("Accounts")
            if not tb_df.empty and self.chart_of_accounts is not None and tb_acct_col:
                if "Period" in tb_df.columns and "Amount" in tb_df.columns and tb_acct_col in tb_df.columns:
                    for period in tb_df["Period"].dropna().unique():
                        period_data = tb_df[tb_df["Period"] == period]
                        total_debit = total_credit = 0
                        for _, row in period_data.iterrows():
                            amt = row["Amount"]
                            acct = str(row[tb_acct_col]).strip().lower()
                            if not acct:
                                continue
                            if acct in self.chart_of_accounts["Accounts"].values:
                                gaap_cat = self.chart_of_accounts.loc[self.chart_of_accounts["Accounts"] == acct, "GAAP_Category"].iloc[0]
                            else:
                                gaap_cat = "Uncategorized"
                            if gaap_cat in ['Assets','Expenses','Cost of Goods Sold']:
                                total_debit += max(amt, 0)
                                total_credit += abs(min(amt, 0))
                            else:
                                total_debit += abs(min(amt, 0))
                                total_credit += max(amt, 0)
                        if abs(total_debit - total_credit) > 0.001:
                            p_str = period.strftime("%b %d, %Y") if hasattr(period, "strftime") else str(period)
                            errors.append(f"TB {p_str} mismatch: Debits = ${total_debit:.2f}, Credits = ${total_credit:.2f}")
            else:
                logging.warning("TB data or COA or account mapping missing. Skipping TB check.")
            
            # 2. Balance Sheet Integrity Check
            if not bs_df.empty and self.chart_of_accounts is not None:
                periods_orig = self.STANDARD_COLUMNS["balance_sheet"]["Periods"]["original"]
                if periods_orig:
                    latest_period = periods_orig[-1]
                    if latest_period in bs_df.columns and "Accounts" in bs_df.columns:
                        bs_df["Accounts"] = bs_df["Accounts"].fillna('')
                        assets = bs_df[bs_df["GAAP_Category"].isin(['Assets','Fixed Assets','Other Current Assets','Other Assets'])][latest_period].sum()
                        liab = bs_df[bs_df["GAAP_Category"].isin(['Liabilities','Other Current Liabilities','Long Term Liabilities'])][latest_period].sum()
                        eqty = bs_df[bs_df["GAAP_Category"]=='Equity'][latest_period].sum()
                        if abs(assets - (liab + eqty)) > 0.001:
                            fperiod = pd.to_datetime(latest_period).strftime("%b %d, %Y")
                            errors.append(f"BS imbalance: Assets ${assets:.2f}, Liabilities ${liab:.2f}, Equity ${eqty:.2f} as of {fperiod}")
                else:
                    logging.warning("No period columns in BS. Skipping BS check.")
            else:
                logging.warning("BS data or COA missing. Skipping BS check.")
            
            # (Additional GAAP checks would be implemented similarly using the loaded data.)
            
            logging.debug("GAAP errors detected successfully.")
            return errors
        except Exception as e:
            logging.error(f"Error in detect_gaap_errors: {str(e)}")
            messagebox.showerror("Error", f"Failed to detect GAAP errors: {str(e)}")
            return []

    # -------------------------------
    # Tax Preparation
    # -------------------------------
    def tax_preparation(self, dfs):
        try:
            adjustments = []
            taxable_income = 0
            p_l_df = dfs.get("p_l", pd.DataFrame())
            trans_df = dfs.get("ledger_transactions", pd.DataFrame())
            p_l_df = p_l_df.fillna({'Accounts': '', **{c: 0 for c in p_l_df.columns if c not in ['Accounts','GAAP_Category']}})
            trans_df = trans_df.fillna({'Accounts': '', 'Date': pd.NaT, 'Transaction Type': '', 'Num': '', 'Name': '',
                                        'Memo/Description': '', 'Split': '', 'Amount': 0, 'Category/Product/Service': ''})
            if p_l_df.empty:
                logging.warning("P&L data empty. Using defaults for tax prep.")
            if trans_df.empty:
                logging.warning("Ledger Transactions empty. Using defaults for tax prep.")
            sales = expenses = 0
            if not p_l_df.empty and not trans_df.empty and self.chart_of_accounts is not None:
                periods_orig = self.STANDARD_COLUMNS["p_l"]["Periods"]["original"]
                if periods_orig:
                    sales_cols = [col for col in periods_orig if col in p_l_df.columns]
                    if "Accounts" in p_l_df.columns and sales_cols:
                        sales = p_l_df[p_l_df["GAAP_Category"].isin(['Income'])][sales_cols].sum().sum()
                        expense_cols = [col for col in periods_orig if col in p_l_df.columns]
                        expenses = p_l_df[p_l_df["GAAP_Category"].isin(['Expenses','Cost of Goods Sold'])][expense_cols].sum().sum()
                    cogs = trans_df[trans_df["GAAP_Category"].isin(['Cost of Goods Sold'])]["Amount"].sum() if "Category/Product/Service" in trans_df.columns else (sales * 0.6 if sales > 0 else 0)
                    capital_items = ["Tools, machinery, and equipment", "Fixed Assets", "Long-term office equipment STYKU",
                                     "STYKU", "Vehicles", "Buildings", "Furniture & fixtures", "Improvements", "Land"]
                    if "Category/Product/Service" in trans_df.columns:
                        for item in capital_items:
                            cap_exp = trans_df[trans_df["GAAP_Category"].isin(['Fixed Assets','Assets'])]["Amount"].sum()
                            if cap_exp > 1000:
                                life = 5
                                annual_exp = cap_exp / life
                                adjustments.append({"Item": item, "Original": cap_exp, "Adjusted": annual_exp, "To BS": cap_exp - annual_exp})
                                expenses -= (cap_exp - annual_exp)
                    depreciation = sum(adj["Adjusted"] for adj in adjustments)
                    adjusted_exp = expenses - sum(adj["Original"] - adj["Adjusted"] for adj in adjustments)
                    taxable_income = max(0, sales - cogs - adjusted_exp - depreciation)
                else:
                    logging.warning("No period columns in P&L. Using default values for tax prep.")
            else:
                logging.warning("Using default values for tax prep due to missing data.")
                cogs = 0
            logging.debug("Tax preparation completed successfully.")
            return adjustments, taxable_income
        except Exception as e:
            logging.error(f"Error in tax_preparation: {str(e)}")
            messagebox.showerror("Error", f"Failed to prepare tax data: {str(e)}")
            return [], 0

    # -------------------------------
    # Analyze – Main Function
    # -------------------------------
    def analyze(self):
        try:
            if not any(self.files.values()) and not self.coa_file:
                messagebox.showerror("Error", "No files uploaded, including Chart of Accounts!")
                return
            dfs = {}
            for sec in ["trial_balance", "ledger_transactions", "p_l", "balance_sheet"]:
                dfs[sec] = self.load_data(sec)
                if dfs[sec].empty:
                    messagebox.showwarning("Warning", f"No data loaded for {sec}. Analysis may be incomplete.")
            self.results_text.delete(1.0, tk.END)
            header_text = f"Comprehensive Accounting Audit as of {datetime.now().strftime('%B %d, %Y')}\n\n"
            self.results_text.insert(tk.END, header_text, "header")
            self.results_text.tag_config("header", font=("Arial", 14, "bold"), foreground="#4A90E2")
            self.results_text.tag_config("error", foreground="#D0021B")
            self.results_text.tag_config("success", foreground="#7ED321")
            self.results_text.tag_config("highlight", font=("Arial", 12, "italic"), foreground="#7ED321")
            
            self.results_text.insert(tk.END, "=== GAAP Compliance Errors ===\n", "header")
            gaap_errors = self.detect_gaap_errors(dfs)
            if gaap_errors:
                for err in gaap_errors:
                    self.results_text.insert(tk.END, f"⚠️ {err}\n", "error")
            else:
                self.results_text.insert(tk.END, "✅ No GAAP errors detected.\n", "success")
            
            self.results_text.insert(tk.END, "\n=== Tax Preparation ===\n", "header")
            p_l_df = dfs.get("p_l", pd.DataFrame())
            adjustments, taxable_income = self.tax_preparation(dfs)
            sales = expenses = 0
            if not p_l_df.empty and "Accounts" in p_l_df.columns:
                periods_orig = self.STANDARD_COLUMNS["p_l"]["Periods"]["original"]
                if periods_orig:
                    s_cols = [c for c in periods_orig if c in p_l_df.columns]
                    if s_cols:
                        sales = p_l_df[p_l_df["GAAP_Category"].isin(['Income'])][s_cols].sum().sum()
                        e_cols = [c for c in periods_orig if c in p_l_df.columns]
                        expenses = p_l_df[p_l_df["GAAP_Category"].isin(['Expenses','Cost of Goods Sold'])][e_cols].sum().sum()
                        self.results_text.insert(tk.END, f"Sales: ${sales:.2f}\n")
                        self.results_text.insert(tk.END, f"Expenses (before adjustments): ${expenses:.2f}\n")
            for adj in adjustments:
                self.results_text.insert(tk.END, f"Adjusted {adj['Item']}: Expensed ${adj['Adjusted']:.2f}, Moved to BS ${adj['To BS']:.2f}\n", "highlight")
            self.results_text.insert(tk.END, f"Estimated COGS: ${(sales * 0.6 if sales > 0 else 0):.2f}\n", "highlight")
            self.results_text.insert(tk.END, f"Taxable Income: ${taxable_income:.2f}\n", "highlight")
            
            combined_df = pd.concat(dfs.values(), keys=dfs.keys(), sort=False)
            combined_df["Error"] = ""
            for err in gaap_errors:
                try:
                    combined_df.loc[combined_df.index[combined_df.apply(lambda row: err.split(":")[0] in str(row), axis=1)], "Error"] = err
                except Exception as ex:
                    logging.warning(f"Error adding GAAP error to combined_df: {str(ex)}")
            output_dir = r"C:\Users\Admin\Downloads\AuditTool"
            os.makedirs(output_dir, exist_ok=True)
            combined_df.to_csv(os.path.join(output_dir, "audit_results.csv"), index=False)
            combined_df.to_excel(os.path.join(output_dir, "audit_results.xlsx"), index=False)
            self.results_text.insert(tk.END, f"\nDetailed results saved to {output_dir}\\audit_results.csv and {output_dir}\\audit_results.xlsx\n", "success")
            logging.debug("Analysis completed successfully.")
        except Exception as e:
            logging.error(f"Error in analyze: {str(e)}")
            messagebox.showerror("Error", f"Failed to analyze data: {str(e)}")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = AuditApp(root)
        root.mainloop()
        logging.debug("GUI application closed normally.")
    except KeyboardInterrupt:
        logging.warning("Application interrupted by user (Ctrl+C).")
        sys.exit(0)
    except Exception as e:
        logging.error(f"Unexpected error in main: {str(e)}")
        messagebox.showerror("Error", f"Application failed to start: {str(e)}")
        sys.exit(1)
