
## 7. One Final Python File (Covers All Day 77 Topics)
"""
Day 77 – Project: Excel Data Processor GUI
Phase 6 Capstone | 84-Day Python & Advanced Excel Mastery Roadmap
Upload Excel → Select processing option → Download result
Covers: tkinter GUI, file dialogs, Excel connect (pandas), error messages & feedback
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from datetime import datetime, timedelta
import random

class ExcelDataProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Processor GUI – Finance Edition | Day 77")
        self.root.geometry("780x620")
        self.root.resizable(False, False)

        # State
        self.input_path = tk.StringVar()
        self.output_df = None
        self.processing_options = [
            "1. Clean Missing Data (fillna)",
            "2. Remove Duplicates",
            "3. Summary Statistics (new sheet)",
            "4. Categorize Amounts (Low/Med/High)",
            "5. Add Totals Row",
            "6. Filter Amount > 0"
        ]

        self._build_ui()
        self._create_sample_if_missing()

    def _build_ui(self):
        # Header
        header = tk.Label(
            self.root,
            text="Excel Data Processor GUI",
            font=("Segoe UI", 18, "bold"),
            fg="#1a5f2a"
        )
        header.pack(pady=(15, 5))

        sub = tk.Label(
            self.root,
            text="Upload → Select Option → Process → Download   |   Day 77 Project",
            font=("Segoe UI", 10),
            fg="#555"
        )
        sub.pack(pady=(0, 15))

        # --- File Selection Frame ---
        file_frame = ttk.LabelFrame(self.root, text="1. Select Input Excel File", padding=10)
        file_frame.pack(fill="x", padx=20, pady=5)

        tk.Entry(file_frame, textvariable=self.input_path, width=70, state="readonly").pack(side="left", padx=(0, 8))
        ttk.Button(file_frame, text="Browse…", command=self.browse_file).pack(side="left")

        # --- Processing Option ---
        opt_frame = ttk.LabelFrame(self.root, text="2. Choose Processing Option", padding=10)
        opt_frame.pack(fill="x", padx=20, pady=8)

        self.option_var = tk.StringVar(value=self.processing_options[0])
        self.combo = ttk.Combobox(
            opt_frame,
            textvariable=self.option_var,
            values=self.processing_options,
            state="readonly",
            width=55
        )
        self.combo.pack(side="left", padx=(0, 10))

        # --- Action Buttons ---
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)

        self.process_btn = ttk.Button(btn_frame, text="▶  Process File", command=self.process_file)
        self.process_btn.pack(side="left", padx=6)

        self.save_btn = ttk.Button(btn_frame, text="💾  Save Result As…", command=self.save_result, state="disabled")
        self.save_btn.pack(side="left", padx=6)

        ttk.Button(btn_frame, text="📄 Create Sample Data", command=self.create_sample_data).pack(side="left", padx=6)

        # --- Progress & Status ---
        self.progress = ttk.Progressbar(self.root, mode="indeterminate", length=700)
        self.progress.pack(pady=(5, 5))

        self.status_var = tk.StringVar(value="Ready. Select a file and processing option.")
        status_label = tk.Label(
            self.root,
            textvariable=self.status_var,
            font=("Segoe UI", 10),
            fg="#0a5",
            anchor="w"
        )
        status_label.pack(fill="x", padx=25)

        # --- Log Area ---
        log_frame = ttk.LabelFrame(self.root, text="Activity Log", padding=5)
        log_frame.pack(fill="both", expand=True, padx=20, pady=10)

        self.log_text = tk.Text(log_frame, height=12, font=("Consolas", 9), state="disabled", wrap="word")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def log(self, message: str):
        """Append a timestamped line to the log Text widget."""
        self.log_text.configure(state="normal")
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.root.update_idletasks()

    def set_status(self, text: str, color: str = "#0a5"):
        self.status_var.set(text)
        # simple color change via config is limited; we keep it green/red via message

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self.input_path.set(path)
            self.log(f"Selected file: {path}")
            self.set_status(f"File loaded: {os.path.basename(path)}")
            self.save_btn.config(state="disabled")
            self.output_df = None

    def _create_sample_if_missing(self):
        if not os.path.exists("sample_expenses.xlsx"):
            self.create_sample_data(silent=True)

    def create_sample_data(self, silent=False):
        """Generate a realistic finance expense Excel for testing."""
        try:
            categories = ["Travel", "Meals", "Software", "Office", "Marketing", "Utilities"]
            rows = []
            start = datetime(2025, 1, 1)
            for i in range(40):
                date = start + timedelta(days=random.randint(0, 120))
                cat = random.choice(categories)
                amount = round(random.uniform(15, 850), 2)
                # inject some missing & duplicates for cleaning demos
                if i % 9 == 0:
                    amount = None
                desc = f"{cat} expense #{i+1}"
                rows.append({
                    "Date": date.strftime("%Y-%m-%d"),
                    "Description": desc,
                    "Category": cat,
                    "Amount": amount,
                    "Paid": random.choice(["Yes", "No", "Yes"])
                })
            # force two duplicate rows
            rows.append(rows[5].copy())
            rows.append(rows[12].copy())

            df = pd.DataFrame(rows)
            df.to_excel("sample_expenses.xlsx", index=False)
            self.input_path.set(os.path.abspath("sample_expenses.xlsx"))
            msg = "sample_expenses.xlsx created successfully (40 rows + intentional missing/duplicates)."
            self.log(msg)
            if not silent:
                messagebox.showinfo("Sample Created", msg)
                self.set_status("Sample data ready.")
        except Exception as e:
            messagebox.showerror("Sample Error", str(e))
            self.log(f"ERROR creating sample: {e}")

    def process_file(self):
        path = self.input_path.get().strip()
        if not path:
            messagebox.showwarning("No File", "Please select an Excel file first.")
            return
        if not os.path.isfile(path):
            messagebox.showerror("File Not Found", f"The file does not exist:\n{path}")
            return

        option = self.option_var.get()
        self.process_btn.config(state="disabled")
        self.save_btn.config(state="disabled")
        self.progress.start(12)
        self.set_status("Processing… please wait")
        self.log(f"Starting: {option}")
        self.root.update()

        try:
            df = pd.read_excel(path)
            self.log(f"Loaded {len(df)} rows × {len(df.columns)} columns")

            if option.startswith("1."):
                # Clean missing
                before = df.isna().sum().sum()
                df = df.fillna({"Amount": 0, "Description": "Unknown", "Category": "Uncategorized", "Paid": "No"})
                after = df.isna().sum().sum()
                self.log(f"Filled {before - after} missing values")
                self.output_df = df

            elif option.startswith("2."):
                before = len(df)
                df = df.drop_duplicates()
                self.log(f"Removed {before - len(df)} duplicate rows")
                self.output_df = df

            elif option.startswith("3."):
                # Summary statistics on numeric columns + value counts
                numeric = df.select_dtypes(include="number")
                summary = numeric.describe().T
                summary["missing"] = df[numeric.columns].isna().sum()
                # also a simple category breakdown if present
                self.output_df = summary  # will be written as single sheet; user can extend
                self.log("Generated descriptive statistics for numeric columns")

            elif option.startswith("4."):
                if "Amount" not in df.columns:
                    raise ValueError("Column 'Amount' is required for categorization.")
                def bucket(x):
                    if pd.isna(x):
                        return "Missing"
                    if x < 100:
                        return "Low"
                    if x < 400:
                        return "Medium"
                    return "High"
                df["Amount_Category"] = df["Amount"].apply(bucket)
                self.log("Added column Amount_Category (Low / Medium / High)")
                self.output_df = df

            elif option.startswith("5."):
                if "Amount" not in df.columns:
                    raise ValueError("Column 'Amount' is required for totals.")
                totals = {col: "" for col in df.columns}
                totals["Description"] = "TOTAL"
                totals["Amount"] = df["Amount"].sum(skipna=True)
                df = pd.concat([df, pd.DataFrame([totals])], ignore_index=True)
                self.log(f"Appended totals row. Sum(Amount) = {totals['Amount']:.2f}")
                self.output_df = df

            elif option.startswith("6."):
                if "Amount" not in df.columns:
                    raise ValueError("Column 'Amount' is required for filtering.")
                before = len(df)
                df = df[df["Amount"] > 0].copy()
                self.log(f"Filtered to Amount > 0: {before} → {len(df)} rows")
                self.output_df = df

            else:
                raise ValueError("Unknown option selected.")

            self.progress.stop()
            self.set_status("Processing complete. You can now save the result.")
            self.log("SUCCESS – result ready in memory. Click 'Save Result As…'")
            self.save_btn.config(state="normal")
            messagebox.showinfo("Done", "Processing finished successfully.\nClick 'Save Result As…' to download.")

        except Exception as e:
            self.progress.stop()
            self.set_status("Error occurred – see log")
            self.log(f"ERROR: {e}")
            messagebox.showerror("Processing Error", f"An error occurred:\n\n{e}")
        finally:
            self.process_btn.config(state="normal")

    def save_result(self):
        if self.output_df is None:
            messagebox.showwarning("Nothing to Save", "Process a file first.")
            return
        path = filedialog.asksaveasfilename(
            title="Save Processed Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="processed_result.xlsx"
        )
        if not path:
            return
        try:
            self.output_df.to_excel(path, index=True if self.option_var.get().startswith("3.") else False)
            self.log(f"Saved → {path}")
            self.set_status(f"Saved: {os.path.basename(path)}")
            messagebox.showinfo("Saved", f"File saved successfully:\n{path}")
        except Exception as e:
            messagebox.showerror("Save Error", str(e))
            self.log(f"SAVE ERROR: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelDataProcessorGUI(root)
    root.mainloop()