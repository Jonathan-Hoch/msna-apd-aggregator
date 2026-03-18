# Aggregates per-file MSNA APD outputs into a single master Excel sheet.
# Computes burst-level statistics (APs/burst, clusters/burst, amplitude, AP frequency)
# and flags outliers by SD. Includes a GUI for point-and-click use.
#
# To run headless: fill in INPUT_DIR and OUTPUT_FILE below, then call process_data() directly.
# To use the GUI: run as-is.

import pandas as pd
import openpyxl
import os
import customtkinter as ctk
from tkinter import filedialog, messagebox
from openpyxl.styles import PatternFill, Font, Alignment
from pathlib import Path
from datetime import datetime
import numpy as np

VERSION = "1.2.0"

# ---- config (headless use) ----
# INPUT_DIR   = Path(r"C:\path\to\your\APD files")
# OUTPUT_FILE = Path(r"C:\path\to\output\APD_Master.xlsx")
#
# To run without the GUI:
#   1. Set the paths above
#   2. Call: process_data(INPUT_DIR, OUTPUT_FILE.parent)
#   3. Delete or ignore everything below the process_data() function


def process_data(input_path, output_path):
    input_dir = Path(input_path)
    output_dir = Path(output_path)

    output_file = output_dir / "APD_Master.xlsx"
    log_file    = output_dir / "processing_log.txt"

    rows = []
    labels = None
    log_entries = [f"MSNA Pipeline v{VERSION} - Run Date: {datetime.now()}\n", "-"*50]

    excel_files = sorted(input_dir.glob("*.xlsx")) + sorted(input_dir.glob("*.xls"))
    if not excel_files:
        raise ValueError("No Excel files found in the selected directory.")

    for filepath in excel_files:
        try:
            if filepath.name.startswith("~$"):
                continue

            df1 = pd.read_excel(filepath, sheet_name=0, header=0)
            if labels is None:
                labels = df1.columns.tolist()

            if len(df1) >= 1:
                data_row = df1.iloc[0].tolist()
                file_duration_sec = df1.iloc[0, 0]
            else:
                data_row = [None] * len(labels)
                file_duration_sec = None

            try:
                df2 = pd.read_excel(filepath, sheet_name=1, header=0)

                col_amp    = df2.iloc[:, 1]   # Column B - Amplitude
                col_spikes = df2.iloc[:, 5]   # Column F - Spikes per Burst
                col_bins   = df2.iloc[:, 9]   # Column J - Bins per Burst

                amp_vals    = pd.to_numeric(col_amp,    errors='coerce').dropna()
                spikes_vals = pd.to_numeric(col_spikes, errors='coerce').dropna()
                bins_vals   = pd.to_numeric(col_bins,   errors='coerce').dropna()

                burst_amplitude    = amp_vals.mean()    if len(amp_vals)    > 0 else None
                aps_per_burst      = spikes_vals.mean() if len(spikes_vals) > 0 else None
                clusters_per_burst = bins_vals.mean()   if len(bins_vals)   > 0 else None
                total_spikes       = spikes_vals.sum()  if len(spikes_vals) > 0 else None

                if total_spikes is not None and file_duration_sec not in (None, 0):
                    ap_frequency = (total_spikes / file_duration_sec) * 60
                else:
                    ap_frequency = None

            except Exception as e2:
                print(f"  Sheet 2 error in {filepath.name}: {e2}")
                aps_per_burst = clusters_per_burst = burst_amplitude = ap_frequency = None

            rows.append(
                [filepath.stem] + data_row +
                [aps_per_burst, clusters_per_burst, burst_amplitude, ap_frequency]
            )
            log_entries.append(f"SUCCESS: {filepath.name}")

        except Exception as e:
            log_entries.append(f"ERROR {filepath.name}: {e}")

    cols = ["File Name"] + labels + [
        "APs per Burst (Rate Coding)",
        "AP Clusters per Burst (Recruitment)",
        "Burst Amplitude (V)",
        "AP Frequency (APs/min)"
    ]

    # ---- statistical QA ----
    stat_cols = [
        "APs per Burst (Rate Coding)",
        "AP Clusters per Burst (Recruitment)",
        "Burst Amplitude (V)",
        "AP Frequency (APs/min)"
    ]
    master_df = pd.DataFrame(rows, columns=cols)
    stats_map = {}
    for col in stat_cols:
        master_df[col] = pd.to_numeric(master_df[col], errors='coerce')
        if not master_df[col].dropna().empty:
            stats_map[col] = {
                'mean': master_df[col].mean(),
                'std':  master_df[col].std()
            }

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Master"

    ws.append(cols)

    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", start_color="2F4F8F")
    for col_idx in range(1, len(cols) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    for row in rows:
        ws.append(row)
        for col_idx in range(1, len(row) + 1):
            ws.cell(row=ws.max_row, column=col_idx).font = Font(name="Arial")

    for col in ws.columns:
        max_len = max((len(str(cell.value)) if cell.value else 0) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)

    # ---- color coding (SD flagging) ----
    fills = {
        "yellow": PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid"),
        "orange": PatternFill(start_color="FFCC00", end_color="FFCC00", fill_type="solid"),
        "red":    PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    }

    for r_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
        for c_idx, cell in enumerate(row, start=1):
            col_name = cols[c_idx - 1]
            if col_name in stat_cols and cell.value is not None:
                try:
                    val = float(cell.value)
                    if col_name in stats_map:
                        m = stats_map[col_name]['mean']
                        s = stats_map[col_name]['std']
                        if s and s > 0:
                            z = abs((val - m) / s)
                            if z >= 3:
                                cell.fill = fills["red"]
                            elif z >= 2:
                                cell.fill = fills["orange"]
                            elif z >= 1:
                                cell.fill = fills["yellow"]
                except:
                    continue

    wb.save(output_file)

    with open(log_file, "w") as f:
        f.write("\n".join(log_entries))

    return output_file


# ---- GUI (delete from here down if running headless) ----

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class MSNAApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"MSNA APD Aggregator v{VERSION}")
        self.geometry("650x480")
        self.configure(fg_color="#1a1a1a")
        self.grid_columnconfigure(0, weight=1)

        self.title_label = ctk.CTkLabel(self, text="MSNA Data Pipeline", font=ctk.CTkFont(size=26, weight="bold"))
        self.title_label.grid(row=0, column=0, pady=(30, 20))

        self.label1 = ctk.CTkLabel(self, text="1. Select Raw Data Folder:", font=ctk.CTkFont(size=14))
        self.label1.grid(row=1, column=0, padx=40, sticky="w")
        self.entry1 = ctk.CTkEntry(self, placeholder_text="Folder path...", width=450, height=35, corner_radius=10)
        self.entry1.grid(row=2, column=0, padx=40, pady=(5, 15), sticky="w")
        self.browse1 = ctk.CTkButton(self, text="Browse", width=80, height=35, corner_radius=10, command=self.get_input, fg_color="#3d3d3d")
        self.browse1.grid(row=2, column=0, padx=(500, 40), pady=(5, 15))

        self.label2 = ctk.CTkLabel(self, text="2. Select Export Destination:", font=ctk.CTkFont(size=14))
        self.label2.grid(row=3, column=0, padx=40, sticky="w")
        self.entry2 = ctk.CTkEntry(self, placeholder_text="Save destination...", width=450, height=35, corner_radius=10)
        self.entry2.grid(row=4, column=0, padx=40, pady=(5, 15), sticky="w")
        self.browse2 = ctk.CTkButton(self, text="Browse", width=80, height=35, corner_radius=10, command=self.get_output, fg_color="#3d3d3d")
        self.browse2.grid(row=4, column=0, padx=(500, 40), pady=(5, 15))

        self.run_button = ctk.CTkButton(self, text="RUN ANALYSIS", fg_color="#1f6aa5", hover_color="#144870",
                                        font=ctk.CTkFont(size=18, weight="bold"), height=55, width=300,
                                        corner_radius=15, command=self.execute)
        self.run_button.grid(row=5, column=0, pady=40)

    def get_input(self):
        folder = filedialog.askdirectory()
        if folder:
            self.entry1.delete(0, "end")
            self.entry1.insert(0, folder)

    def get_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.entry2.delete(0, "end")
            self.entry2.insert(0, folder)

    def show_custom_success(self, filename):
        pop = ctk.CTkToplevel(self)
        pop.title("Analysis Complete")
        pop.geometry("400x220")
        pop.attributes("-topmost", True)
        pop.configure(fg_color="#242424")
        pop.focus_set()
        pop.columnconfigure(0, weight=1)

        ctk.CTkLabel(pop, text="SUCCESS!", font=ctk.CTkFont(size=20, weight="bold"), text_color="#2d8a4e").pack(pady=(25, 10))
        ctk.CTkLabel(pop, text=f"Master Sheet Created:\n{filename}", font=ctk.CTkFont(size=13)).pack(pady=10)
        ctk.CTkButton(pop, text="OK", width=120, height=35, corner_radius=10, command=pop.destroy).pack(pady=20)

    def execute(self):
        in_p = self.entry1.get()
        out_p = self.entry2.get()
        if not in_p or not out_p:
            messagebox.showwarning("Missing Info", "Please select both folders.")
            return
        try:
            res = process_data(in_p, out_p)
            self.show_custom_success(os.path.basename(res))
            os.startfile(os.path.dirname(res))
        except Exception as e:
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    app = MSNAApp()
    app.mainloop()
