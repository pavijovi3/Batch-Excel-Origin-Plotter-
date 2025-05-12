import sys
import os
import threading
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar

import originpro as op
import win32com.client

# ——————————————————————————————————————————————————————————————
# Shutdown hook: ensure Origin closes if an exception occurs
# ——————————————————————————————————————————————————————————————
def origin_shutdown_exception_hook(exctype, value, traceback):
    try:
        op.exit()
    except:
        pass
    sys.__excepthook__(exctype, value, traceback)

sys.excepthook = origin_shutdown_exception_hook

# ——————————————————————————————————————————————————————————————
# Core processing function for a single file
# ——————————————————————————————————————————————————————————————
def process_file(xlsx_path, template_path, progress_callback=None):
    # 1) Read and clean the Record Sheet
    excel = pd.ExcelFile(xlsx_path)
    raw = (
        excel
        .parse('Record Sheet', header=[0,1])
        .drop(index=0)
        .reset_index(drop=True)
    )

    # 2) Build wide CSV (drop Capacity columns, keep SpeCap & Voltage)
    cycles = len(raw.columns) // 3
    out = {}
    for i in range(cycles):
        spe = pd.to_numeric(raw.iloc[:, i*3+1], errors='coerce')
        vol = pd.to_numeric(raw.iloc[:, i*3+2], errors='coerce')
        out[f"Cycle{i+1}_SpeCap"]  = spe
        out[f"Cycle{i+1}_Voltage"] = vol

    csv_path = os.path.splitext(xlsx_path)[0] + "_plotdata.csv"
    pd.DataFrame(out).to_csv(csv_path, index=False)

    # 3) Launch and show Origin
    if op.oext:
        op.new()
        op.set_show(True)
    origin_app = win32com.client.Dispatch("Origin.ApplicationSI")
    origin_app.Visible = True
    try:
        origin_app.Execute("win -r")  # restore main window
    except:
        pass

    # 4) Import CSV into a single worksheet
    wks = op.new_sheet()
    wks.from_file(csv_path, False)

    # 5) Set worksheet axes to XY, XY, XY...
    for i in range(cycles):
        wks.cols_axis('X', i*2)
        wks.cols_axis('Y', i*2 + 1)

    # 6) Create graph from template
    gr = op.new_graph(template=template_path)
    layer = gr[0]

    # 7) Plot each cycle individually, naming each trace
    for i in range(cycles):
        colx = i*2       # even columns = SpeCap
        coly = i*2 + 1   # odd  columns = Voltage
        plot = layer.add_plot(wks, coly=coly, colx=colx)
        plot.name = f"Cycle {i+1}"
        if progress_callback:
            progress_callback(100 // cycles)

    # 8) Finalize plot
    layer.rescale()
    op.lt_exec('win -s T')  # tile all windows

    # 9) Add Notes window
    nt = op.new_notes()
    nt.append("Excel input:   " + os.path.basename(xlsx_path))
    nt.append("Template used: " + os.path.basename(template_path))
    nt.append("CSV export:    " + os.path.basename(csv_path))
    nt.view = 1

    # 10) Auto-save Origin project next to the Excel file
    proj_path = os.path.abspath(os.path.splitext(xlsx_path)[0] + ".opju")
    op.save(proj_path)

# ——————————————————————————————————————————————————————————————
# GUI setup for batch processing
# ——————————————————————————————————————————————————————————————
def run_gui():
    root = tk.Tk()
    root.title("Batch Excel → Origin Plotter")

    # Excel files picker (multiple)
    tk.Label(root, text="1) Select Excel files:").pack(anchor='w', padx=10, pady=(10,0))
    files_ent = tk.Entry(root, width=60)
    files_ent.pack(padx=10)
    tk.Button(root, text="Browse… (multi-select)", bg='lightblue',
              command=lambda: _browse_files(files_ent)
    ).pack(pady=(0,5))

    # Origin template picker
    tk.Label(root, text="2) Select Origin template:").pack(anchor='w', padx=10)
    templ_ent = tk.Entry(root, width=60)
    templ_ent.pack(padx=10)
    tk.Button(root, text="Browse…", bg='lightblue',
              command=lambda: _browse(templ_ent, [("Origin templates","*.otpu;*.otp")])
    ).pack(pady=(0,10))

    # Generate & Plot button
    tk.Button(root, text="Generate & Plot All", width=30, bg='lightblue',
              command=lambda: _start_batch(files_ent.get(), templ_ent.get())
    ).pack(pady=5)

    # Exit button
    tk.Button(root, text="Exit", width=30, bg='coral', command=root.destroy).pack(pady=(0,10))

    # Progress bar
    progress = Progressbar(root, length=400, mode='determinate')
    progress.pack(pady=(5,15))

    # Utility functions
    def update_progress(v):
        progress['value'] = min(100, progress['value'] + v)
        root.update_idletasks()

    def _browse(entry, ftypes):
        p = filedialog.askopenfilename(filetypes=ftypes)
        if p:
            entry.delete(0, tk.END)
            entry.insert(0, p)

    def _browse_files(entry):
        paths = filedialog.askopenfilenames(filetypes=[("Excel files","*.xlsx")])
        if paths:
            entry.delete(0, tk.END)
            entry.insert(0, ";".join(paths))

    def _start_batch(files_str, templ):
        if not files_str:
            return messagebox.showerror("Error", "Select at least one Excel file.")
        files = files_str.split(';')
        for f in files:
            if not os.path.exists(f):
                return messagebox.showerror("Error", f"File not found: {f}")
        if not os.path.exists(templ):
            return messagebox.showerror("Error", "Select a valid template.")
        progress['value'] = 0
        threading.Thread(target=lambda: _worker_batch(files, templ), daemon=True).start()

    def _worker_batch(files, templ):
        try:
            total_files = len(files)
            for idx, f in enumerate(files, start=1):
                # Reset per-file progress
                progress['value'] = 0
                process_file(f, templ, progress_callback=update_progress)
            messagebox.showinfo("Done", f"Processed {total_files} files successfully.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    root.mainloop()

if __name__ == "__main__":
    run_gui()
