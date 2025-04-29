import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import numpy as np
from scipy import stats
import warnings

warnings.filterwarnings("ignore", category=RuntimeWarning)

def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file_path:
        return
    try:
        global df
        df = pd.read_excel(file_path)
        numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
        if not numeric_cols:
            messagebox.showerror("Error", "No numeric columns found.")
            return
        column_dropdown['values'] = numeric_cols
        column_dropdown.current(0)
        calculate_button.config(state=tk.NORMAL)
    except Exception as e:
        messagebox.showerror("Error", str(e))

def calculate_stats():
    selected_col = column_var.get()
    data = df[selected_col].dropna()

    if data.empty:
        messagebox.showerror("Error", "Selected column has no valid numeric data.")
        return

    mean = np.mean(data)
    median = np.median(data)
    mode = stats.mode(data, keepdims=False).mode
    p25 = np.percentile(data, 25)
    p75 = np.percentile(data, 75)
    moment2 = stats.moment(data, moment=2)
    moment3 = stats.moment(data, moment=3)

    result = (
        f"Analyzing: {selected_col}\n\n"
        f"Mean: {mean:.2f}\n"
        f"Median: {median:.2f}\n"
        f"Mode: {mode}\n"
        f"25th Percentile: {p25:.2f}\n"
        f"75th Percentile: {p75:.2f}\n"
        f"2nd Central Moment: {moment2:.2f}\n"
        f"3rd Central Moment: {moment3:.2f}"
    )

    result_box.config(state='normal')
    result_box.delete(1.0, tk.END)
    result_box.insert(tk.END, result)
    result_box.config(state='disabled')

# --- Styling function for hover ---
def on_enter(e):
    e.widget['background'] = '#3e8e41'

def on_leave(e):
    e.widget['background'] = '#4CAF50'

# --- Main Window ---
root = tk.Tk()
root.title("Business Management System ┃ Developed by Grp-No.-102")
root.geometry("600x500")
root.configure(bg="#f9f9f9")

# --- Header ---
header = tk.Label(root, text="Excel Analyzer", font=("Segoe UI", 20, "bold"), bg="#f9f9f9", fg="#333")
header.pack(pady=15)

# --- File Frame ---
frame = tk.Frame(root, bg="white", bd=1, relief="solid")
frame.pack(padx=20, pady=10, fill="x")

load_btn = tk.Button(frame, text="Choose Excel File", font=("Segoe UI", 12), bg="#4CAF50", fg="white", bd=0, pady=8, command=load_file)
load_btn.pack(pady=10, padx=10)
load_btn.bind("<Enter>", on_enter)
load_btn.bind("<Leave>", on_leave)

# --- Column dropdown ---
column_var = tk.StringVar()
column_dropdown = ttk.Combobox(root, textvariable=column_var, state="readonly", font=("Segoe UI", 11))
column_dropdown.pack(pady=10)

# --- Calculate Button ---
calculate_button = tk.Button(root, text="Calculate Stats", font=("Segoe UI", 12), bg="#2196F3", fg="white", bd=0, pady=8, state=tk.DISABLED, command=calculate_stats)
calculate_button.pack(pady=10)

# --- Results Frame ---
result_frame = tk.Frame(root, bg="white", bd=1, relief="solid")
result_frame.pack(padx=20, pady=10, fill="both", expand=True)

result_box = tk.Text(result_frame, wrap="word", font=("Consolas", 11), state='disabled', bg="#ffffff", fg="#222")
result_box.pack(fill="both", expand=True, padx=10, pady=10)

# Run app 
root.mainloop()