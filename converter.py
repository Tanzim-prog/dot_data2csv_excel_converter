import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import sys


# -------------------- DATA SCHEMA -------------------- #
WINE_HEADERS = [
    "Cultivars",
    "Alcohol",
    "Malic acid",
    "Ash",
    "Alcalinity of ash",
    "Magnesium",
    "Total phenols",
    "Flavanoids",
    "Nonflavanoid phenols",
    "Proanthocyanins",
    "Color intensity",
    "Hue",
    "OD280/OD315 of diluted wines",
    "Proline"
]


# -------------------- DATA LOADING LOGIC -------------------- #

def load_data_file(path: Path) -> pd.DataFrame:
    """Load .data file, auto-detect delimiter, force known Wine dataset headers."""

    sample = path.read_text(encoding="utf-8", errors="ignore")[:2048]
    delimiter = "," if "," in sample else None

    df = pd.read_csv(
        path,
        delimiter=delimiter,
        header=None,
        engine="python",
        on_bad_lines="skip"
    )

    # FORCE CORRECT HEADERS ALWAYS
    if len(df.columns) != len(WINE_HEADERS):
        raise ValueError(
            f"Dataset has {len(df.columns)} columns but Wine dataset requires {len(WINE_HEADERS)}."
        )

    df.columns = WINE_HEADERS

    return df


def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.drop_duplicates().reset_index(drop=True)
    df = df.ffill()  # Fixed: replaced fillna(method="ffill") with ffill()
    return df


def save_output(df: pd.DataFrame, output_path: Path):
    ext = output_path.suffix.lower()

    if ext == ".xlsx":
        try:
            df.to_excel(output_path, index=False, engine='openpyxl')
        except ImportError:
            # Suggest installation of openpyxl
            raise ImportError(
                "Excel export requires 'openpyxl' package.\n\n"
                "Please install it by running:\n"
                "pip install openpyxl\n\n"
                "Or save the file as CSV instead."
            )
    elif ext == ".csv":
        df.to_csv(output_path, index=False)
    else:
        raise ValueError("Unsupported output format. Please use .xlsx or .csv")


# -------------------- GUI LOGIC -------------------- #

def browse_input():
    path = filedialog.askopenfilename(
        title="Select .data File",
        filetypes=[("Data File", "*.data"), ("All Files", "*.*")]
    )
    if path:
        input_var.set(path)


def browse_output():
    path = filedialog.asksaveasfilename(
        title="Save Output File",
        defaultextension=".xlsx",
        filetypes=[
            ("Excel", "*.xlsx"), 
            ("CSV", "*.csv")
        ]
    )
    if path:
        output_var.set(path)


def convert():
    input_path = input_var.get()
    output_path = output_var.get()

    if not input_path:
        messagebox.showerror("Error", "Select an input .data file.")
        return
    if not output_path:
        messagebox.showerror("Error", "Choose an output file path.")
        return

    try:
        progress_bar["value"] = 20
        root.update_idletasks()

        df = load_data_file(Path(input_path))
        progress_bar["value"] = 60
        root.update_idletasks()

        df = clean_df(df)
        progress_bar["value"] = 80
        root.update_idletasks()

        save_output(df, Path(output_path))
        progress_bar["value"] = 100

        style.configure("Green.Horizontal.TProgressbar", troughcolor="#d0ffd0", background="#00b300")
        progress_bar.configure(style="Green.Horizontal.TProgressbar")

        messagebox.showinfo("Success", "File converted successfully.")

    except ImportError as e:
        # Special handling for missing openpyxl
        error_msg = str(e)
        if "openpyxl" in error_msg:
            error_msg += "\n\nYou can still save as CSV format."
        messagebox.showerror("Conversion Failed", error_msg)
        progress_bar["value"] = 0
    except Exception as e:
        messagebox.showerror("Conversion Failed", str(e))
        progress_bar["value"] = 0


def shutdown():
    root.destroy()


# -------------------- BUILD UI -------------------- #

root = tk.Tk()
root.title("Data → CSV/XLSX Converter")
root.geometry("700x340")  # Slightly taller for better error visibility
root.configure(bg="#f4f4f4")
root.resizable(False, False)

style = ttk.Style()
style.theme_use("clam")

style.configure("TButton", font=("Segoe UI", 11), padding=6)
style.configure("TEntry", padding=4)
style.configure("TLabel", background="#f4f4f4", font=("Segoe UI", 11))
style.configure("Horizontal.TProgressbar", thickness=10)

input_var = tk.StringVar()
output_var = tk.StringVar()

title_lbl = tk.Label(root, text="DATA → CSV/XLSX Converter", font=("Segoe UI", 16, "bold"), bg="#f4f4f4")
title_lbl.pack(pady=15)

# Info label about requirements
info_text = "Note: Excel export requires 'openpyxl' package (pip install openpyxl)"
info_lbl = tk.Label(root, text=info_text, font=("Segoe UI", 9), bg="#f4f4f4", fg="#666666")
info_lbl.pack(pady=(0, 10))

frame = tk.Frame(root, bg="#f4f4f4")
frame.pack(pady=10)

ttk.Label(frame, text="Input (.data):").grid(row=0, column=0, sticky="w")
ttk.Entry(frame, textvariable=input_var, width=55).grid(row=0, column=1, padx=10)
ttk.Button(frame, text="Browse", command=browse_input).grid(row=0, column=2)

ttk.Label(frame, text="Output (csv/xlsx):").grid(row=1, column=0, sticky="w", pady=15)
ttk.Entry(frame, textvariable=output_var, width=55).grid(row=1, column=1, padx=10)
ttk.Button(frame, text="Save As", command=browse_output).grid(row=1, column=2)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=500, mode="determinate")
progress_bar.pack(pady=15)

button_frame = tk.Frame(root, bg="#f4f4f4")
button_frame.pack(pady=5)

ttk.Button(button_frame, text="Convert", command=convert, width=20).grid(row=0, column=0, padx=10)
ttk.Button(button_frame, text="Exit", command=shutdown, width=20).grid(row=0, column=1, padx=10)

root.mainloop()