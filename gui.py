import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# Initialize window
root = tk.Tk()
root.title("Data Cleaning Automation Tool")
root.geometry("500x450")

# Variables
input_file = tk.StringVar()
output_folder = tk.StringVar()

remove_duplicates = tk.BooleanVar()
remove_blanks = tk.BooleanVar()
trim_spaces = tk.BooleanVar()
title_case = tk.BooleanVar()

# -----------------------------
# Functions
# -----------------------------

def browse_input():
    file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file:
        input_file.set(file)

def browse_output():
    folder = filedialog.askdirectory()
    if folder:
        output_folder.set(folder)

def clean_data():
    if not input_file.get() or not output_folder.get():
        messagebox.showerror("Error", "Please select input file and output folder")
        return

    try:
        df = pd.read_excel(input_file.get())

        # Remove duplicates
        if remove_duplicates.get():
            df = df.drop_duplicates()

        # Remove blank rows
        if remove_blanks.get():
            df = df.dropna(how='all')

        # Process text columns only
        text_cols = df.select_dtypes(include='object').columns

        # Trim spaces
        if trim_spaces.get():
            for col in text_cols:
                df[col] = df[col].astype(str).str.strip()

        # Title case
        if title_case.get():
            for col in text_cols:
                df[col] = df[col].astype(str).str.title()

        # Save file
        output_path = os.path.join(output_folder.get(), "cleaned_data.xlsx")
        df.to_excel(output_path, index=False)

        messagebox.showinfo("Success", f"File saved at:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))

def reset_paths():
    input_file.set("")
    output_folder.set("")

def reset_checkboxes():
    remove_duplicates.set(False)
    remove_blanks.set(False)
    trim_spaces.set(False)
    title_case.set(False)

# -----------------------------
# UI Layout
# -----------------------------

tk.Label(root, text="Select Input Excel File").pack(pady=5)
tk.Entry(root, textvariable=input_file, width=50).pack()
tk.Button(root, text="Browse", command=browse_input).pack(pady=5)

tk.Label(root, text="Select Output Folder").pack(pady=5)
tk.Entry(root, textvariable=output_folder, width=50).pack()
tk.Button(root, text="Browse", command=browse_output).pack(pady=5)

tk.Label(root, text="Select Cleaning Options").pack(pady=10)

tk.Checkbutton(root, text="Remove Duplicate Rows", variable=remove_duplicates).pack(anchor='w')
tk.Checkbutton(root, text="Remove Blank Rows", variable=remove_blanks).pack(anchor='w')
tk.Checkbutton(root, text="Trim Spaces (Text Columns)", variable=trim_spaces).pack(anchor='w')
tk.Checkbutton(root, text="Convert to Title Case", variable=title_case).pack(anchor='w')

tk.Button(root, text="Clean Data", bg="green", fg="white", command=clean_data).pack(pady=10)
tk.Button(root, text="Reset Paths", command=reset_paths).pack(pady=5)
tk.Button(root, text="Reset Checkboxes", command=reset_checkboxes).pack(pady=5)

# Run app
root.mainloop()