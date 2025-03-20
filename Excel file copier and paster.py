import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import openpyxl
import csv
import re
import xlwings as xw

# -------------------------------------------------------------------
# 1) Define the list of variables in the desired order
# 2) Create a dictionary mapping each variable to its cell reference
# -------------------------------------------------------------------
VARIABLES = [
    "mean_relative_frontal_alpha_power",
    "mean_abs_frontal_alpha_power",
    "mean_relative_frontal_beta_power",
    "mean_abs_frontal_beta_power",
    "mean_relative_frontal_theta_power",
    "mean_abs_frontal_theta_power",
    "mean_relative_parietal_alpha_power",
    "mean_abs_parietal_alpha_power",
    "mean_relative_parietal_beta_power",
    "mean_abs_parietal_beta_power",
    "mean_relative_parietal_theta_power",
    "mean_abs_parietal_theta_power",
    "mean_relative_occipital_alpha_power",
    "mean_abs_occipital_alpha_power",
    "mean_relative_occipital_beta_power",
    "mean_abs_occipital_beta_power",
    "mean_relative_occipital_theta_power",
    "mean_abs_occipital_theta_power",
    "mean_relative_central_alpha_power",
    "mean_abs_central_alpha_power",
    "mean_relative_central_beta_power",
    "mean_abs_central_beta_power",
    "mean_relative_central_theta_power",
    "mean_abs_central_theta_power",
    "mean_relative_t7_alpha_power",
    "mean_abs_t7_alpha_power",
    "mean_relative_t8_alpha_power",
    "mean_abs_t8_alpha_power"
]

variable_cell_map = {}
start_row_for_input = 2  # We start at B2 for the first variable
for i, var_name in enumerate(VARIABLES):
    variable_cell_map[var_name] = f"B{start_row_for_input + i}"
# -------------------------------------------------------------------

def select_folder():
    """Opens a file dialog to select the subfolder."""
    folder_path = filedialog.askdirectory()
    folder_path_entry.delete(0, tk.END)
    folder_path_entry.insert(0, folder_path)

def select_output_excel():
    """Opens a file dialog to select the existing output Excel file."""
    output_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    output_file_entry.delete(0, tk.END)
    output_file_entry.insert(0, output_file_path)

def extract_and_write():
    """Extracts data from the chosen variable (mapped to a cell) in each input file,
    verifies that each expected file (from Column A in the output workbook) is present,
    and updates the existing Excel workbook using xlwings (so that charts and other
    Excel features are preserved)."""
    folder_path = folder_path_entry.get()
    output_file_path = output_file_entry.get()
    chosen_variable = variable_var.get()  # The selected variable from the dropdown
    output_col_letters = output_col_entry.get().upper().strip()
    row_range_str = output_row_range_entry.get().strip()

    if not all([folder_path, output_file_path, chosen_variable, output_col_letters, row_range_str]):
        messagebox.showerror("Error", "Please fill in all fields.")
        return

    # Check that the output file exists
    if not os.path.exists(output_file_path):
        messagebox.showerror("Error", f"Output file not found: {output_file_path}")
        return

    # Look up the corresponding cell reference for the chosen variable
    cell_to_extract = variable_cell_map.get(chosen_variable, None)
    if not cell_to_extract:
        messagebox.showerror("Error", f"No cell mapping found for {chosen_variable}")
        return

    # Parse the row range (e.g. "2:12") into start_row and end_row
    try:
        start_row_str, end_row_str = row_range_str.split(':')
        start_row = int(start_row_str)
        end_row = int(end_row_str)
    except ValueError:
        messagebox.showerror("Error", "Invalid row range format. Use something like '2:12'.")
        return

    try:
        # Open the existing workbook using xlwings
        output_wb = xw.Book(output_file_path)
        output_sheet = output_wb.sheets[0]

        # -----------------------------------------------------------
        # New functionality: Verify that files exist for each expected identifier.
        # Read the expected identifiers from Column A in the given row range.
        # -----------------------------------------------------------
        expected_ids = output_sheet.range(f"A{start_row}:A{end_row}").value
        if not isinstance(expected_ids, list):
            expected_ids = [expected_ids]
        # Get all files in the folder with supported extensions.
        folder_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.xlsx', '.xls', '.csv'))]
        mapping = {}
        missing_ids = []
        for identifier in expected_ids:
            if identifier is None or str(identifier).strip() == "":
                missing_ids.append("Empty cell")
                continue
            # Search for a file whose name contains the identifier string.
            matches = [f for f in folder_files if str(identifier) in f]
            if not matches:
                missing_ids.append(str(identifier))
            else:
                mapping[str(identifier)] = matches[0]  # take the first match

        if missing_ids:
            messagebox.showerror("Error", f"Missing files for identifiers: {', '.join(missing_ids)}")
            output_wb.close()
            return
        # -----------------------------------------------------------
        # Now process each row based on the expected identifiers.
        # -----------------------------------------------------------
        for i, identifier in enumerate(expected_ids):
            current_row = start_row + i
            # Use the mapping to get the corresponding file.
            file_name = mapping.get(str(identifier))
            full_file_path = os.path.join(folder_path, file_name)
            try:
                # For Excel files, use openpyxl to extract the value
                if file_name.lower().endswith(('.xlsx', '.xls')):
                    input_wb = openpyxl.load_workbook(full_file_path, data_only=True)
                    input_sheet = input_wb.active
                    extracted_value = input_sheet[cell_to_extract].value
                    input_wb.close()
                # For CSV files
                else:
                    with open(full_file_path, 'r', newline='', encoding='utf-8') as csvfile:
                        reader = csv.reader(csvfile)
                        rows = list(reader)
                        cell_col_letters = ''.join(filter(str.isalpha, cell_to_extract))
                        cell_row = int(''.join(filter(str.isdigit, cell_to_extract))) - 1
                        col_index = openpyxl.utils.column_index_from_string(cell_col_letters) - 1
                        extracted_value = rows[cell_row][col_index]

                # Convert the extracted value to a float
                try:
                    numeric_value = float(extracted_value)
                except (ValueError, TypeError) as e:
                    messagebox.showerror("Error", f"Value '{extracted_value}' in file {file_name} cannot be converted to a number: {e}")
                    output_wb.close()
                    return

                # Write the numeric value to the output Excel file in the chosen column, on the same row.
                cell_address = f"{output_col_letters}{current_row}"
                output_sheet.range(cell_address).value = numeric_value

            except FileNotFoundError:
                messagebox.showerror("Error", f"Input file not found: {full_file_path}")
                output_wb.close()
                return
            except (KeyError, IndexError):
                messagebox.showerror("Error", f"Cell {cell_to_extract} not found in file {file_name}")
                output_wb.close()
                return
            except Exception as e:
                messagebox.showerror("Error", f"Unexpected error processing file {file_name}: {e}")
                output_wb.close()
                return

        # Auto-adjust column widths for the entire used range
        output_sheet.api.Columns.AutoFit()

        # Center-align the updated cells in the output column
        updated_range = output_sheet.range(f"{output_col_letters}{start_row}:{output_col_letters}{end_row}")
        updated_range.api.HorizontalAlignment = -4108  # -4108 corresponds to center alignment in Excel COM

        # Save the workbook without closing it
        # (If the workbook is already open, we get that instance; otherwise, we open it)
        for book in xw.books:
            if book.fullname.lower() == output_file_path.lower():
                output_wb = book
                break
        else:
            output_wb = xw.Book(output_file_path)

        output_wb.save()
        messagebox.showinfo("Success", "Data extracted and updated successfully.")

    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")

# ----------------- Light/Dark Mode Toggle ----------------- #
dark_mode = False

def toggle_theme():
    """Switch between light and dark themes using ttk style."""
    global dark_mode
    dark_mode = not dark_mode
    if dark_mode:
        style.theme_use("alt")
        style.configure('.', background='#2d2d2d', foreground='#ffffff')
        style.configure('TLabel', background='#2d2d2d', foreground='#ffffff')
        style.configure('TButton', background='#4d4d4d', foreground='#ffffff')
        style.configure('TEntry', fieldbackground='#4d4d4d', foreground='#ffffff')
        style.configure('TCombobox', fieldbackground='#4d4d4d', background='#2d2d2d', foreground='#ffffff')
        root.configure(bg='#2d2d2d')
        toggle_btn.config(text="Switch to Light Mode")
    else:
        style.theme_use("clam")
        style.configure('.', background='#f0f0f0', foreground='#000000')
        style.configure('TLabel', background='#f0f0f0', foreground='#000000')
        style.configure('TButton', background='#e0e0e0', foreground='#000000')
        style.configure('TEntry', fieldbackground='#ffffff', foreground='#000000')
        style.configure('TCombobox', fieldbackground='#ffffff', background='#f0f0f0', foreground='#000000')
        root.configure(bg='#f0f0f0')
        toggle_btn.config(text="Switch to Dark Mode")

# --------------------- Main GUI Setup --------------------- #
root = tk.Tk()
root.title("Excel/CSV Data Extractor")

# Use ttk.Style for a modern look
style = ttk.Style(root)
style.theme_use("clam")
root.configure(bg='#f0f0f0')
style.configure('.', background='#f0f0f0', foreground='#000000')

# Create and position widgets using ttk
pad_opts = {'padx': 5, 'pady': 5}

ttk.Label(root, text="Subfolder:").grid(row=0, column=0, sticky="w", **pad_opts)
folder_path_entry = ttk.Entry(root, width=50)
folder_path_entry.grid(row=0, column=1, **pad_opts)
ttk.Button(root, text="Browse", command=select_folder).grid(row=0, column=2, **pad_opts)

ttk.Label(root, text="Output Excel:").grid(row=1, column=0, sticky="w", **pad_opts)
output_file_entry = ttk.Entry(root, width=50)
output_file_entry.grid(row=1, column=1, **pad_opts)
ttk.Button(root, text="Browse", command=select_output_excel).grid(row=1, column=2, **pad_opts)

ttk.Label(root, text="Variable to Extract:").grid(row=2, column=0, sticky="w", **pad_opts)
variable_var = tk.StringVar(root)
variable_var.set(VARIABLES[0])
variable_dropdown = ttk.Combobox(root, textvariable=variable_var, values=VARIABLES, state="readonly", width=60)
variable_dropdown.current(0)
variable_dropdown.grid(row=2, column=1, **pad_opts, sticky="w")

ttk.Label(root, text="Output Column (e.g., F):").grid(row=3, column=0, sticky="w", **pad_opts)
output_col_entry = ttk.Entry(root, width=5)
output_col_entry.grid(row=3, column=1, sticky="w", **pad_opts)

ttk.Label(root, text="Output Row Range (e.g., 2:12):").grid(row=4, column=0, sticky="w", **pad_opts)
output_row_range_entry = ttk.Entry(root, width=10)
output_row_range_entry.grid(row=4, column=1, sticky="w", **pad_opts)

ttk.Button(root, text="Extract and Write", command=extract_and_write).grid(row=5, column=1, sticky="e", **pad_opts)
toggle_btn = ttk.Button(root, text="Switch to Dark Mode", command=toggle_theme)
toggle_btn.grid(row=5, column=2, sticky="e", **pad_opts)

root.mainloop()
