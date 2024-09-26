from Document_converters import create_docs_from_excel_Billit, create_docs_from_excel_Erelonen, create_docs_from_excel_Rappels
from checks import check_missing_relatiecode

import os
import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
from datetime import datetime

from utils import showWarning, showInfo, showError

CONFIG_FILE = os.path.join(os.path.expanduser("~"), "excel_converter_settings.ini")

# Add the month mapping dictionary
month_mapping = {
    "January": 1,
    "February": 2,
    "March": 3,
    "April": 4,
    "May": 5,
    "June": 6,
    "July": 7,
    "August": 8,
    "September": 9,
    "October": 10,
    "November": 11,
    "December": 12
}

# Refactored save function for documents
def save_converted_data(docs, output_folder, filename="converted_data.xlsx"):
    """Saves the converted Document objects into an Excel file."""
    docs_data = [factuur.to_dict() for factuur in docs]
    df = pd.DataFrame(docs_data)
    output_file_path = os.path.join(output_folder, filename)
    df.to_excel(output_file_path, index=False)
    return output_file_path


def clean_text(text):
    if isinstance(text, str):
        return text.strip().replace("–", "-")  # Replace any special dashes with a standard one
    return text


def ensure_save_folder_exists(save_folder):
    if not os.path.exists(save_folder):
        os.makedirs(save_folder)


def load_dataframes(input_file_path, reference_file_path):
    input_df = pd.read_excel(input_file_path)
    reference_df = pd.read_csv(reference_file_path, delimiter=';', encoding='utf-8')
    return input_df, reference_df


def validate_reference_columns(reference_df, reference_file_path):
    if list(reference_df.columns) != ['Name', 'Relatiecode']:
        raise ValueError(f"Reference file: {reference_file_path} must have exactly 2 columns: 'Name' and 'Relatiecode'")


def clean_and_normalize_dataframes(input_df, reference_df):
    input_df[input_df.columns[0]] = input_df[input_df.columns[0]].apply(clean_text)
    reference_df['Name'] = reference_df['Name'].apply(clean_text)
    return input_df, reference_df


def merge_dataframes(input_df, reference_df):
    merged_df = pd.merge(input_df, reference_df, how='left', left_on=input_df.columns[0], right_on='Name')
    merged_df.drop('Name', axis=1, inplace=True)
    return merged_df


def prepare_excel_file(input_file_path, save_folder, reference_file_path="reference_codes.csv"):
    try:

        ensure_save_folder_exists(save_folder)

        # Load input and reference DataFrames
        input_df, reference_df = load_dataframes(input_file_path, reference_file_path)

        # Validate reference columns
        validate_reference_columns(reference_df, reference_file_path)

        # Clean and normalize the data
        input_df, reference_df = clean_and_normalize_dataframes(input_df, reference_df)

        # Merge DataFrames
        merged_df = merge_dataframes(input_df, reference_df)

        # Handle missing 'Relatiecode'
        check_missing_relatiecode(merged_df, input_df, reference_file_path)

        # Rearrange columns and save the final DataFrame
        cols = merged_df.columns.tolist()
        cols.insert(1, cols.pop(cols.index('Relatiecode')))
        merged_df = merged_df[cols]

        current_time = datetime.now().strftime("%H%M")
        save_path = os.path.join(save_folder, f"{current_time}.xlsx")

        merged_df.to_excel(save_path, index=False)

        return save_path

    except Exception as e:
        showError(f"An error occurred: {e}")
        return None


def select_excel_file():
    """Opens a file dialog to select an Excel file."""
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    return file_path


def select_output_folder():
    """Opens a file dialog to select an output folder."""
    folder_path = filedialog.askdirectory()
    return folder_path


def save_excel_file(df, folder_path):
    """Saves a DataFrame to an Excel file in the specified folder."""
    nowtime = datetime.now().strftime("%H_%M")
    filename = f"{nowtime}.xlsx"
    full_path = os.path.join(folder_path, filename)
    try:
        df.to_excel(full_path, index=False)
        showInfo(f"Excel file saved at {full_path}")
    except Exception as e:
        showError(f"Error saving the Excel file: {e}")


def convert():
    """Converts the selected Excel file to documents and saves them."""
    input_file_path = default_input_file.get()

    if not input_file_path:
        showWarning("No input file selected.")
        return

    output_folder = default_output_folder.get()

    if not output_folder:
        showWarning("No output folder set.")
        return

    conversion_type = conversion_type_var.get()

    # Get selected month and year (correct month handling)
    selected_month_name = month_var.get()
    selected_month = month_mapping[selected_month_name]  # Map the month name to a number
    selected_year = int(year_var.get())  # Year is already a number

    # Billit conversion
    if conversion_type == "Billit":
        docs = create_docs_from_excel_Billit(input_file_path, output_folder)
        # Set boekhpl_reknr for Billit conversion to 700002 for all documents
        for doc in docs:
            doc.boekhpl_reknr = 700002

        # Save the processed data to an Excel file in the output folder
        docs_data = [factuur.to_dict() for factuur in docs]
        df = pd.DataFrame(docs_data)
        output_file_path = os.path.join(output_folder, "converted_data.xlsx")
        df.to_excel(output_file_path, index=False)

    # Erelonen conversion - first run prepare_excel_file, then use the prepared file for conversion
    elif conversion_type == "Erelonen":
        prepared_file_path = prepare_excel_file(input_file_path, output_folder)
        if prepared_file_path:
            # Pass the correct month and year for filtering
            docs = create_docs_from_excel_Erelonen(prepared_file_path, skip_rows=2, selected_year=selected_year, selected_month=selected_month)

            # Proceed with the rest of the Erelonen conversion process
            for doc in docs:
                relatiecode_str = str(doc.relatiecode)  # Ensure relatiecode is treated as a string
                if relatiecode_str.startswith("H"):
                    doc.boekhpl_reknr = 700010
                elif relatiecode_str.startswith("D"):
                    doc.boekhpl_reknr = 700020
                elif relatiecode_str.startswith("L"):
                    doc.boekhpl_reknr = 700030
                elif relatiecode_str.startswith("G"):
                    doc.boekhpl_reknr = 700040
                elif relatiecode_str.startswith("R"):
                    doc.boekhpl_reknr = 700050
                else:
                    doc.boekhpl_reknr = "CONTROLEREN"

            # Save the processed data to an Excel file in the output folder
            docs_data = [factuur.to_dict() for factuur in docs]
            df = pd.DataFrame(docs_data)
            output_file_path = os.path.join(output_folder, "converted_data.xlsx")
            df.to_excel(output_file_path, index=False)

    elif conversion_type == "Rappels":
        docs = create_docs_from_excel_Rappels(input_file_path)

    else:
        showWarning("Invalid conversion type selected.")
        return

    showInfo(f"Files saved to {output_folder}")



# Function to update month and year visibility based on conversion type
def update_month_year_visibility(*args):
    """Show month and year dropdowns only for Erelonen conversion."""
    if conversion_type_var.get() == "Erelonen":
        month_label.grid(row=3, column=0, sticky='w', pady=5)
        month_dropdown.grid(row=3, column=1, padx=5)
        year_label.grid(row=4, column=0, sticky='w', pady=5)
        year_dropdown.grid(row=4, column=1, padx=5)
    else:
        month_label.grid_forget()
        month_dropdown.grid_forget()
        year_label.grid_forget()
        year_dropdown.grid_forget()


def rearrange_columns(docs_data):
    """
    Rearranges columns or modifies data within the document list (docs_data)
    based on specific logic for each conversion type.
    """
    for doc_data in docs_data:
        # Example: Adjust columns order or modify data
        doc_data['omschr (D)'] = doc_data.pop('omschrijving (H)', None)  # Move 'omschrijving' to 'omschr (D)'
        doc_data['boekhpl_reknr (D)'] = doc_data.pop('boekhpl_reknr (D)', None)  # Example for reordering

    return docs_data  # Return the modified docs data


def set_default_input_file():
    """Sets the default input file path."""
    file_path = select_excel_file()
    if file_path:
        default_input_file.set(file_path)


def set_default_output_folder():
    """Sets the default output folder path."""
    folder_path = select_output_folder()
    if folder_path:
        default_output_folder.set(folder_path)


# Create the main window
root = tk.Tk()
root.title("Excel to Document Converter")

# Create a tab control widget
tab_control = ttk.Notebook(root)
tab1 = ttk.Frame(tab_control)
tab_control.add(tab1, text="Convert")
tab_control.pack(expand=1, fill="both")

# Default input file and output folder
default_input_file = tk.StringVar()
default_output_folder = tk.StringVar()
conversion_type_var = tk.StringVar(value="Billit")

# Tab 1: Convert
frame1 = tk.Frame(tab1, padx=10, pady=10)
frame1.pack(padx=10, pady=10)

tk.Label(frame1, text="Input Excel File:").grid(row=0, column=0, sticky='w')
tk.Entry(frame1, textvariable=default_input_file, width=50).grid(row=0, column=1, padx=5)
tk.Button(frame1, text="Browse...", command=lambda: default_input_file.set(filedialog.askopenfilename())).grid(row=0, column=2)

tk.Label(frame1, text="Default Output Folder:").grid(row=1, column=0, sticky='w')
tk.Entry(frame1, textvariable=default_output_folder, width=50).grid(row=1, column=1, padx=5)
tk.Button(frame1, text="Browse...", command=lambda: default_output_folder.set(filedialog.askdirectory())).grid(row=1, column=2)

tk.Label(frame1, text="Conversion Type:").grid(row=2, column=0, sticky='w')
tk.OptionMenu(frame1, conversion_type_var, "Billit", "Erelonen", "Rappels", command=update_month_year_visibility).grid(row=2, column=1, padx=5)

# Month and Year selectors (Initially hidden)
month_label = tk.Label(frame1, text="Select Month:")
month_var = tk.StringVar(value=datetime.now().strftime("%B"))
month_dropdown = tk.OptionMenu(frame1, month_var, *[datetime(2000, i, 1).strftime('%B') for i in range(1, 13)])

year_label = tk.Label(frame1, text="Select Year:")
year_var = tk.StringVar(value=datetime.now().strftime("%Y"))
year_dropdown = tk.OptionMenu(frame1, year_var, *[str(i) for i in range(2020, 2031)])

# Convert Button
tk.Button(frame1, text="Convert", command=convert, width=30).grid(row=5, columnspan=3, pady=10)

conversion_type_var.trace_add('write', update_month_year_visibility)
update_month_year_visibility()

# Start the GUI event loop
root.mainloop()
