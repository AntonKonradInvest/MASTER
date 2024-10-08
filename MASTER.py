import csv
import os
import unidecode
import pandas as pd
import tkinter as tk

from datetime import datetime
from tkinter import ttk, filedialog, simpledialog, messagebox

from Document import Document

CONFIG_FILE = os.path.join(os.path.expanduser("~"), "excel_converter_settings.ini")
RELATIECODES_FILE = "relatiecodes.csv"
GROOTBOEKREKENINGEN_FILE = "grootboekrekeningen.csv"

RELATIECODES_DF = pd.read_csv(RELATIECODES_FILE, delimiter=';', encoding='utf-8', dtype={'Relatiecode': str})

HIGHEST = None
LOWEST = None

EXPECTED_COLUMNS_BILLIT = [
    "Order nummer", "Datum", "Vervaldag", "Totaal inclusief", "Totaal exclusief", "Relatiecode"
]

EXPECTED_COLUMNS_ERELONEN = [
    "Gebouw", "Factuurnr", "Totaal brutto", "Totaal netto", "Totaal BTW", "Vervaldag", "Betaald", "Documentnummer",
    "Documentdatum", "Relatiecode"
]

EXPECTED_COLUMNS_RAPPELS = [
    "Gebouw", "Relatiecode", "Factuurnr", "Totaal brutto", "Totaal netto", "Totaal BTW", "Doc nr", "Documentdatum"
]

FACTUURNUMMERS = []

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


def view_reference_df():
    """Opens a new window to display and edit the reference DataFrame."""
    # Load the reference DataFrame
    reference_df = load_reference_df()

    # Create a new window
    view_window = tk.Toplevel()
    view_window.title("View and Edit Reference DataFrame")

    # Create a treeview to display the DataFrame
    tree = ttk.Treeview(view_window, columns=("Name", "Relatiecode"), show='headings')
    tree.heading("Name", text="Name")
    tree.heading("Relatiecode", text="Relatiecode")

    # Insert rows from reference_df into the treeview
    for index, row in reference_df.iterrows():
        tree.insert("", tk.END, values=(row['Name'], row['Relatiecode']))

    tree.pack(fill=tk.BOTH, expand=True)

    # Frame for search and editing options
    control_frame = tk.Frame(view_window)
    control_frame.pack(pady=10)

    # Search fields
    tk.Label(control_frame, text="Search Name:").grid(row=0, column=0)
    search_name_entry = tk.Entry(control_frame)
    search_name_entry.grid(row=0, column=1, padx=5)

    tk.Label(control_frame, text="Search Relatiecode:").grid(row=1, column=0)
    search_code_entry = tk.Entry(control_frame)
    search_code_entry.grid(row=1, column=1, padx=5)

    # Search button functionality
    def search_reference():
        search_name = search_name_entry.get().strip().lower()
        search_code = search_code_entry.get().strip()

        # Clear the treeview first
        tree.delete(*tree.get_children())

        # Filter the DataFrame based on search inputs
        filtered_df = reference_df
        if search_name:
            filtered_df = filtered_df[filtered_df['Name'].str.lower().str.contains(search_name)]
        if search_code:
            filtered_df = filtered_df[filtered_df['Relatiecode'].astype(str).str.contains(search_code)]

        # Repopulate the treeview with filtered results
        for index2, row2 in filtered_df.iterrows():
            tree.insert("", tk.END, values=(row2['Name'], row2['Relatiecode']))

    tk.Button(control_frame, text="Search", command=search_reference).grid(row=0, column=2, rowspan=2, padx=5)

    # Function to edit a selected row
    def edit_selected_row():
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a row to edit.")
            return

        # Get the selected row values
        selected_values = tree.item(selected_item, "values")
        current_name = selected_values[0]
        current_code = selected_values[1]

        # Prompt for new values
        new_name = simpledialog.askstring("Edit Name", f"Edit Name (current: {current_name}):",
                                          initialvalue=current_name)
        new_code = simpledialog.askstring("Edit Relatiecode", f"Edit Relatiecode (current: {current_code}):",
                                          initialvalue=current_code)

        if new_name and new_code:
            # Update the reference DataFrame
            reference_df.loc[
                (reference_df['Name'] == current_name) & (reference_df['Relatiecode'] == current_code), ['Name',
                                                                                                         'Relatiecode']] = [
                new_name, new_code]

            # Update the treeview
            tree.item(selected_item, values=(new_name, new_code))

    # Button to trigger editing
    tk.Button(control_frame, text="Edit Selected Row", command=edit_selected_row).grid(row=2, column=0, columnspan=3,
                                                                                       pady=5)

    # Function to save changes
    def save_changes():
        # Save the DataFrame back to CSV
        save_reference_df(reference_df)
        log_message("127: Info: Changes saved successfully!")

    # Save button
    tk.Button(control_frame, text="Save Changes", command=save_changes).grid(row=3, column=0, columnspan=3, pady=5)

    # Close button
    tk.Button(control_frame, text="Close", command=view_window.destroy).grid(row=4, column=0, columnspan=3, pady=5)


def load_reference_df():
    """Loads the reference CSV into a DataFrame."""
    try:
        reference_df = pd.read_csv(RELATIECODES_FILE, delimiter=';', encoding='utf-8')
        return reference_df
    except Exception as e:
        log_message(f"142: Error loading the reference file: {e}")
        return pd.DataFrame(columns=['Name', 'Relatiecode'])


def save_reference_df(reference_df):
    """Saves the reference DataFrame back to the CSV file."""
    reference_df.to_csv(RELATIECODES_FILE, sep=';', index=False, encoding='utf-8')


def get_column_letter(col_idx):
    """Convert a 0-based column index to an Excel-style column letter."""
    col_letter = ''
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        col_letter = chr(65 + remainder) + col_letter
    return col_letter


def check_missing_values_in_columns(df, conversion_folder, filter_column=None):
    # Columns to check for missing values
    columns = ["boekjaar_boekjaar (H)", "dagboek_dagboek (H)", "factuur (H)", "periode_periode (H)",
               "btwregimes_btwregime (H)", "factdat (H)", "relaties_code (H)", "vervdat (H)",
               "valuta_code (H)", "tebet (H)", "statusfact_status (H)", "codefcbd_codefcbd (H)",
               "basis (H)", "btwtebet (H)", "boekhpl_reknr (D)", "datum (D)", "bedrag (D)", "btwcodes_btwcode (D)"]

    missing_values_locations = {}

    # Filter based on populated rows if a filter_column is provided
    if filter_column:
        populated_df = df[df[filter_column].notna() & (df[filter_column] != '')]
    else:
        populated_df = df  # Check all rows if no filter column is specified

    # Iterate over columns and find missing values
    for column in columns:
        if column in df.columns:
            # Check for missing values (NaN or empty strings)
            missing = populated_df[column].isna() | (populated_df[column] == '')
            if missing.any():
                col_idx = df.columns.get_loc(column)  # Get the column index
                col_letter = get_column_letter(col_idx + 1)  # Convert index to Excel column letter (1-based index)

                # Get the row indices where the column has missing values
                missing_rows = missing[missing].index.tolist()

                # Store the missing locations as (column_letter, row_number)
                missing_locations = [(col_letter, row + 1) for row in missing_rows]
                missing_values_locations[column] = missing_locations
            else:
                missing_values_locations[column] = []  # No missing values in this column
        else:
            # If the column does not exist in the DataFrame
            missing_values_locations[column] = [("Unknown", row + 1) for row in range(len(df))]

    # Save missing values locations to a text file
    missing_order_nummers_file = os.path.join(conversion_folder, "missing_values.txt")
    with open(missing_order_nummers_file, "w") as f:
        for column, locations in missing_values_locations.items():
            for location in locations:
                f.write(f"Missing value in cell '{column}' at location: {location}\n")

    return missing_values_locations


def check_missing_relatiecodes(merged_df, name_column):
    """
    General function to check missing 'Relatiecodes' for any DataFrame.
    Handles both Billit and Erelonen cases.
    """
    # Normalize names
    merged_df = merged_df.copy()
    merged_df['cleaned_name'] = merged_df[name_column].apply(lambda name: unidecode.unidecode(str(name)).strip())

    # Find rows with missing Relatiecode
    missing_codes = merged_df[merged_df['Relatiecode'].isnull()]

    if not missing_codes.empty:
        # Get unique missing names
        missing_codes_list = missing_codes['cleaned_name'].unique().tolist()

        for i, cleaned_name in enumerate(missing_codes_list, start=1):
            code = None
            while not code:
                code = simpledialog.askstring("Missing Code",
                                              f"Enter Relatiecode for {cleaned_name} ({i}/{len(missing_codes_list)}):")
                index = merged_df[merged_df['cleaned_name'] == cleaned_name].index[0]
                merged_df.at[index, 'Relatiecode'] = code

                # Append to reference file
                new_row = pd.DataFrame({'Name': [cleaned_name], 'Relatiecode': [code]})
                with open(RELATIECODES_FILE, mode='a', newline='', encoding='utf-8') as f:
                    new_row.to_csv(f, header=False, index=False, sep=';')

        log_message(f"233: All missing Relatiecodes have been entered.")


def check_column_names(df, expected_columns):
    """Checks if the DataFrame has the expected columns. Returns only missing columns."""
    actual_columns = set(df.columns.str.strip())
    expected_columns = set(expected_columns)
    missing_columns = expected_columns - actual_columns
    return missing_columns


def clean_text(text):
    """
    Normalizes text by:
    - Removing accents or special characters using unidecode
    - Stripping leading/trailing whitespace
    - Replacing special dashes with standard dashes
    - Converting to lowercase

    Args:
        text (str): Input string to be cleaned and normalized.

    Returns:
        str: Cleaned and normalized string.
    """
    if isinstance(text, str):
        # Remove accents and special characters
        text = unidecode.unidecode(text)

        # Normalize by stripping whitespace and replacing non-standard dashes
        return text.strip().replace("–", "-")

    return text


def ensure_save_folder_exists(save_folder):
    if not os.path.exists(save_folder):
        os.makedirs(save_folder)


def validate_reference_file(reference_df):
    """
    Validates the reference file:
    - Ensures the reference file contains exactly 'Name' and 'Relatiecode' columns.
    - Ensures that one unique name has only one 'Relatiecode'.
    - Allows one 'Relatiecode' to be assigned to multiple names.
    """
    # Check if the reference file has the correct columns
    expected_columns = {'Name', 'Relatiecode'}
    actual_columns = set(reference_df.columns.str.strip())  # Strip any leading/trailing spaces from column names
    missing_columns = expected_columns - actual_columns

    if missing_columns:
        raise ValueError(f"Reference file must contain the following columns: {expected_columns}. "
                         f"Missing columns: {missing_columns}")

    # Check for duplicate names with different 'Relatiecode'
    name_duplicates = reference_df.groupby('Name')['Relatiecode'].nunique()
    name_inconsistencies = name_duplicates[name_duplicates > 1]

    if not name_inconsistencies.empty:
        inconsistent_names = name_inconsistencies.index.tolist()
        raise ValueError(f"Duplicate 'Relatiecode' entries found for names: {inconsistent_names}")

    # Since multiple names sharing the same 'Relatiecode' is allowed, no need to check that.

    # If no inconsistencies found
    log_message("300: Reference file is valid. No duplicate 'Relatiecode' entries found for the same name.")


def clean_and_normalize_dataframe(input_df, name_column):
    input_df['cleaned_name'] = input_df[input_df.columns[name_column]].apply(clean_text)
    return input_df


def merge_dataframes(input_df, name_column):
    # Clean and normalize the data
    input_df_cleaned = clean_and_normalize_dataframe(input_df, name_column)

    # Merge using the 'cleaned_name' for matching, and the 'Name' column from RELATIECODES_DF
    merged_df = pd.merge(input_df_cleaned, RELATIECODES_DF, how='left', left_on='cleaned_name', right_on='Name')

    # Drop 'Name' from RELATIECODES_DF if not needed in the final result
    merged_df.drop('Name', axis=1, inplace=True)

    return merged_df


def create_docs_from_excel_Billit(prepared_df, conversion_folder):
    """Reads Excel file and creates a list of Document objects for conversion Billit."""
    try:
        missing_columns = check_column_names(prepared_df, EXPECTED_COLUMNS_BILLIT)
        if missing_columns:
            log_message(f"326: Missing prepared_df columns: {missing_columns}")
            return []
    except Exception as e:
        log_message(f"329: Error reading the Excel file: {e}")
        return []

    docs = []
    missing_order_nummers = []
    for index, row in prepared_df.iterrows():
        try:
            # Handle Order nummer and check for missing
            order_nummer = row.get("Order nummer", "")
            if pd.isna(order_nummer) or not order_nummer.strip():
                # If Order nummer is missing, log the factuurnr or row number
                factuurnr = row.get("Factuurnr", f"Row {index + 2}")
                missing_order_nummers.append(factuurnr)
                continue  # Skip this row

            # Convert to string for further processing
            order_nummer = str(order_nummer)

            # Split only if valid
            if "-" in order_nummer:
                nummer_parts = order_nummer.split("-")
            else:
                nummer_parts = [None, None, None]

            dagboek, boekjaar, nummer = nummer_parts
            factuurcode = "F"
            if dagboek == "CN1":
                dagboek = "CN"
                factuurcode = "C"

            # Handle dates
            datum = pd.to_datetime(row["Datum"], errors='coerce')
            vervaldatum = pd.to_datetime(row["Vervaldag"], errors='coerce')

            # Ensure basisbedrag and tebetalen are numbers
            tebetalen = abs(row.get("Totaal inclusief", 0))  # Default to 0 if NaN
            basisbedrag = abs(row.get("Totaal exclusief", 0))  # Default to 0 if NaN

            # Check for missing dates
            if pd.isnull(datum) or pd.isnull(vervaldatum):
                raise ValueError("Invalid dates in the row.")

            # Create Document object
            doc = Document(
                boekjaar=boekjaar,
                dagboek=dagboek,
                nummer=nummer,
                datum=datum,
                relatiecode=row["Relatiecode"],
                vervaldatum=vervaldatum,
                tebetalen=tebetalen,
                basisbedrag=basisbedrag,
                omschrijving=row['Betreft']
            )
            doc.codefcbd_codefcbd = factuurcode
            docs.append(doc)

        except Exception as e:
            log_message(f"338: Error processing row: {e}")
            continue

    # Write missing Order nummers to a file
    if missing_order_nummers:
        missing_order_nummers_file = os.path.join(conversion_folder, "missing_order_nummers.txt")
        with open(missing_order_nummers_file, "w") as f:
            for factuurnr in missing_order_nummers:
                f.write(f"Missing Order nummer for factuurnr: {factuurnr}\n")

    return docs


def create_docs_from_excel_Erelonen(filepath):
    """Reads Excel file and creates a list of Document objects for conversion Erelonen, with options to skip rows and filter by year and month."""
    global HIGHEST, LOWEST
    try:
        df = pd.read_excel(filepath)
        # Filtering MAAND EN JAAR TOE TE VOEGEN
        missing_columns = check_column_names(df, EXPECTED_COLUMNS_ERELONEN)
        if missing_columns:
            log_message(f"408: Missing columns :{missing_columns}")
            return []

    except Exception as e:
        log_message(f"412: Error reading the Excel file: {e}")
        return []

    docs = []

    for _, row in df.iterrows():
        try:
            datum = row["Documentdatum"]
            vervaldatum = pd.to_datetime(row["Vervaldag"], errors='coerce')
            tebetalen = abs(row["Totaal brutto"])
            basisbedrag = abs(row["Totaal netto"])
            factuurnr = str(row["Factuurnr"])

            # Filter factuurnummers that don't start with "AF"
            if not factuurnr.startswith("AF"):
                continue

            if pd.isnull(datum) or pd.isnull(vervaldatum):
                raise ValueError("Invalid dates in the row.")

            # Split only if valid
            if "/" in factuurnr:
                nummer_parts = factuurnr.split("/")
            else:
                nummer_parts = [None, None]

            dagboek, factuurnummer = nummer_parts
            factuurnummer = factuurnummer[2::]
            FACTUURNUMMERS.append(factuurnummer)

            if dagboek == "AF1":
                dagboek = "VK2"
            elif dagboek == "AF2":
                dagboek = "VK3"

            # Create the Document object
            doc = Document(
                boekjaar=datum.year,
                dagboek=dagboek,
                nummer=factuurnummer,
                datum=datum,
                relatiecode=row["Relatiecode"],
                vervaldatum=vervaldatum,
                tebetalen=tebetalen,
                basisbedrag=basisbedrag,
                omschrijving="",
                rekening="NA"
            )

            if HIGHEST is None or factuurnummer > HIGHEST:
                HIGHEST = factuurnummer

            if LOWEST is None or factuurnummer < LOWEST:
                LOWEST = factuurnummer

            if doc.relatiecode.startswith("H"):
                doc.boekhpl_reknr = "700010"
            elif doc.relatiecode.startswith("D"):
                doc.boekhpl_reknr = "700020"
            elif doc.relatiecode.startswith("L"):
                doc.boekhpl_reknr = "700030"
            elif doc.relatiecode.startswith("G"):
                doc.boekhpl_reknr = "700040"
            elif doc.relatiecode.startswith("R"):
                doc.boekhpl_reknr = "700050"

            docs.append(doc)

        except Exception as e:
            log_message(f"448: Error processing row: {e}")
            continue

    return docs


def create_docs_from_excel_Rappels(file_path):
    """Reads Excel file and creates a list of Document objects for conversion Rappels."""
    possible_factuurnummer_columns = ["Factuurnummer", "Factuurnr"]

    try:
        df = pd.read_excel(file_path)
        missing_columns = check_column_names(df, EXPECTED_COLUMNS_RAPPELS)

        if missing_columns:
            log_message(f"463: Missing column(s): {missing_columns}")
            return []

        factuurnummer_col = next((col for col in possible_factuurnummer_columns if col in df.columns), None)
        if not factuurnummer_col:
            log_message("468: Missing 'Factuurnummer' or 'Factuurnr' column.")
            return []

    except Exception as e:
        log_message(f"472: Error reading the Excel file: {e}")
        return []

    docs = []
    for _, row in df.iterrows():
        try:
            datum = pd.to_datetime(row["Documentdatum"], errors='coerce')
            tebetalen = abs(row["Totaal brutto"])
            basisbedrag = abs(row["Totaal netto"])

            factuurnr = row[factuurnummer_col]
            log_message(factuurnr)
            dagboek = "VK4"

            factuurnr = factuurnr[4:]

            if pd.isnull(datum):
                raise ValueError("Invalid date in the row.")

            doc = Document(
                boekjaar=datum.year,
                dagboek=dagboek,
                nummer=factuurnr,
                datum=datum,
                relatiecode=row["Relatiecode"],
                tebetalen=tebetalen,
                basisbedrag=basisbedrag,
                vervaldatum=datum + pd.Timedelta(days=15),
                omschrijving=""
            )
            docs.append(doc)
        except Exception as e:
            log_message(f"504: Error processing row: {e}")
            continue

    docs.sort(key=lambda x: x.nummer)

    return docs


def initialize_grootboekrekeningen_file():
    """Initializes the grootboekrekeningen.csv file if it doesn't exist."""
    if not os.path.exists(GROOTBOEKREKENINGEN_FILE):
        with open(GROOTBOEKREKENINGEN_FILE, mode='w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['Name', 'Grootboekrekening'])  # Add header to the CSV file
        log_message(f"{GROOTBOEKREKENINGEN_FILE} file created with header.")


def load_and_merge_file(input_path, type):
    """
    Helper function to load, clean, and merge the input file with RELATIECODES_DF.
    This function is shared between Billit and Erelonen file preparation.
    """
    try:
        # Load the input DataFrame
        if type == "erelonen":
            input_df = pd.read_excel(input_path, header=1)
            input_df = clean_and_normalize_dataframe(input_df, 0)
            merged_df = merge_dataframes(input_df, 0)

        elif type == "billit":
            input_df = pd.read_excel(input_path)
            input_df = clean_and_normalize_dataframe(input_df, 3)
            merged_df = merge_dataframes(input_df, 3)

        # Validate the reference file
        validate_reference_file(RELATIECODES_DF)

        # Merge with the reference file (RELATIECODES_DF)

        return merged_df

    except Exception as e:
        log_message(f"537: Error loading or merging data: {e}")
        return None


def prepare_erelonen_excel_file(input_path, month, year, save_folder):
    try:
        ensure_save_folder_exists(save_folder)

        # Check if input path exists
        if not input_path or not os.path.exists(input_path):
            log_message("547: Please select a valid input file.")
            return None

        # Load the Excel file (skip the first two rows to get headers from the third row)
        df = pd.read_excel(input_path, skiprows=2)
        if df is None or df.empty:
            log_message("553: The loaded DataFrame is empty.")
            return None

        # Rename columns for clarity (adjust according to the actual input data structure)
        df.rename(columns={
            'Unnamed: 0': 'Gebouw',
            'Unnamed: 1': 'Factuurnr',
            'Unnamed: 2': 'Totaal brutto',
            'Unnamed: 3': 'Totaal netto',
            'Unnamed: 4': 'Totaal BTW',
            'Unnamed: 5': 'Vervaldag',
            'Unnamed: 6': 'Betaald',
            'Unnamed: 7': 'Documentnummer',
            'Unnamed: 8': 'Documentdatum'
        }, inplace=True)

        log_message(f"569: Columns in merged_df: {list(df.columns)}")

        # Merge with the reference DataFrame (RELATIECODES_DF) to get the 'Relatiecode'
        merged_df = pd.merge(df, RELATIECODES_DF, how='left', left_on='Gebouw', right_on='Name')

        # Check if the "Relatiecode" column exists after the merge
        if 'Relatiecode' not in merged_df.columns:
            log_message("576: Error: 'Relatiecode' column is missing after merging with the reference data.")
            return None

        # Filter rows where 'Factuurnr' starts with 'AF'
        df_filtered = merged_df[merged_df['Factuurnr'].str.startswith('AF', na=False)]

        check_missing_relatiecodes(df_filtered, "Gebouw")

        # Convert 'Documentdatum' to datetime and log any invalid entries
        df_filtered = df_filtered.copy()
        df_filtered['Documentdatum'] = pd.to_datetime(df_filtered['Documentdatum'], errors='coerce')

        # Log and filter out rows with invalid 'Documentdatum' values (NaT after conversion)
        invalid_dates = df_filtered[df_filtered['Documentdatum'].isna()]
        if not invalid_dates.empty:
            log_message(
                f"590: Warning: Found invalid dates in the following rows:\n{invalid_dates[['Factuurnr', 'Documentdatum']]}")

        # Keep only rows with valid dates
        df_filtered = df_filtered[df_filtered['Documentdatum'].notna()]

        # Filter by the specified year and month
        df_filtered = df_filtered[
            (df_filtered['Documentdatum'].dt.year == year) &
            (df_filtered['Documentdatum'].dt.month == month)
            ]

        # Reordering columns: Move 'Relatiecode' to the second position if it exists
        cols = df_filtered.columns.tolist()
        if 'Relatiecode' in cols:
            cols.insert(1, cols.pop(cols.index('Relatiecode')))
            df_filtered = df_filtered[cols]

        # Save the filtered DataFrame to an Excel file
        output_file_path = os.path.join(save_folder, "erelonen.xlsx")
        print(df_filtered.columns)

        df_filtered.to_excel(output_file_path, index=False)

        log_message(f"611: Data saved to {output_file_path}")
        return output_file_path

    except Exception as e:
        log_message(f"615: An error occurred: {e}")
        return None


def prepare_billit_excel_file(input_path, save_folder):
    try:
        ensure_save_folder_exists(save_folder)

        # Check if input path exists
        if not input_path or not os.path.exists(input_path):
            log_message("625: Please select a valid input file.")
            return None

        # Load and merge data using helper function
        merged_df = load_and_merge_file(input_path, "billit")
        if merged_df is None:
            return None

        log_message(f"633: Columns in merged_df: {list(merged_df.columns)}")

        # Filter rows based on 'Order nummer' for facturen (VK-) and creditnota's (CN1-)
        df_facturen = merged_df[merged_df['Order nummer'].str.startswith('VK-', na=False)]
        df_creditnota = merged_df[merged_df['Order nummer'].str.startswith('CN1-', na=False)]

        # Combine both facturen and creditnota's into a single DataFrame
        df_filtered = pd.concat([df_facturen, df_creditnota], ignore_index=True)

        # Handle missing Relatiecodes for the combined DataFrame
        check_missing_relatiecodes(df_filtered, "Bedrijf")

        # Convert 'Datum' to datetime format
        df_filtered = df_filtered.copy()  # This ensures a deep copy
        df_filtered.loc[:, 'Datum'] = pd.to_datetime(df_filtered['Datum'], errors='coerce')

        # Save to an Excel file
        output_file_path = os.path.join(save_folder, "prepared_billit_file.xlsx")
        df_filtered.to_excel(output_file_path, index=False)

        return df_filtered

    except Exception as e:
        log_message(f"656: An error occurred: {e}")
        return None


def select_excel_file():
    """Opens a file dialog to select an Excel file."""
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    return file_path


def select_output_folder():
    """Opens a file dialog to select an output folder."""
    folder_path = filedialog.askdirectory()
    return folder_path


def save_excel_file(df, folder_path, add=""):
    nowtime = datetime.now().strftime("%H_%M")
    filename = f"{nowtime + add}.xlsx"
    full_path = os.path.join(folder_path, filename)
    try:
        df.to_excel(full_path, index=False)
        log_message(f"678: Excel file saved at {full_path}")
    except Exception as e:
        log_message(f"680: Error saving the Excel file: {e}")


def convert():
    initialize_grootboekrekeningen_file()
    output_folder = default_output_folder.get()
    if not output_folder:
        log_message("686: No output folder set.")
        return

    current_time_folder = datetime.now().strftime("%H_%M")
    output_folder_path = os.path.join(output_folder, current_time_folder)
    ensure_save_folder_exists(output_folder_path)

    conversion_type = conversion_type_var.get()

    selected_month_name = month_var.get()
    selected_month = month_mapping[selected_month_name]
    selected_year = int(year_var.get())

    if conversion_type == "Billit":
        # Prepare the Billit file, which loads and merges data
        prepared_df = prepare_billit_excel_file(default_input_file.get(), output_folder_path)
        # Check if the DataFrame `prepared_df` is valid
        if prepared_df is not None and not prepared_df.empty:
            # Call the function that creates the list of Document objects
            docs = create_docs_from_excel_Billit(prepared_df, output_folder_path)
            # Check if `docs` contains any documents
            if docs:
                # Convert each Document object to a dictionary and create a DataFrame
                docs_data = [factuur.to_dict() for factuur in docs]
                df = pd.DataFrame(docs_data)

                if 'factuur (H)' in df.columns:
                    # Remove leading zeros and convert to numeric
                    df['factuur (H)'] = pd.to_numeric(df['factuur (H)'].astype(str).str.lstrip('0'), errors='coerce')
                    # Sort the DataFrame by 'factuur (H)'
                    df = df.sort_values(by='factuur (H)', na_position='last')

                if not df.empty:
                    # Check for missing values in the specified columns
                    missing_values = check_missing_values_in_columns(df, output_folder_path)

                    # If no missing values were found in any of the relevant columns
                    if all(len(locations) == 0 for locations in missing_values.values()):
                        save_excel_file(df, output_folder_path, "billit")
                        log_message("727: Info: Excel file saved without missing values.")
                    else:
                        # Log the missing values
                        log_message("730: Warning: Some columns contain empty cells.")
                        for column, locations in missing_values.items():
                            if locations:  # If there are missing values in this column
                                log_message(f"733: Missing values in column '{column}': {locations}")

                        # Save the file with a different name to indicate missing cells
                        save_excel_file(df, output_folder_path, "withemptycells")
                        log_message("737: Info: Excel file saved with missing values.")

                else:
                    log_message("740: Warning: No data to save in Billit file.")
            else:
                log_message("742: Warning: No documents created for Billit conversion.")
        else:
            log_message("744: Warning: The prepared Billit DataFrame is empty.")


    elif conversion_type == "Erelonen":

        prepared_file_path = prepare_erelonen_excel_file(default_input_file.get(), selected_month, selected_year,
                                                         output_folder_path)

        if prepared_file_path:
            docs = create_docs_from_excel_Erelonen(prepared_file_path)

            docs_data = [factuur.to_dict() for factuur in docs]

            sorted_data = pd.DataFrame(docs_data).sort_values(by='factuur (H)', na_position='last')
            print(sorted_data["boekhpl_reknr (D)"])
            # Perform the check for missing rekeningnummers just before saving
            sorted_data = check_missing_rekeningnummers(sorted_data, "relaties_code (H)")  # Assuming "Relatiecode" column has the names
            print(sorted_data["boekhpl_reknr (D)"])

            # Save the updated sorted_data to Excel after the missing rekeningnummers check

            save_excel_file(sorted_data, output_folder_path, str(LOWEST))

    log_message(f"752: Files saved to {output_folder_path}")


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


def check_missing_rekeningnummers(df, name_column):
    """
    General function to check missing 'Grootboekrekeningen' for any DataFrame.
    Handles Erelonen cases.
    """
    # Load grootboekrekeningen.csv into a DataFrame, ensuring columns are treated as strings
    if os.path.exists(GROOTBOEKREKENINGEN_FILE):
        grootboek_df = pd.read_csv(GROOTBOEKREKENINGEN_FILE, delimiter=';', encoding='utf-8', dtype=str)
    else:
        grootboek_df = pd.DataFrame(columns=['Code', 'Grootboekrekening'])  # Ensure 'Code' and 'Grootboekrekening' are strings

    # Ensure 'Code' column is always treated as a string and normalized
    grootboek_df['Code'] = grootboek_df['Code'].str.strip().str.lower()  # Normalize 'Code' in CSV
    df = df.copy()  # Make a copy of the dataframe to avoid modifying in place
    df['cleaned_name'] = df[name_column].astype(str).str.strip().str.lower()  # Normalize 'cleaned_name'

    log_message(f"Checking missing rekeningen for: {df['cleaned_name'].unique()}")  # Debugging

    # Find rows where Grootboekrekening is "NA"
    missing_rekeningen = df[df['boekhpl_reknr (D)'] == "NA"]

    if not missing_rekeningen.empty:
        # Get unique missing names
        missing_names_list = missing_rekeningen['cleaned_name'].unique().tolist()

        for i, cleaned_name in enumerate(missing_names_list, start=1):
            rekening = None

            # Check if a rekening already exists in grootboek_df for this name
            existing_rekening = grootboek_df[grootboek_df['Code'] == cleaned_name]  # Use 'Code' for checking

            log_message(f"Checking for {cleaned_name} in CSV...")  # Debugging

            if not existing_rekening.empty:
                # Use the existing rekening if it is already in the grootboek_df
                rekening = existing_rekening.iloc[0]['Grootboekrekening']
                log_message(f"Using existing grootboekrekening for {cleaned_name}: {rekening}")
            else:
                # Prompt for the rekening if it doesn't exist in the CSV file
                while not rekening:
                    rekening = simpledialog.askstring(
                        "Missing rekening",
                        f"Enter erelonen grootboekrekening for {cleaned_name} ({i}/{len(missing_names_list)}):"
                    )

            # Update the rekening in the DataFrame
            df.loc[df['cleaned_name'] == cleaned_name, 'boekhpl_reknr (D)'] = rekening

            # Append to the grootboek_df if it's a new entry
            if existing_rekening.empty:
                new_row = pd.DataFrame({'Code': [cleaned_name], 'Grootboekrekening': [rekening]})
                grootboek_df = pd.concat([grootboek_df, new_row], ignore_index=True)

        # Save the updated grootboek_df to the CSV file
        grootboek_df.to_csv(GROOTBOEKREKENINGEN_FILE, sep=';', index=False)

    log_message("All missing rekeningnummers have been entered.")

    return df  # Return the updated DataFrame


def log_message(message):
    logbook.config(state='normal')
    logbook.insert(tk.END, f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")
    logbook.config(state='disabled')
    logbook.yview(tk.END)


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
input_file_path = default_input_file.get()
output_folder_path = default_output_folder.get()
conversion_type_var = tk.StringVar(value="Billit")
code_factuur = tk.StringVar(value="Factuur")

# Tab 1: Convert
frame1 = tk.Frame(tab1, padx=10, pady=10)
frame1.pack(padx=10, pady=10)

tk.Label(frame1, text="Input Excel File:").grid(row=0, column=0, sticky='w')
tk.Entry(frame1, textvariable=default_input_file, width=50).grid(row=0, column=1, padx=5)
tk.Button(frame1, text="Browse...", command=lambda: default_input_file.set(filedialog.askopenfilename())).grid(row=0,
                                                                                                               column=2)

tk.Label(frame1, text="Default Output Folder:").grid(row=1, column=0, sticky='w')
tk.Entry(frame1, textvariable=default_output_folder, width=50).grid(row=1, column=1, padx=5)
tk.Button(frame1, text="Browse...", command=lambda: default_output_folder.set(filedialog.askdirectory())).grid(row=1,
                                                                                                               column=2)
tk.Label(frame1, text="Conversion Type:").grid(row=2, column=0, sticky='w')
tk.OptionMenu(frame1, conversion_type_var, "Billit", "Erelonen", "Rappels").grid(row=2, column=1, padx=5)

# Month and Year selectors (Initially hidden)
month_label = tk.Label(frame1, text="Select Month:")
month_var = tk.StringVar(value=datetime.now().strftime("%B"))
month_dropdown = tk.OptionMenu(frame1, month_var, *[datetime(2000, i, 1).strftime('%B') for i in range(1, 13)])

year_label = tk.Label(frame1, text="Select Year:")
year_var = tk.StringVar(value=datetime.now().strftime("%Y"))
year_dropdown = tk.OptionMenu(frame1, year_var, *[str(i) for i in range(datetime.now().year - 4, datetime.now().year + 4)])

# Convert Button
tk.Button(frame1, text="Convert", command=convert, width=30).grid(row=5, columnspan=3, pady=10)
tk.Button(root, text="View and Edit Reference DataFrame", command=view_reference_df).pack(pady=20)

conversion_type_var.trace_add('write', update_month_year_visibility)
update_month_year_visibility()

# Create a logbook (Text widget) to display logs at the bottom of the GUI
logbook = tk.Text(root, height=10, width=80, state='disabled')
logbook.pack(pady=10)

# Start the GUI event loop
root.mainloop()