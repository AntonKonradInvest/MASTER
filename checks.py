from tkinter import simpledialog

import pandas as pd

from utils import showInfo


def check_missing_factuurnummers(df):
    """Checks if the 'Factuur' column in the DataFrame has missing numbers in the sequence."""
    factuur_column = df["factuur (H)"].dropna().astype(int)  # Convert to int and drop any NaN values
    sorted_factuurnummers = sorted(factuur_column.unique())  # Sort unique factuur numbers

    # Create a complete range from the minimum to the maximum factuurnummer
    expected_range = list(range(min(sorted_factuurnummers), max(sorted_factuurnummers) + 1))

    # Find the missing factuurnummers
    missing_factuurnummers = set(expected_range) - set(sorted_factuurnummers)

    if missing_factuurnummers:
        return f"Missing invoice numbers: {sorted(missing_factuurnummers)}"
    else:
        return "No missing invoice numbers in the Factuur column."


def check_non_numeric_boekhpl_reknr(df):
    """Checks if the 'boekhpl_reknr (D)' column contains the string 'CONTROLEREN' and logs related factuurnummers."""
    boekhpl_column = df["boekhpl_reknr (D)"].dropna()  # Drop any missing values

    # Find rows where the value is 'CONTROLEREN'
    control_values = boekhpl_column[boekhpl_column.astype(str) == "CONTROLEREN"]

    if not control_values.empty:
        # Find the corresponding factuurnummers where 'CONTROLEREN' is found
        factuurnummers = df.loc[control_values.index, "factuur (H)"]
        messages = [f"Please check grootboekrekening for factuur: {factuur}" for factuur in factuurnummers]
        return "\n".join(messages)
    else:
        return "All values OK for 'boekhpl_reknr (D)'."


def check_missing_relatiecode(merged_df, input_df, reference_file_path):
    # Find rows with missing Relatiecode
    missing_codes = merged_df[merged_df['Relatiecode'].isnull()]

    if not missing_codes.empty:
        # Get a list of unique missing names
        missing_codes_list = missing_codes[input_df.columns[0]].unique().tolist()
        total_missing = len(missing_codes_list)

        for i, name in enumerate(missing_codes_list, start=1):
            # Show progress to the user
            progress_message = f"Entering Relatiecode {i} of {total_missing} for {name}"
            print(progress_message)

            # Ask for the code once for each unique name
            code = simpledialog.askstring("Missing Code", f"Enter Relatiecode for {name} ({i}/{total_missing}):")
            if code:
                # Update the 'Relatiecode' for all rows with the same name
                merged_df.loc[merged_df[input_df.columns[0]] == name, 'Relatiecode'] = code

                # Append the new code to the reference CSV
                new_row = pd.DataFrame({'Name': [name], 'Relatiecode': [code]})
                with open(reference_file_path, mode='a', newline='', encoding='utf-8') as f:
                    new_row.to_csv(f, header=False, index=False, sep=';')

        # Show final message when done
        showInfo(f"All missing Relatiecodes have been entered.")
