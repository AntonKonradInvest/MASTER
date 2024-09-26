import os
import pandas as pd
from Document import Document
from utils import showError


def check_column_names(df, expected_columns):
    """Checks if the DataFrame has the expected columns. Returns only missing columns."""
    actual_columns = set(df.columns.str.strip())
    expected_columns = set(expected_columns)
    missing_columns = expected_columns - actual_columns
    return missing_columns


def create_docs_from_excel_Billit(file_path, conversion_folder):
    """Reads Excel file and creates a list of Document objects for conversion Billit."""
    expected_columns = [
        "Order nummer", "Datum", "Vervaldag", "Totaal inclusief", "Totaal exclusief", "Relatiecode"
    ]
    try:
        df = pd.read_excel(file_path)
        missing_columns = check_column_names(df, expected_columns)

        if missing_columns:
            showError(f"Missing columns: {missing_columns}")
            return []
    except Exception as e:
        showError(f"Error reading the Excel file: {e}")
        return []

    docs = []
    missing_order_nummers = []

    for index, row in df.iterrows():
        try:
            # Handle Order nummer and check for missing
            order_nummer = row.get("Order nummer", "")
            if pd.isna(order_nummer) or not order_nummer.strip():
                # If Order nummer is missing, log the factuurnr or row number
                factuurnr = row.get("Factuurnr", f"Row {index+2}")
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
            factuurcode = ""
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
            )
            doc.codefcbd_codefcbd = factuurcode
            docs.append(doc)

        except Exception as e:
            print(f"Error processing row: {e}")  # Debugging print
            showError(f"Error processing row: {e}")
            continue

    # Write missing Order nummers to a file
    if missing_order_nummers:
        missing_order_nummers_file = os.path.join(conversion_folder, "missing_order_nummers.txt")
        with open(missing_order_nummers_file, "w") as f:
            for factuurnr in missing_order_nummers:
                f.write(f"Missing Order nummer for factuurnr: {factuurnr}\n")

    return docs


def create_docs_from_excel_Erelonen(file_path, skip_rows=0, selected_year=None, selected_month=None):
    """Reads Excel file and creates a list of Document objects for conversion Erelonen, with options to skip rows and filter by year and month."""

    expected_columns = [
        "Documentdatum", "Vervaldag", "Totaal brutto", "Totaal netto", "Relatiecode"
    ]
    possible_factuurnummer_columns = ["Factuurnummer", "Factuurnr"]

    try:
        # Read the file and skip the first 'skip_rows' rows
        df = pd.read_excel(file_path, skiprows=skip_rows)
        missing_columns = check_column_names(df, expected_columns)

        if missing_columns:
            showError(f"Missing column(s): {missing_columns}")
            return []

        factuurnummer_col = next((col for col in possible_factuurnummer_columns if col in df.columns), None)
        if not factuurnummer_col:
            showError("Missing 'Factuurnummer' or 'Factuurnr' column.")
            return []

        # Convert dates and filter based on selected year and month
        df["Documentdatum"] = pd.to_datetime(df["Documentdatum"], errors='coerce')
        if selected_year and selected_month:
            df = df[(df["Documentdatum"].dt.year == selected_year) &
                    (df["Documentdatum"].dt.month == selected_month)]

    except Exception as e:
        showError(f"Error reading the Excel file: {e}")
        return []

    docs = []
    for _, row in df.iterrows():
        try:
            datum = row["Documentdatum"]
            vervaldatum = pd.to_datetime(row["Vervaldag"], errors='coerce')
            tebetalen = abs(row["Totaal brutto"])
            basisbedrag = abs(row["Totaal netto"])

            factuurnr = str(row[factuurnummer_col])

            # Filter factuurnummers that don't start with "AF"
            if not factuurnr.startswith("AF"):
                continue

            if pd.isnull(datum) or pd.isnull(vervaldatum):
                raise ValueError("Invalid dates in the row.")

            # Create the Document object
            doc = Document(
                boekjaar=datum.year,
                dagboek="Erelonen",  # You can modify this as needed
                nummer=factuurnr,
                datum=datum,
                relatiecode=row["Relatiecode"],
                vervaldatum=vervaldatum,
                tebetalen=tebetalen,
                basisbedrag=basisbedrag,
            )
            docs.append(doc)

        except Exception as e:
            showError(f"Error processing row: {e}")
            continue

    return docs


def create_docs_from_excel_Rappels(file_path):
    """Reads Excel file and creates a list of Document objects for conversion Rappels."""
    expected_columns = [
        "Gebouw", "Relatiecode", "Factuurnr", "Totaal brutto", "Totaal netto", "Totaal BTW", "Doc nr", "Documentdatum"
    ]
    possible_factuurnummer_columns = ["Factuurnummer", "Factuurnr"]

    try:
        df = pd.read_excel(file_path)
        missing_columns = check_column_names(df, expected_columns)

        if missing_columns:
            showError(f"Missing column(s): {missing_columns}")
            return []

        factuurnummer_col = next((col for col in possible_factuurnummer_columns if col in df.columns), None)
        if not factuurnummer_col:
            showError("Missing 'Factuurnummer' or 'Factuurnr' column.")
            return []

    except Exception as e:
        showError(f"Error reading the Excel file: {e}")
        return []

    docs = []
    for _, row in df.iterrows():
        try:
            datum = pd.to_datetime(row["Documentdatum"], errors='coerce')
            tebetalen = abs(row["Totaal brutto"])
            basisbedrag = abs(row["Totaal netto"])

            factuurnr = row[factuurnummer_col]
            print(factuurnr)
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
                vervaldatum=datum + pd.Timedelta(days=15)
            )
            docs.append(doc)
        except Exception as e:
            showError(f"Error processing row: {e}")
            continue

    docs.sort(key=lambda x: x.nummer)

    return docs
