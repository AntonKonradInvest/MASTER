import pandas as pd
from tkinter import filedialog, messagebox
import tkinter as tk


# Function to select the CSV file
def select_csv_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    return file_path


# Function to save the cleaned CSV file
def save_cleaned_csv_file(cleaned_df):
    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if save_path:
        cleaned_df.to_csv(save_path, index=False, sep=';', encoding='utf-8')
        messagebox.showinfo("Success", f"Cleaned CSV file saved at: {save_path}")
    else:
        messagebox.showwarning("Warning", "No save location selected.")


# Function to clean the CSV file (remove duplicates)
def clean_csv_file(file_path):
    try:
        # Load the CSV file
        df = pd.read_csv(file_path, delimiter=';', encoding='utf-8')

        # Ensure it has exactly 2 columns ('Name' and 'Relatiecode')
        if list(df.columns) != ['Name', 'Relatiecode']:
            raise ValueError("CSV file must have exactly 2 columns: 'Name' and 'Relatiecode'")

        # Remove duplicate rows
        cleaned_df = df.drop_duplicates()

        # Save the cleaned CSV file
        save_cleaned_csv_file(cleaned_df)

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


# Main function to select, clean, and save the CSV file
def process_csv_file():
    # Select the reference CSV file
    file_path = select_csv_file()
    if not file_path:
        messagebox.showwarning("Warning", "No file selected.")
        return

    # Clean the CSV file
    clean_csv_file(file_path)


# GUI setup using tkinter
root = tk.Tk()
root.title("CSV File Cleaner")

# GUI button to trigger the file processing
button = tk.Button(root, text="Select and Clean CSV File", command=process_csv_file, width=30, height=2)
button.pack(pady=20)

# Start the GUI loop
root.mainloop()
