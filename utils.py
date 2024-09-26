from tkinter import messagebox


def showWarning(message):
    messagebox.showwarning("Warning", message)


def showInfo(message):
    messagebox.showinfo("Info", message)


def showError(message):
    messagebox.showerror("Error", message)