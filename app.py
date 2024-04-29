import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog
import openpyxl
from extract import extract_tickets
from classify import classify_tickets
import os


def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    workbook = openpyxl.load_workbook(file_path)
    sheet_combo["values"] = workbook.sheetnames
    sheet_combo.current(0)
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)


def select_output_folder():
    folder_path = filedialog.askdirectory()
    output_entry.delete(0, tk.END)
    output_entry.insert(0, folder_path)


def update_progress(v):
    progress_bar["value"] = v
    root.update_idletasks()


def reset_form():
    file_entry.delete(0, tk.END)
    sheet_combo.set("")
    sheet_combo["values"] = []
    type_combo.current(0)
    output_entry.delete(0, tk.END)
    progress_bar["value"] = 0


def process_tickets():
    excel_file = file_entry.get()
    selected_type = type_combo.get()
    selected_sheet = sheet_combo.get()
    output_folder = output_entry.get()
    if (
        excel_file != ""
        and selected_type != ""
        and selected_sheet != ""
        and output_folder != ""
    ):
        extracted_file = "./extracted_" + os.path.basename(excel_file)
        final_output = output_folder + "/tagged_" + os.path.basename(excel_file)
        update_progress(10)
        tickets_processed = extract_tickets(excel_file, selected_sheet, extracted_file)
        update_progress(50)
        classify_tickets(extracted_file, selected_type, final_output)
        update_progress(80)
        os.remove(extracted_file)
        update_progress(100)
    else:
        tickets_processed = "No"

    # dialog = tk.Toplevel(root)
    # dialog.title("Success")
    # message = tk.Label(dialog, text=f"{tickets_processed} Tickets Processed")
    # message.pack(padx=20, pady=10)
    # exit_button = ttk.Button(dialog, text="Exit", command=root.quit)
    # exit_button.pack(pady=10)
    # dialog.update_idletasks()

    dialog = tk.Toplevel(root)
    dialog.title("Success")
    message = tk.Label(dialog, text=f"{tickets_processed} Tickets Processed")
    message.pack(padx=20, pady=10)

    ok_button = ttk.Button(
        dialog, text="OK", command=lambda: [reset_form(), dialog.destroy()]
    )
    ok_button.pack(pady=10)

    exit_button = ttk.Button(dialog, text="Exit App", command=root.quit)
    exit_button.pack(pady=10)

    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()
    x = (screen_width // 2) - (dialog.winfo_reqwidth() // 2)
    y = (screen_height // 2) - (dialog.winfo_reqheight() // 2)
    dialog.geometry(f"+{int(x)}+{int(y)}")
    dialog.update_idletasks()


root = ttk.Window(title="FreshDesk Ticket Tagger", resizable=(False, False))
root.geometry("800x250")

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

x = (screen_width // 2) - (root.winfo_reqwidth() // 2)
y = (screen_height // 2) - (root.winfo_reqheight() // 2)

root.geometry(f"+{int(x)}+{int(y)}")

frame = ttk.Frame(root)
frame.pack(padx=10, pady=10)

file_label = ttk.Label(frame, text="Select Excel File:", width=20)
file_label.grid(row=0, column=0, sticky="w", pady=5)
file_entry = ttk.Entry(frame, width=50)
file_entry.grid(row=0, column=1, sticky="w", pady=5)
file_button = ttk.Button(
    frame, text="Browse", command=browse_file, style="primary.TButton"
)
file_button.grid(row=0, column=2, sticky="w", pady=5)

type_label = ttk.Label(frame, text="Select Type:", width=20)
type_label.grid(row=1, column=0, sticky="w", pady=5)
type_combo = ttk.Combobox(frame, values=["FSD", "Data Science"])
type_combo.grid(row=1, column=1, sticky="w", pady=5)
type_combo.current(0)

sheet_label = ttk.Label(frame, text="Select Sheet:", width=20)
sheet_label.grid(row=2, column=0, sticky="w", pady=5)
sheet_combo = ttk.Combobox(frame)
sheet_combo.grid(row=2, column=1, sticky="w", pady=5)

output_label = ttk.Label(frame, text="Output Folder:", width=20)
output_label.grid(row=3, column=0, sticky="w", pady=5)
output_entry = ttk.Entry(frame, width=50)
output_entry.grid(row=3, column=1, sticky="w", pady=5)
output_button = ttk.Button(
    frame, text="Browse", command=select_output_folder, style="primary.TButton"
)
output_button.grid(row=3, column=2, sticky="w", pady=5)

process_button = ttk.Button(
    frame, text="Process Tickets", command=process_tickets, style="success.TButton"
)
process_button.grid(row=4, column=1, pady=10, sticky="w")


progress_bar = ttk.Progressbar(
    frame,
    orient="horizontal",
    length=800,
    mode="determinate",
    bootstyle="warning-striped",
)
progress_bar.grid(row=5, column=0, columnspan=6, pady=10, sticky="w")
progress_bar["maximum"] = 100


root.mainloop()
