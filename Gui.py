from excel_comparison import main, adjust_column_width
import tkinter as tk
from tkinter import ttk, filedialog, Text, Toplevel, Scrollbar, messagebox
import os
import threading
import queue
import csv
from openpyxl import Workbook


def select_file(entry):
    filename = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx *.csv')])
    if filename:
        entry.delete(0, tk.END)
        entry.insert(0, filename)


def save_to_file(data, output_file):
    _, ext = os.path.splitext(output_file)
    if ext == ".csv":
        with open(output_file, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(data)
    else:
        new_workbook = Workbook()
        new_sheet = new_workbook.active
        for row in data:
            new_sheet.append(row)
        adjust_column_width(new_sheet)
        new_workbook.save(output_file)


def compare_files(file1_entry, file2_entry, q, copy_only_new_clients=False, case_sensitive=False):
    file1 = file1_entry.get()
    file2 = file2_entry.get()
    if not file1 or not file2:
        q.put(("update_label", "Error: Please select both files."))
        return
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                               filetypes=[('Excel Files', '*.xlsx'), ('CSV Files', '*.csv')])
    if not output_file:
        q.put(("update_label", "Error: No save location selected."))
        return
    compare_button['state'] = 'disabled'
    threading.Thread(target=run_comparison,
                     args=(file1, file2, q, copy_only_new_clients, case_sensitive, output_file)).start()


def run_comparison(file1, file2, q, copy_only_new_clients, case_sensitive, output_file):
    try:
        data = main(file1, file2, q, copy_only_new_clients, case_sensitive)
        if data is not None:
            save_to_file(data, output_file)
    except Exception as e:
        q.put(("update_label", str(e)))
    finally:
        q.put(("update_button", "normal"))


def update_gui(q):
    try:
        while True:
            msg = q.get(0)
            cmd, arg = msg
            if cmd == "update_label":
                status_label['text'] = arg
            elif cmd == "update_progress":
                progress['value'] = arg
            elif cmd == "update_button":
                compare_button['state'] = arg
                copy_new_clients_button['state'] = arg
    except queue.Empty:
        pass
    finally:
        root.after(100, update_gui, q)


def display_log():
    log_window = Toplevel(root)
    log_window.title('excel_comparison.log')
    text_area = Text(log_window, wrap='word', height=15, width=50)
    text_area.grid(row=0, column=0, sticky='nsew')
    scrollbar = Scrollbar(log_window, command=text_area.yview)
    scrollbar.grid(row=0, column=1, sticky='ns')
    text_area['yscrollcommand'] = scrollbar.set
    try:
        with open('excel_comparison.log', 'r') as log_file:
            text_area.insert('1.0', log_file.read())
    except IOError:
        messagebox.showerror('Error', 'Could not open excel_comparison.log')
        if os.path.exists('excel_comparison.log'):
            os.remove('excel_comparison.log')
        logging.basicConfig(filename='excel_comparison.log', level=logging.INFO)


root = tk.Tk()
root.title("Excel File Comparison")

frame = ttk.Frame(root, padding="10")
frame.grid(row=0, column=0, sticky="NSEW")

file1_entry = ttk.Entry(frame, width=50)
file1_entry.grid(row=0, column=0, sticky="WE")
file1_button = ttk.Button(frame, text="Select File 1", command=lambda: select_file(file1_entry))
file1_button.grid(row=0, column=1, sticky=tk.W, padx=5)

file2_entry = ttk.Entry(frame, width=50)
file2_entry.grid(row=1, column=0, sticky="WE")
file2_button = ttk.Button(frame, text="Select File 2", command=lambda: select_file(file2_entry))
file2_button.grid(row=1, column=1, sticky=tk.W, padx=5)


compare_button = ttk.Button(frame, text="Update Database")
compare_button.grid(row=2, column=0, columnspan=2, pady=10)

copy_new_clients_button = ttk.Button(frame, text="Generate New Clients")
copy_new_clients_button.grid(row=3, column=0, columnspan=2, pady=10)

progress = ttk.Progressbar(frame, length=100)
progress.grid(row=4, column=0, columnspan=2, pady=10)

status_label = ttk.Label(frame, text="")
status_label.grid(row=5, column=0, columnspan=2)

case_sensitive_var = tk.BooleanVar()
case_sensitive_checkbutton = ttk.Checkbutton(frame, text="Case-sensitive comparison", variable=case_sensitive_var)
case_sensitive_checkbutton.grid(row=7, column=0, columnspan=2, pady=10)

compare_button['command'] = lambda: compare_files(file1_entry, file2_entry, q, False, case_sensitive_var.get())
copy_new_clients_button['command'] = lambda: compare_files(file1_entry, file2_entry, q, True, case_sensitive_var.get())

log_button = ttk.Button(frame, text="Show Log", command=display_log)
log_button.grid(row=6, column=0, columnspan=2, pady=10)

q = queue.Queue()
root.after(100, update_gui, q)

root.mainloop()
