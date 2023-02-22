import openpyxl
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import openpyxl
from data_creation.df_creation_counts import create_counts_df
import pandas as pd
from table_creation.table_creation_counts import add_dataframe_to_excel_sheet
from table_creation.table_creation_severity import table_creation_for_severity
from table_creation.table_creation_conformance import table_creation_for_conformance
from table_creation.table_creation_issuetype import table_creation_for_issuetype
from table_creation.table_creation_status import table_creation_for_status
from table_creation.table_creation_status_counts import table_creation_for_status_counts

# Create the GUI
root = Tk()
root.title("Report Automation")

# Create a custom style
style = ttk.Style()
style.theme_create("custom_style", parent="alt", settings={
    "TNotebook": {"configure": {"background": "#b9d9eb"}},
    "TNotebook.Tab": {"configure": {"background": "#b9d9eb", "foreground": "#000000", "padding": [10, 5]}, "map": {"background": [("selected", "#ffffff")]}},
    "TFrame": {"configure": {"background": "#b9d9eb"}},
    "TLabel": {"configure": {"background": "#b9d9eb"}},
    "TButton": {"configure": {"background": "#ffffff", "foreground": "#000000", "font": ("Arial", 14), "padding": [10, 5], "borderwidth": 0, "relief": "flat", "border-radius": 20}, "map": {"background": [("active", "#b9d9eb")]}},
    "TEntry": {"configure": {"font": ("Arial", 14), "padding": [10, 5]}, "map": {"background": [("active", "#ffffff")]}},
})
style.theme_use("custom_style")

# Function to handle file selection
def select_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        # Clear the previous page
        for widget in root.winfo_children():
            widget.destroy()

        # Display the workbook name
        wb_name = ttk.Label(root, text="Workbook Location : {}".format(file_path))
        wb_name.pack(pady=20)

        # Create a new page with the sheet input field and submit button
        input_frame = ttk.Frame(root, style="TFrame")
        input_frame.pack(pady=20)
        sheet_label = ttk.Label(input_frame, text="Sheet Name:")
        sheet_label.grid(row=0, column=0, padx=10)
        sheet_entry = ttk.Entry(input_frame, width=30)
        sheet_entry.insert(END, "Data & Chart")
        sheet_entry.grid(row=0, column=1, padx=10)
        submit_button = ttk.Button(root, text="Submit", command=lambda: submit_file(file_path, sheet_entry.get()))
        submit_button.pack(pady=20)

        # Center the new page
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry('{}x{}+{}+{}'.format(width, height, x, y))

# Function to handle file submission
def submit_file(file_path, sheet_name):
    wb = openpyxl.load_workbook(file_path ,data_only=True)
    if sheet_name in wb.sheetnames:
        response = messagebox.askyesno("Sheet Exists", "The sheet {} already exists in the selected file. Do you want to delete it and continue?".format(sheet_name))
        if response == True:
            wb.remove(wb[sheet_name])
        else:
            return

    df , new_df = create_counts_df(wb)
    last_row = add_dataframe_to_excel_sheet(wb , sheet_name ,  new_df)
# code Ended to add the Counts Table and Chart to the Page

    table_creation_for_severity(wb, sheet_name, last_row)
    table_creation_for_conformance(wb, sheet_name, last_row + 15)
    table_creation_for_issuetype(wb, sheet_name, last_row + 30)
    table_name = table_creation_for_status(wb, sheet_name, df , last_row + 52)
    table_creation_for_status_counts(wb, sheet_name, table_name, last_row + 45)
    wb.save(file_path)

    # Display a status message
    status_label = ttk.Label(root, text="Changes saved.")
    status_label.pack()

    # Close the program after 3 seconds
    root.after(1000, root.destroy)

# Add a button to select a file
select_button = ttk.Button(root, text="Select File", command=select_file)
select_button.place(relx=0.5, rely=0.5, anchor=CENTER)
select_button.config(width=20)

root.geometry("700x300")

# Run the GUI
root.mainloop()
   
