import pandas as pd
import sys
import tkinter as tk
from tkinter import filedialog
import time

print("Select the database files DB1, DB2, DB3 etc")
time.sleep(3)

# create tkinter root window (it won't be shown)
root = tk.Tk()
root.withdraw()

# show file selection dialog for multiple files
selected_db_files = filedialog.askopenfilenames(title="Select the database files DB2, DB3 etc",
                                                filetypes=[("All Files", "*.xlsx")])

# check if user cancelled
if not selected_db_files:
    print("DB file(s) selection cancelled")
    sys.exit()
else:
    print(str(len(selected_db_files)) + " DB file(s) selected")

# data importation
print("Importing the DB files. This will take a few minutes ...")

# import data
tables = []

for i in range(len(selected_db_files)):
    tables.append(pd.read_excel(selected_db_files[i], engine='openpyxl'))


print("Upload of DB1, DB2, DB3 etc completed successfully!")
