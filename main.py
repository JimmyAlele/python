import pandas as pd
import sys
import tkinter as tk
from tkinter import filedialog
import time
import openpyxl

print("Select the database files DB1, DB2, DB3 etc")
# time.sleep(0)

# create tkinter root window (it won't be shown)
root = tk.Tk()
root.withdraw()

# show file selection dialog for multiple files
selected_db_files = filedialog.askopenfilenames(title="Select the database files DB1, DB2, DB3 etc",
                                                filetypes=[("All Files", "*.xlsx")])

# check if user cancelled
if not selected_db_files:
    print("DB file(s) selection cancelled")
    sys.exit()
else:
    print(str(len(selected_db_files)) + " DB file(s) selected")

# data importation
print("Importing the DB files. This will take a few minutes ...")
print()

# import data
tables = []

start_time = time.time()
for i in range(len(selected_db_files)):
    tables.append(pd.read_excel(selected_db_files[i], engine='openpyxl'))

print("DB1, DB2, DB3 etc imported successfully!")
end_time = time.time()
elapsed_time = (end_time - start_time) / 60

print("File importation took " + str(elapsed_time) + " minutes.")
print()

# count of dfs
table_index = 0
count_tables = len(selected_db_files)

start_time_loop = time.time()
for tbl in tables:
    # table_index is used to access workbook path from selected_db_files variable
    # table_index is updated at the end of for loop
    workbook_path = selected_db_files[table_index]
    workbook = openpyxl.load_workbook(workbook_path)
    sheet = workbook['Sheet1']
    new_workbook_path = workbook_path[:-5] + "_updated.xlsx"

    # check main_tbl and save column names with filters
    # check for >= or <= and issue warning
    col_names = []
    for col in tbl.columns:
        first_value = tbl.loc[0, col]
        if (">=" in str(first_value)) & ("<=" in str(first_value)):
            print("Select the proper filter condition. Do not leave it as >= or <=")
            print("WARNING! " + str(col) + " column filter will be ignored")
            print("")
        elif ">=" in str(first_value):
            col_names.append(col)
        elif "<=" in str(first_value):
            col_names.append(col)
        else:
            pass

    # to hold tables to be filtered
    check_tables = []

    # listing tables to be filtered
    for count in range(count_tables):
        if count != table_index:
            check_tables.append(count)

    # actual Excel indexing
    start_col = 62

    for check_tbl in check_tables:
        filtered_tbl = tables[check_tbl].copy()

        # first row index 0 holds conditions. So avoid it
        for row in range(1, len(tbl)):
            # time filter
            work_tbl = filtered_tbl[filtered_tbl['Time'] == tbl.loc[row, 'Time']].copy()

            # loop through the column names in tbl and check for filters
            for col in col_names:
                first_value = tbl.loc[0, col]
                filter_value = tbl.loc[row, col]

                # check if value of tbl row is blank or not
                # ignore if blank
                if pd.notna(filter_value):
                    # check if the filter value is numeric or string
                    if isinstance(filter_value, (int, float)):
                        if ">=" in str(first_value):
                            work_tbl['Check'] = pd.to_numeric(work_tbl[col], errors="coerce")
                            work_tbl = work_tbl[
                                (work_tbl['Check'] >= filter_value) & (work_tbl['Check'].notna())].copy()
                            work_tbl.drop('Check', axis=1, inplace=True)
                        elif "<=" in str(first_value):
                            work_tbl['Check'] = pd.to_numeric(work_tbl[col], errors="coerce")
                            work_tbl = work_tbl[
                                (work_tbl['Check'] <= filter_value) & (work_tbl['Check'].notna())].copy()
                            work_tbl.drop('Check', axis=1, inplace=True)

                    elif isinstance(filter_value, str):
                        work_tbl = work_tbl[work_tbl[col] == filter_value].copy()

                    # if it is not a string, and it is not an int
                    else:
                        print("WARNING! Check filter value of row " + str(row) + " and column " + str(
                            col) + " in DB" + str(check_tbl + 1))
                        print("The filter value has been ignored")
                        print()
                else:
                    pass

            # looping through rows and filtering finished.
            # update the Excel file
            sheet.cell(row=row + 2, column=start_col, value="Updated")

        print("check_tbl " + str(check_tbl))

        # actual Excel indexing
        start_col = 80

    # save and close the workbook
    workbook.save(new_workbook_path)
    workbook.close()

    # next table_index
    table_index += 1

end_time_loop = time.time()
elapsed_time = (end_time_loop - start_time_loop) / 60
print("Filtering took " + str(elapsed_time) + " minutes.")
