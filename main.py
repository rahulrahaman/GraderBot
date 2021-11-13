from openpyxl import load_workbook
import os
import webbrowser
import numpy as np


def get_values(col_name, row_range):
    values = []
    for i in range(row_range[0], row_range[1] + 1):
        val = sheet[f'{col_name}{i}'].value
        if val is None:
            val = ''
        values.append(val)
    return values


# The path to main directory, roster-grade file
main_directory = 'F:/Downloads/Linear-model-tutorial/'
excel_fname = ''
excel_file = os.path.join(main_directory, excel_fname)

# Tutorial name, here I am assuming the folder name and sheet names
# are given in this pattern for different tutorials.
tutorial_id = 10
tutorial_directory = f'Tut-{tutorial_id}-linear-models'
sheet_name = f'Tutorial {tutorial_id}'

# Define the columns and row-range for ids, define column for marks
ids_stored_at_col = 'C'
ids_stored_at_row = [2, 77]
store_grade_at_col = 'D'

# Get all the list of submission files from tutorial folder
path = os.path.join(main_directory, tutorial_directory)
fnames = os.listdir(path)
file_list = [os.path.join(path, fname) for fname in fnames]

# Load the student-roster excel file for reading and editing
wb = load_workbook(excel_file)
sheet = wb[sheet_name]

# Go through all files and do stuff
for i in range(len(fnames)):
    # Get ID e.t.c from fine name
    filen = fnames[i]
    pathn = file_list[i]
    id_ = filen[:9]

    # Search ID in the excel with the student roster
    for rowid in range(ids_stored_at_row[0], ids_stored_at_row[1]+1):
        id_cell = f'{ids_stored_at_col}{rowid}'
        id_at_cell = sheet[id_cell].value

        # If a match is obtained, print for validation
        if id_at_cell == id_:
            print(f'Found at row {rowid} with student ID {id_at_cell}')
            break

    # Check if the item is already graded, if yes then skip
    grade_cell = f'{store_grade_at_col}{rowid}'
    grade_at_cell = sheet[grade_cell].value
    if grade_at_cell is not None or grade_at_cell == '':
        continue

    # Open the submission file for grading. Once done, input the marks
    webbrowser.open(pathn)
    print(f'ID {id_}: Please input the grade')
    grade = input()

    # Store the marks to the corresponding row and grade column, save the file
    store_at = f'{store_grade_at_col}{rowid}'
    sheet[store_at] = int(grade)
    wb.save(filename=excel_file)
