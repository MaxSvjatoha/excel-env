import os
import sys

import utils.util as utils

import numpy as np
import pandas as pd

import openpyxl

if __name__ == '__main__':
    
    # get cwd
    cwd = os.getcwd()

    # get cwd's parent
    parent = os.path.dirname(cwd)

    input_folder_name = 'Arbetsmapp datainsamling'

    # get path to input folder
    input_folder_path = os.path.join(parent, input_folder_name)

    print(f"input_folder_path: {input_folder_path}")

    # get all files in input folder
    input_folder_subdirs = os.listdir(input_folder_path)

    print(f"input_folder_subdirs: {input_folder_subdirs}")
    
    # Enter first subfolder and get all excel files
    first_subfolder = input_folder_subdirs[0]
    first_subfolder_path = os.path.join(input_folder_path, first_subfolder)
    first_subfolder_files = os.listdir(first_subfolder_path)
    first_subfolder_excel_files = [file for file in first_subfolder_files if file.endswith('.xlsx')]
    print(f"first_subfolder_excel_files: {first_subfolder_excel_files}")
    
    # Read first excel file
    first_excel_file = first_subfolder_excel_files[0]
    first_excel_file_path = os.path.join(first_subfolder_path, first_excel_file)

    # Now use openpyxl to read all sheets in first excel file and save them in cwd
    wb = openpyxl.load_workbook(first_excel_file_path)
    sheets = wb.sheetnames
    print(f"sheets: {sheets}")
    
    # Filter out sheets that are not relevant - keep only those with "scope" in the name
    sheets = [sheet for sheet in sheets if 'scope' in sheet.lower()]
    print(f"sheets: {sheets}")

    # Print all workbook data
    for sheet in sheets:
        ws = wb[sheet]
        print(f"ws: {ws}")
        for row in ws.iter_rows():
            for cell in row:
                print(cell.value, type(cell.value))
    
    # Save entire workbook to cwd
    utils.workbook_to_excel(input_workbook = wb, save_name = 'test.xlsx')

    sys.exit(0)
    
    
    
        