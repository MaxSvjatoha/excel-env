import os
from openpyxl import Workbook
import numpy as np
import pandas as pd

from typing import List, Dict

# Note: Internal testing functions are prefixed with an underscore (_) and are not meant for production use

def _generate_excel(cols: int, rows: int) -> int:
    '''
    Generate an excel file with 'cols' many columns and 'rows' many rows. 
    Populate it with column headers and random numbers.

    Args:
        cols (int): Number of columns
        rows (int): Number of rows
    Returns:
        int: 0 if successful, 1 if not
    '''

    # Make sure the number of columns and rows are valid
    if cols < 1:
        print('[_generate_excel] Error: Number of columns must be greater than 0')
        return 1
    if rows < 1:
        print('[_generate_excel] Error: Number of rows must be greater than 0')
        return 1

    # Generate the excel file
    try:
        wb = Workbook()
        ws = wb.active

        for i in range(1, cols + 1):
            ws.cell(row=1, column=i).value = 'Column {}'.format(i)
        for i in range(2, rows + 2):
            for j in range(1, cols + 1):
                ws.cell(row=i, column=j).value = np.random.randint(1, 100)

        wb.save(os.path.join(os.getcwd(), 'test.xlsx'))
        return 0
    except Exception as e:
        print(e)
        return 1

def excel_to_list(path: str) -> List[List[str]]:
    '''
    Read an excel file and return a list of lists of strings

    Args:
        path (str): Path to the excel file
    Returns:
        List[List[str]]: List of lists of strings
    '''
    df = pd.read_excel(path)
    return df.values.tolist()

def excel_to_dict(file_path: str, sheet_name: str) -> Dict[str, list]:
    '''
    Read a single sheet in an Excel file and return a dictionary 
    with column headers as keys and the rest of the column as the values

    Args:
        file_path (str): Path to the Excel file
        sheet_name (str): Name of the sheet to read from

    Returns:
        dict: Dictionary with column headers as keys and the rest of the column
        as a list of values
    '''
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    result_dict = {}
    for column in df:
        result_dict[column] = df[column].tolist()
    return result_dict