import os
from openpyxl import Workbook
import numpy as np
import pandas as pd

from typing import List, Dict, Union

import difflib

# Note: Internal testing functions are prefixed with an underscore (_) and are not meant for production use

def _generate_excel(cols: int, rows: int, col_names: Union[List, None] = None, save_name: str = 'test.xlsx') -> int:
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
    
    # If column names are provided, make sure they match the number of columns
    if col_names is not None and len(col_names) != cols:
        print('[_generate_excel] Error: Number of column names does not match number of columns')
        return 1
        
    # Make sure the save name is valid
    if save_name == '':
        print('[_generate_excel] Error: Save name cannot be empty')
        return 1
    if not save_name.endswith('.xlsx'):
        save_name += '.xlsx'
        print(f'[_generate_excel] Warning: Save name does not end with .xlsx. Automatically appending .xlsx to save name')

    # Generate the excel file
    try:
        wb = Workbook()
        ws = wb.active

        # Generate column headers
        if col_names is None:
            for i in range(1, cols + 1):
                ws.cell(row=1, column=i).value = 'Column {}'.format(i)
        else:
            for i in range(1, cols + 1):
                ws.cell(row=1, column=i).value = col_names[i - 1]

        # Generate random numbers
        for i in range(2, rows + 2):
            for j in range(1, cols + 1):
                ws.cell(row=i, column=j).value = np.random.randint(1, 100)

        # Save the excel file
        wb.save(os.path.join(os.getcwd(), save_name))
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
    try:
        df = pd.read_excel(path)
    except Exception as e:
        print(f"[excel_to_list] Error: {e}")
        return []
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
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        result_dict = {}
        for column in df:
            result_dict[column] = df[column].tolist()
    except Exception as e:
        print(f"[excel_to_dict] Error: {e}")
        return {}
    return result_dict

def normalize_list(data: Union[List[str], str] = '') -> List[str]:
    '''
    Use a lambda function to normalize the list values by removing all non-alphanumeric characters and converting to lowercase
    Source: https://stackoverflow.com/questions/1276764/stripping-everything-but-alphanumeric-chars-from-a-string-in-python/1276779#1276779

    Args:
        data (Union[List, str]): List of strings or a single string
    
    Returns:
        List[str]: List of normalized strings
    '''

    # If the data is a string, convert it to a list
    if isinstance(data, str):
        # Check if the string is empty
        if data == '':
            print('[normalize_list] Warning: Input is an empty string. Returning empty list.')
            return []
        data = [data]

    return list(map(lambda x: ''.join(e for e in x if e.isalnum()).lower(), data))
    
def match_lists(list_1: List, list_2: List) -> Dict:
    '''
    Use diff to match two lists and returns a nested dictionary with the following structure:
    {
        'list_1_value': {
        'match': 'best_matching_string_in_list_2_here'
        'score': 'ratio_of_similarity_here'
        },
        'another_list_1_value': {
        'match': 'best_matching_string_in_list_2_here'
        'score': 'ratio_of_similarity_here'
        }
    }

    Example output:
    {
        'CO2e emissions': {
        'match': 'CO2e emissons (kg)'
        'score': '0.875'
        },
        'Waste generation': {
        'match': 'Waste genration (kg)'
        'score': '0.973'
        }
    }

    Args:
        list_1 (List): First list of strings
        list_2 (List): Second list of strings

    Returns:
        Dict: Output dictionary with the above structure
    '''
    output_dict = {}
    print(f"List 1: {list_1}")

    for item in list_1:
        best_match = 0
        best_match_string = ''

        for item_2 in list_2:
            ratio = difflib.SequenceMatcher(None, item, item_2).ratio()
            if ratio > best_match:
                best_match = ratio
                best_match_string = item_2

        output_dict[item] = {
            'match': best_match_string,
            'score': np.round(best_match, 3)
        }
    return output_dict