import sys
import os

# Allows imports from sibling directories
# Source: https://stackoverflow.com/questions/70395407/import-module-from-a-sibling-directory-in-python3-10/73081295#73081295
sys.path.insert(0, '.')

import utils.util as utils

if __name__ == '__main__':
    # Generate an excel file
    if utils._generate_excel(5, 5) == 0:
        print('Excel file generated successfully')
    else:
        print('Failed to generate excel file')
        sys.exit(1)

    # Read the excel file
    data = utils.excel_to_list('test.xlsx')
    print(data)

    # Convert the excel file to a dictionary
    data_dict = utils.excel_to_dict('test.xlsx', 'Sheet')
    print(data_dict)