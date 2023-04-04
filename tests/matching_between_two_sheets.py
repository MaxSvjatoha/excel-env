import sys

# Settings
save_name_1 = 'input_test1.xlsx'
save_name_2 = 'input_test2.xlsx'

# Allows imports from sibling directories
# Source: https://stackoverflow.com/questions/70395407/import-module-from-a-sibling-directory-in-python3-10/73081295#73081295
sys.path.insert(0, '.')

import utils.util as utils

if __name__ == '__main__':
    # Generate 2 excel files with different (but similar) column names, read them and attempt to merge them into a single file based on the column names

    # Create column names with a climate impact theme
    col_names = ['CO2e emissions', 'Energy consumption', 'Water consumption', 'Waste generation', 'Land use']

    # Generate the first excel file
    utils._generate_excel(5, 5, col_names=col_names, save_name=save_name_1)

    # Generate the second excel file where each column name is similar, but not identical to the first excel file
    col_names = ['CO2e emissions (kg)', 'Energy consumption (kWh)', 'Water consumption (L)', 'Waste generation (kg)', 'Land use (m2)']

    # Add misspellings to some of the column names
    col_names[0] = 'CO2e emissons (kg)'
    col_names[3] = 'Waste genration (kg)'

    # Generate the second excel file
    utils._generate_excel(5, 5, col_names=col_names, save_name=save_name_2)

    # Read the excel files (as dictionaries)
    data_dict_1 = utils.excel_to_dict(save_name_1, 'Sheet')
    data_dict_2 = utils.excel_to_dict(save_name_2, 'Sheet')

    # Extract and normalize the keys
    keys_1, mappings_1 = utils.normalize_list(data = list(data_dict_1.keys()), keep_originals = True)
    keys_2, mappings_2 = utils.normalize_list(data = list(data_dict_2.keys()), keep_originals = True)

    # Match the two sets of keys
    matches = utils.match_lists(list_1 = keys_1, list_2 = keys_2)

    # Check for the following condition - mappings_2[key_2] == matches[mappings_1[key_1]]
    for key_1, key_2 in zip(mappings_1, mappings_2):
        if mappings_2[key_2] == matches[mappings_1[key_1]]['match']:
            # Extend the first dictionary with the second dictionary
            data_dict_1[key_1] += (data_dict_2[key_2])

    utils.dict_to_excel(data_dict_1)