import os
import sys

import json
import numpy as np

import utils.util as utils

# Setup: Import settings from settings.json and initialize required variables
with open("settings.json", "r") as f:
    settings = json.load(f)
    assert(type(settings) == dict)
    settings["Current working directory"] = os.getcwd()
    settings["Parent directory"] = os.path.dirname(settings["Current working directory"])
settings["Input file folder path"] = os.path.join(settings["Parent directory"], settings["Input file folder name"])
settings["Output file folder path"] = os.path.join(settings["Parent directory"], settings["Output file folder name"])

# Setup: Load relevant folder and file paths and extract their names
input_folder_paths = utils._get_input_folders(path = settings["Input file folder path"], settings = settings)
input_file_paths = utils._get_input_files(path = settings["Input file folder path"], settings = settings)
input_folder_names = [file.split("\\")[-1] for file in input_folder_paths]
input_file_names = [file.split("\\")[-1] for file in input_file_paths]

if __name__ == '__main__':
    summary_file = os.path.join(settings["Output file folder path"], settings["Summary file name"])

    # Get the summary file and its sheet names
    summary_wb = utils.excel_to_workbook(summary_file)
    summary_sheets = summary_wb.sheetnames

    # Match input file names to summary file sheet names
    matches = utils.match_lists(input_folder_names, summary_sheets, filter_doubles = True)
    # print(f"Matches: {matches}")

    # Create a dictionary of input data and initialize subdicts for each input file
    input_data_dict = {}
    for file in input_file_paths:
        input_data_key = file.split("\\")[-2] # Use folder name as key
        # Check that key is in matches dict
        if input_data_key in matches.keys():
            input_data_dict[input_data_key] = {} # Initialize subdict to store sheet data

    print("Processing input files...")
    for key in input_data_dict.keys():
        wb = utils.excel_to_workbook(file)
        scope_sheets = [sheet for sheet in wb.sheetnames if 'scope' in sheet.lower()]
        for sheet in scope_sheets:
            input_data_dict[key][sheet] = utils._get_scope_data(wb, sheet)
        wb.close()

    print("Writing data to summary file:")
    for key in input_data_dict.keys():
        # use 'matches' dict to find where to write the data
        print(f"Writing data from {key} to sheet {matches[key]['match']}")

        # load sheet in summary file which matches matches[key]['match']
        # print(f"Loading sheet {summary_wb[matches[key]['match']]} from summary file")
        summary_sheet = summary_wb[matches[key]['match']]
        max_row = summary_sheet.max_row
        max_col = summary_sheet.max_column

        match_count = 0 # Keep track of how many matches are found for a given subitem
        match_dict = {} # Keep track of which row and column a given subitem is found in

        # write data to sheet. format is:
        # input_data_dict[input_folder_name][summary_sheet_name][cell_with_name_associated_with_data]
        for item in input_data_dict[key].keys():
            for subitem in input_data_dict[key][item].keys():
                # Match subitem to row and column in summary sheet
                for row in range(1, max_row + 1):
                    for col in range(1, max_col + 1):
                        if summary_sheet.cell(row = row, column = col).value == subitem:
                            # print(f"Found match for {subitem} at row {row} and column {col}")
                            match_count += 1
                            match_dict[match_count] = {"row": row, "col": col}
                if match_count > 1:
                    if 'Scope 1'.lower() in item.lower():
                        write_key = min(match_dict.keys()) # use match_dict entry with the lowest key
                    elif 'Scope 3'.lower() in item.lower():
                        write_key = max(match_dict.keys()) # use match_dict entry with the highest key
                    else: 
                        pass # TODO # Handler (non-urgent)
                elif match_count == 0:
                    continue # go back to top of loop
                else:
                    write_key = list(match_dict.keys())[0] # use match_dict entry with the only key
                
                # Extract data from input_data_dict and write to summary sheet. If none, generate random data
                if input_data_dict[key][item][subitem] is not None:
                    write_data = input_data_dict[key][item][subitem] # Write data to summary sheet
                else:
                    write_data = np.random.randint(0, 100)
                
                summary_sheet.cell(row = match_dict[write_key]["row"], column = match_dict[write_key]["col"] + 1).value = write_data
                
                # Reset match_count and match_dict
                match_count = 0
                match_dict = {}

    summary_wb.save(os.path.join(settings['Output file folder path'], settings["Output file name"]))
    summary_wb.close()

    print("Script execution finished successfully")
    sys.exit(0)