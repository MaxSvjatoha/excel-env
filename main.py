import os
import sys

import json
import numpy as np
import openpyxl

import itertools

import utils.util as utils

# Setup: Import settings from settings.json and initialize required variables
with open("settings.json", "r") as f:
    settings = json.load(f)
    assert(type(settings) == dict)
    settings["Current working directory"] = os.getcwd()
    settings["Parent directory"] = os.path.dirname(settings["Current working directory"])
settings["Input file folder path"] = os.path.join(settings["Parent directory"], settings["Input file folder name"])
settings["Output file folder path"] = os.path.join(settings["Parent directory"], settings["Output file folder name"])

# Setup: Import scope 2 special cases dictionary
with open("scope_2_dict.json", "r") as f:
    scope_2_dict = json.load(f)
    assert(type(scope_2_dict) == dict)

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
    if "Missmatchningar" not in summary_sheets:
        summary_mismatches = summary_wb.create_sheet("Missmatchningar")
    else:
        pass # TODO - decide whether new sheet should be created or if starting row should be set to max_row + 1

    mismatch = False # Track match state
    mismatch_count = 0 # Keep track of how many mismatches are found in total
    mismatch_dict = {} # Keep track of which subitems are not found in the summary sheet

    summary_mismatches.cell(row = 1, column = 1).value = "Input folder name"
    summary_mismatches.cell(row = 1, column = 2).value = "Scope"
    summary_mismatches.cell(row = 1, column = 3).value = "Entry name"
    summary_mismatches.cell(row = 1, column = 4).value = "Value"

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
        write_key = None # Keep track of which match_dict entry to use when writing data
        write_data = None # Keep track of which data to write to the summary sheet
        
        coords_dict = {} # EXPERIMENTAL: keep track of the scanned cell value and its coordinates here

        # write data to sheet. format is:
        # input_data_dict[input_folder_name][scope][summary_sheet_name]
        for item in input_data_dict[key].keys():
            for subitem in input_data_dict[key][item].keys():
                for row, col in itertools.product(range(1, max_row + 1), range(1, max_col + 1)):

                    # Load cell value
                    cell_value = summary_sheet.cell(row = row, column = col).value
                    
                    # Pre-process cell value
                    if type(cell_value) == str:
                        cell_value = utils.preprocess_cell(cell_value)

                    # Try to match subitem to cell_value
                    if subitem == cell_value:
                        print(f"Found match for {subitem} at row {row} and column {col}")
                        match_count += 1
                        match_dict[match_count] = {"row": row, "col": col}
                    elif str(subitem) in str(cell_value):
                        print(f"{subitem} part of {cell_value} at row {row} and column {col}")
                        match_count += 1
                        match_dict[match_count] = {"row": row, "col": col}
                    else:
                        #print(f"No match for {subitem} at row {row} and column {col}")
                        pass

                # Handle various amounts of matches
                if match_count > 1:
                    if 'Scope 1'.lower() in item.lower():
                        write_key = min(match_dict.keys()) # use match_dict entry with the lowest key
                    elif 'Scope 3'.lower() in item.lower():
                        write_key = max(match_dict.keys()) # use match_dict entry with the highest key
                    else:
                        print(f"WARNING: Multiple matches found for {subitem} in sheet {matches[key]['match']}") 
                        pass # TODO # Handler (non-urgent)
                elif match_count == 0:
                    
                    # Handle special cases before registering mismatches
                    # NOTE - this could be a preprocessing step before the row/col scan instead
                    # NOTE 2 - this should be refactored and preferably done in a separate function

                    # Electricity 
                    if all(x in subitem.lower() for x in ['källa', 'inköpt el']):
                        print(f"found 'källa' and 'inköpt el' in subitem: {subitem}, using special case 1")
                        special_case = scope_2_dict["1"]
                        match_dict[special_case["name"]] = {"row": special_case["row"], "col": special_case["col"]}
                        write_key = special_case["name"]
                    elif all(x in subitem.lower() for x in ['kwh', 'elanvändning']):
                        print(f"found 'kwh' and 'elanvändning' in subitem: {subitem}, using special case 2")
                        special_case = scope_2_dict["2"]
                        match_dict[special_case["name"]] = {"row": special_case["row"], "col": special_case["col"]}
                        write_key = special_case["name"]
                    # Heat
                    elif all(x in subitem.lower() for x in ['källa', 'värme']):
                        print(f"found 'källa' and 'värme' in subitem: {subitem}, using special case 3")
                        special_case = scope_2_dict["3"]
                        match_dict[special_case["name"]] = {"row": special_case["row"], "col": special_case["col"]}
                        write_key = special_case["name"]
                    elif all(x in subitem.lower() for x in ['kwh', 'värme']):
                        print(f"found 'kwh' and 'värme' in subitem: {subitem}, using special case 4")
                        special_case = scope_2_dict["4"]
                        match_dict[special_case["name"]] = {"row": special_case["row"], "col": special_case["col"]}
                        write_key = special_case["name"]
                    # Cooling
                    elif all(x in subitem.lower() for x in ['källa', 'kyla']):
                        print(f"found 'källa' and 'kyla' in subitem: {subitem}, using special case 5")
                        special_case = scope_2_dict["5"]
                        match_dict[special_case["name"]] = {"row": special_case["row"], "col": special_case["col"]}
                        write_key = special_case["name"]
                    elif all(x in subitem.lower() for x in ['kwh', 'kyla']):
                        print(f"found 'kwh' and 'kyla' in subitem: {subitem}, using special case 6")
                        special_case = scope_2_dict["6"]
                        match_dict[special_case["name"]] = {"row": special_case["row"], "col": special_case["col"]}
                        write_key = special_case["name"]
                    else:
                        # Register mismatch
                        #print(f"WARNING: No match found for {subitem} in sheet {matches[key]['match']}")
                        mismatch_count += 1
                        if key not in mismatch_dict.keys():
                            mismatch_dict[key] = {}
                        if item not in mismatch_dict[key].keys():
                            mismatch_dict[key][item] = {}
                        if subitem not in mismatch_dict[key][item].keys():
                            mismatch_dict[key][item][subitem] = mismatch_count
                            # write to mismatch sheet
                            summary_mismatches.cell(row = mismatch_count+1, column = 1).value = key
                            summary_mismatches.cell(row = mismatch_count+1, column = 2).value = item
                            summary_mismatches.cell(row = mismatch_count+1, column = 3).value = subitem
                            summary_mismatches.cell(row = mismatch_count+1, column = 4).value = "No match found"
                        else:
                            print(f"WARNING: {subitem} already in mismatch_dict")
                            pass # TODO # Handler? (this bit was never reached in testing)
                        continue
                else:
                    write_key = list(match_dict.keys())[0] # use match_dict entry with the only key
                
                # Extract data from input_data_dict and write to summary sheet. If none, generate random data
                if input_data_dict[key][item][subitem] is not None:
                    write_data = input_data_dict[key][item][subitem] # Write data to summary sheet
                else:
                    write_data = "GENERATED: " + str(np.random.randint(0, 100))
                
                # Write data to summary sheet if write_key is not None
                if write_key is not None:
                    summary_sheet.cell(row = match_dict[write_key]["row"], column = match_dict[write_key]["col"] + 1).value = write_data
                
                # Reset match_count and match_dict
                match_count = 0
                match_dict = {}

    summary_wb.save(os.path.join(settings['Output file folder path'], settings["Output file name"]))
    summary_wb.close()

    print("Mismatched subitems:")
    for key in mismatch_dict.keys():
        print(f"Key: {key}, value: {mismatch_dict[key]}")
    print(f"Total number of mismatches: {mismatch_count}")

    print("Script execution finished successfully")
    sys.exit(0)