import os
import sys

import json

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
    print(f"Matches: {matches}")

    # Create a dictionary of input data and initialize subdicts for each input file
    input_data_dict = {}
    for file in input_file_paths:
        input_data_key = file.split("\\")[-2] # Use folder name as key
        # Check that key is in matches dict
        if input_data_key in matches.keys():
            input_data_dict[input_data_key] = {} # Initialize subdict to store sheet data

    print("Processing input files: ")
    for key in input_data_dict.keys():
        wb = utils.excel_to_workbook(file)
        scope_sheets = [sheet for sheet in wb.sheetnames if 'scope' in sheet.lower()]
        for sheet in scope_sheets:
            input_data_dict[key][sheet] = utils._get_scope_data(wb, sheet)
        wb.close()

    for key in input_data_dict.keys():
        # use 'matches' dict to find where to write the data
        print(f"Writing data from {key} to sheet {matches[key]['match']}")
        # TODO: continue here

    #print(f"Saving summary file {settings['Output file name']} to {settings['Output file folder path']}")
    #summary_wb.save(os.path.join(settings['Output file folder path'], settings["Output file name"]))
    summary_wb.close()

    print("Script execution finished")