import utils.util as utils

if __name__ == '__main__':
    # Generate an excel file
    if utils._generate_excel(5, 5) == 0:
        print('Excel file generated successfully')
    else:
        print('Failed to generate excel file')
        exit(1)

    # Read the excel file
    data = utils.excel_to_list('test.xlsx')
    print(data)

    # Convert the excel file to a dictionary
    data_dict = utils.excel_to_dict('test.xlsx', 'Sheet')
    print(data_dict)