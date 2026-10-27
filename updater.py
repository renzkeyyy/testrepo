import openpyxl

file_path = "test_data.xlsx"

try:
    # loads the workbook and checks for the 'RawData' tab
    workbook = openpyxl.load_workbook(file_path)
    print("Workbook loaded successfully.")

    if 'RawData' in workbook.sheetnames:
        sheet = workbook['RawData']
        print(f"Success! Located the 'RawData' tab in {file_path}.")
        
        # quick test to check if the sheet is accessible
        print(f"Cell A1 contains: {sheet['A1'].value}")
    else:
        print(f"Error: 'RawData' tab not found in {file_path}.")

except PermissionError:
    # when file is open in Excel
    print(f"Warning: {file_path} is currently open. Please close Excel and try again.")
except FileNotFoundError:
    # when file is not found
    print(f"Error: Could not find {file_path}.")
except Exception as e:
    # catch-all for any other exceptions
    print(f"An unexpected error occurred: {e}")
