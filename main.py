import openpyxl
from openpyxl import Workbook

def create_excel_file(file_name):
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    
    # Add headers
    ws.append(["Account Number", "Date", "Dollar Amount"])
    
    wb.save(file_name)
    print(f"Excel file '{file_name}' created with headers.")

def add_data_to_excel(file_name, account_number, date, dollar_amount):
    wb = openpyxl.load_workbook(file_name)
    ws = wb["Data"]
    
    # Append data to the sheet
    ws.append([account_number, date, dollar_amount])
    
    wb.save(file_name)
    print(f"Data added to '{file_name}': {account_number}, {date}, {dollar_amount}")

# Usage example
file_name = "account_data.xlsx"

# Create the Excel file with headers (run this once)
create_excel_file(file_name)

# Add data to the Excel file
add_data_to_excel(file_name, "123456", "2024-07-20", 250.75)
add_data_to_excel(file_name, "789012", "2024-07-21", 145.00)
add_data_to_excel(file_name, "345678", "2024-07-22", 320.40)
