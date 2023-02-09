import os
from receiveinput import process_html_file
import pandas as pd
import openpyxl

folder_path = "C:/Users/blair/OneDrive/Documents/GitHub/Whole-Foods-Import/Whole Foods PDFs"

orders_compiled = []
for filename in os.listdir(folder_path):
    file_path = os.path.join(folder_path, filename)
    if os.path.isfile(file_path):
        sales_line = process_html_file(file_path)
        orders_compiled.append(sales_line)

# Read the Excel file into a DataFrame
df = pd.read_excel("Whole Foods Stores.xlsx")


def add_lists_together(data_list):
    sales_header = {
        'Document Type': [],
        'No.': [],
        'Sell-to Customer No.': [],
        'Ship-to Code': [],
        'Ship-to Name': [],
        'Ship-to Address': [],
        'Ship-to Address 2': [],
        'Ship-to City': [],
        'Order Date': [],
        'Shipment Date': [],
        'Location Code': [],
        'Salesperson Code': [],
        'Ship-to State': [],
        'Ship-to Country/Region Code': [],
        'External Document No.': []
    }

    sales_line = {
        'Document Type': [],
        'Document No.': [],
        'Line No.': [],
        'Type': [],
        'No.': [],
        'Location Code': [],
        'Shipment Date': [],
        'Description': [],
        'Quantity': [],
        'Unit Price': []
    }

    for data in data_list:
        sales_line['Document Type'] += data['Document Type']
        sales_line['Document No.'] += data['Document No.']
        sales_line['Line No.'] += data['Line No.']
        sales_line['Type'] += data['Type']
        sales_line['No.'] += data['No.']
        sales_line['Location Code'] += data['Location Code']
        sales_line['Shipment Date'] += data['Shipment Date']
        sales_line['Description'] += data['Description']
        sales_line['Quantity'] += data['Quantity']
        sales_line['Unit Price'] += data['Unit Price']
        sales_header['Document Type'] += ['Order']
        sales_header['No.'] += [data['Document No.'][0]]
        customer_number = int(data['Customer Number'])
        sales_header['Sell-to Customer No.'] += [customer_number]
        sales_header['Ship-to Code'] += ['']
        sales_header['Ship-to Name'] += [df.loc[df["Customer ID"] == customer_number, "Customer"].values[0]]
        sales_header['Ship-to Address'] += [df.loc[df["Customer ID"] == customer_number, 'Address Line 1'].values[0]]
        sales_header['Ship-to Address 2'] += [df.loc[df["Customer ID"] == customer_number, 'Address Line 2'].values[0]]
        sales_header['Ship-to City'] += [df.loc[df["Customer ID"] == customer_number, 'City'].values[0]]
        sales_header['Order Date'] += [data['Order Date']]
        sales_header['Shipment Date'] += [data['Shipment Date'][0]]
        sales_header['Location Code'] += ['ELIZABETH']
        sales_header['Salesperson Code'] += ['']
        sales_header['Ship-to State'] += [df.loc[df["Customer ID"] == customer_number, 'State'].values[0]]
        sales_header['Ship-to Country/Region Code'] += ['US']
        sales_header['External Document No.'] += [data['Order Number']]
    return sales_line, sales_header


sales_line_data, sales_header_data = add_lists_together(orders_compiled)

# Create two DataFrames
sales_line_df = pd.DataFrame(sales_line_data)
sales_header_df = pd.DataFrame(sales_header_data)

# Write each DataFrame to its own sheet in the same Excel file
with pd.ExcelWriter('orders.xlsx') as writer:
    sales_line_df.to_excel(writer, sheet_name='Sales Line', index=False)
    sales_header_df.to_excel(writer, sheet_name='Sales Header', index=False)


