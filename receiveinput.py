from bs4 import BeautifulSoup
import pandas as pd


# Read the Excel file into a DataFrame
df = pd.read_excel("Whole Foods Products.xlsx")


# Function to look up the "Name to Use", "Description", and "Type" based on the "Whole Foods Name"
def lookup_name_to_use(value):
    item_code = df.loc[df["Whole Foods Name"] == value.upper()]["Name to Use"].values[0]
    item_description = df.loc[df["Whole Foods Name"] == value.upper()]["Description"].values[0]
    item_type = df.loc[df["Whole Foods Name"] == value.upper()]["Type"].values[0]
    return item_code, item_description, item_type


def extract_order_data(data_file):
    order_data = {}
    order_data["Order Number"] = data_file.select_one("td:-soup-contains('Order Number:') + td").text
    order_data["Order Date"] = data_file.select_one("td:-soup-contains('Order Date:') + td").text
    order_data["Expected Delivery Date"] = data_file.select_one("td:-soup-contains('Expected Delivery Date:') + td").text
    order_data["Store No"] = data_file.select_one("td:-soup-contains('Store No:') + td").text
    order_data["Customer ID"] = data_file.select_one("td:-soup-contains('Account No:') + td").text
    order_data["Subteam"] = data_file.select_one("td:-soup-contains('Subteam:') + td").text
    order_data["Buyer"] = data_file.select_one("td:-soup-contains('Buyer:') + td").text
    return order_data


def extract_item_data(soup):
    item_data = []
    for row in soup.find_all('tr'):
        cells = row.find_all('td')
        if len(cells) == 7:
            item = {
                'code': cells[1].text.strip(),
                'quantity': int(cells[2].text.strip().split("\xa0")[0]),
                'cost': float(cells[5].text.strip()),
            }
            item_data.append(item)
    return item_data


def process_data(data):
    item_types = {
        'Tea': {'codes': [], 'descriptions': [], 'quantities': [], 'costs': [], 'type': []},
        'Gus': {'codes': [], 'descriptions': [], 'quantities': [], 'costs': [], 'type': []},
        'Gorgie': {'codes': [], 'descriptions': [], 'quantities': [], 'costs': [], 'type': []},
        'Chip': {'codes': [], 'descriptions': [], 'quantities': [], 'costs': [], 'type': []}
    }
    for item in data:
        code, description, item_type = lookup_name_to_use(item['code'])
        if item_type in item_types:
            item_types[item_type]['codes'].append(code)
            item_types[item_type]['descriptions'].append(description)
            item_types[item_type]['quantities'].append(item['quantity'])
            item_types[item_type]['costs'].append(item['cost'])
            item_types[item_type]['type'].append('Item')
    return item_types


def add_header_to_item_types(item_types):
    for item_type in ['Tea', 'Gus', 'Gorgie', 'Chip']:
        if len(item_types[item_type]['codes']) > 0:
            if item_type == 'Tea':
                item_types[item_type]['codes'].insert(0, 'T')
                item_types[item_type]['descriptions'].insert(0, '*******************20 OZ TEA******************')
            elif item_type == 'Gus':
                item_types[item_type]['codes'].insert(0, 'G')
                item_types[item_type]['descriptions'].insert(0, '*******************GUS******************')
            elif item_type == 'Gorgie':
                item_types[item_type]['codes'].insert(0, 'R')
                item_types[item_type]['descriptions'].insert(0, '*******GORGIE - NY AND CT BOTTLE DEPOSIT- .6 PER CASE******************')
            elif item_type == 'Chip':
                item_types[item_type]['codes'].insert(0, 2)
                item_types[item_type]['descriptions'].insert(0, '*******************2 OZ CHIPS******************')
            item_types[item_type]['quantities'].insert(0, 0)
            item_types[item_type]['costs'].insert(0, 0)
            item_types[item_type]['type'].insert(0, '')
    return item_types


def get_final_lists(item_types, order_data):
    codes = []
    descriptions = []
    quantities = []
    costs = []
    item_or_not = []

    for item_type in item_types:
        codes += item_types[item_type]['codes']
        descriptions += item_types[item_type]['descriptions']
        quantities += item_types[item_type]['quantities']
        costs += item_types[item_type]['costs']
        item_or_not += item_types[item_type]['type']

    data = {
        'Document Type': ['Order'] * (6 + len(codes)),
        'Document No.': [f'S{order_data["Order Number"]}'] * (6 + len(codes)),
        'Line No.': [10000 * (i + 1) for i in range(6 + len(codes))],
        'Type': [''] * 6 + item_or_not,
        'No.': [''] * 6 + codes,
        'Location Code': ['ELIZABETH'] * (6 + len(codes)),
        'Shipment Date': [order_data["Expected Delivery Date"]] * (6 + len(codes)),
        'Description': ['Expected Delivery Date:', order_data["Expected Delivery Date"],
                        f'Store No: {order_data["Store No"]}', f'Account No: {order_data["Customer ID"]}',
                        f'Subteam: {order_data["Subteam"]}', f'Buyer: {order_data["Buyer"]}'] + descriptions,
        'Quantity': [0] * 6 + quantities,
        'Unit Price': [0] * 6 + costs,
        'Customer Number': order_data["Customer ID"],
        'Order Date': order_data["Order Date"],
        'Order Number': order_data["Order Number"]
    }
    return data


def process_html_file(html_file_path):
    with open(html_file_path, "r") as file:
        html = file.read()
        soup = BeautifulSoup(html, "html.parser")
        # print(soup.prettify())
        order_data = extract_order_data(soup)
        data = extract_item_data(soup)
        item_types = process_data(data)
        item_types = add_header_to_item_types(item_types)
        data = get_final_lists(item_types, order_data)
    return data


def add_lists_together(data_list, wholefoodsinfo):
    df = wholefoodsinfo
    sales_header = {
        'Document Type': [],
        'No.': [],
        'Sell-to Customer No.': [],
        'Bill-to Customer No.': [],
        'Ship-to Code': [],
        'Order Date': [],
        'Due Date': [],
        'Location Code': [],
        'Customer Posting Group': [],
        'Customer Price Group': [],
        'Gen. Bus. Posting Group': [],
        'Sell-to Customer Name': [],
        'Sell-to Customer Name 2': [],
        'Sell-to Address': [],
        'Sell-to Address 2': [],
        'Sell-to City': [],
        'Sell-to ZIP code': [],
        'External Document No.': [],
        'Delivery Info': [],
        'Delivery Exception': [],
        'Route': [],
        'Store Delivery Location': []
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
        customer_number = (data['Customer Number'])
        sales_header['Document Type'] += ['Order']
        sales_header['No.'] += [data['Document No.'][0]]
        sales_header['Sell-to Customer No.'] += [customer_number]
        sales_header['Bill-to Customer No.'] += ['WHOLE FOODS CORPORAT']
        sales_header['Ship-to Code'] += [df.loc[df['No.'] == customer_number, "Ship-to Code"].values[0]]
        sales_header['Ship-to Name'] += [
            df.loc[df['No.'] == customer_number, "Sell-to Customer Name"].values[0]]
        sales_header['Ship-to Address'] += [df.loc[df['No.'] == customer_number, 'Sell-to Address'].values[0]]
        sales_header['Ship-to Address'] += [df.loc[df['No.'] == customer_number, 'Sell-to Address 2'].values[0]]
        sales_header['Ship-to City'] += [df.loc[df['No.'] == customer_number, 'Sell-to City'].values[0]]
        sales_header['Order Date'] += [data['Order Date']]
        sales_header['Due Date'] += ['']
        sales_header['Location Code'] += ['ELIZABETH']
        sales_header['Customer Posting Group'] += ['WHOLE FOODS']
        sales_header['Customer Price Group'] += ['WHOLE FOOD']
        sales_header['Gen. Bus. Posting Group'] += ['ALL']
        sales_header['Sell-to Customer Name'] += [df.loc[df['No.'] == customer_number, "Sell-to Customer Name"].values[0]]
        sales_header['Sell-to Customer Name 2'] += ['']
        sales_header['Sell-to Address'] += [df.loc[df['No.'] == customer_number, 'Sell-to Address'].values[0]]
        sales_header['Sell-to Address 2'] += [df.loc[df['No.'] == customer_number, 'Sell-to Address 2'].values[0]]
        sales_header['Sell-to City'] += [df.loc[df['No.'] == customer_number, 'Sell-to City'].values[0]]
        sales_header['Sell-to ZIP code'] += [df.loc[df['No.'] == customer_number, 'Sell-to ZIP Code'].values[0]]
        sales_header['External Document No.'] += [data['Order Number']]
        sales_header['Delivery Info'] += [df.loc[df['No.'] == customer_number, 'Delivery Info'].values[0]]
        sales_header['Delivery Exception'] += [df.loc[df['No.'] == customer_number, 'Delivery Exception'].values[0]]
        sales_header['Route'] += [df.loc[df['No.'] == customer_number, 'Route'].values[0]]
        sales_header['Store Delivery Location'] += [df.loc[df['No.'] == customer_number, 'Store Delivery Location'].values[0]]

    return sales_line, sales_header


def make_the_excel(sales_line, sales_header):
    # Create two DataFrames
    sales_line_df = pd.DataFrame(sales_line)
    sales_header_df = pd.DataFrame(sales_header)

    # Write each DataFrame to its own sheet in the same Excel file
    with pd.ExcelWriter('orders.xlsx', engine='openpyxl', mode='w') as writer:
        workbook = writer.book

        # Create another sheet in the workbook
        sales_header_sheet = workbook.create_sheet("Sales Header")

        # Write headers
        sales_header_sheet['A1'] = "WHOLE FOODS IMPORT"
        sales_header_sheet['B1'] = "Sales Header"
        sales_header_sheet['C1'] = "36"

        # Write the data in sales_header_df starting from cell A3
        start_row = 3
        sales_header_df.to_excel(writer, sheet_name="Sales Header", startrow=start_row, index=False)

        # Create a new sheet in the workbook
        sales_line_sheet = workbook.create_sheet("Sales Line")

        # Write the values in cells A1:C1
        sales_line_sheet['A1'] = "WHOLE FOODS IMPORT"
        sales_line_sheet['B1'] = "Sales Line"
        sales_line_sheet['C1'] = "37"

        # Write the data in sales_line_df starting from cell A3
        start_row = 3
        sales_line_df.to_excel(writer, sheet_name="Sales Line", startrow=start_row, index=False)

