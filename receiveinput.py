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

