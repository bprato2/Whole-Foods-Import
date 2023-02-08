from bs4 import BeautifulSoup
import pandas as pd

# Read the Excel file into a DataFrame
df = pd.read_excel("Whole Foods Info.xlsx")


# Function to lookup the "Name to Use", "Description", and "Type" based on the "Whole Foods Name"
def lookup_name_to_use(value):
    item_code = df.loc[df["Whole Foods Name"] == value.upper()]["Name to Use"].values[0]
    item_description = df.loc[df["Whole Foods Name"] == value.upper()]["Description"].values[0]
    item_type = df.loc[df["Whole Foods Name"] == value.upper()]["Type"].values[0]
    return item_code, item_description, item_type


with open("order_122886246.html", "r") as file:
    html = file.read()
    soup = BeautifulSoup(html, "html.parser")
# print(soup.prettify())

'''# Extract Order Number
order_number_tds = soup.select("td:-soup-contains('Order Number:') + td")
order_number = order_number_tds[0].text if order_number_tds else None

# Extract Order Date
order_date_tds = soup.select("td:-soup-contains('Order Date:') + td")
order_date = order_date_tds[0].text if order_date_tds else None

# Extract Expected Delivery Date
delivery_date_tds = soup.select("td:-soup-contains('Expected Delivery Date:') + td")
delivery_date = delivery_date_tds[0].text if delivery_date_tds else None

# Extract Store No
store_no_tds = soup.select("td:-soup-contains('Store No:') + td")
store_no = store_no_tds[0].text if store_no_tds else None

# Extract Account No
account_no_tds = soup.select("td:-soup-contains('Account No:') + td")
customer_ID = account_no_tds[0].text if account_no_tds else None

# Extract Subteam
subteam_tds = soup.select("td:-soup-contains('Subteam:') + td")
subteam = subteam_tds[0].text if subteam_tds else None

# Extract Buyer
buyer_tds = soup.select("td:-soup-contains('Buyer:') + td")
buyer = buyer_tds[0].text if buyer_tds else None'''

def extract_order_data(soup):
    order_data = {}
    order_data["Order Number"] = soup.select_one("td:-soup-contains('Order Number:') + td").text
    order_data["Order Date"] = soup.select_one("td:-soup-contains('Order Date:') + td").text
    order_data["Expected Delivery Date"] = soup.select_one("td:-soup-contains('Expected Delivery Date:') + td").text
    order_data["Store No"] = soup.select_one("td:-soup-contains('Store No:') + td").text
    order_data["Customer ID"] = soup.select_one("td:-soup-contains('Account No:') + td").text
    order_data["Subteam"] = soup.select_one("td:-soup-contains('Subteam:') + td").text
    order_data["Buyer"] = soup.select_one("td:-soup-contains('Buyer:') + td").text
    return order_data


order_data = extract_order_data(soup)

# Print extracted data
'''print("Order Number:", order_number)
print("Order Date:", order_date)
print("Expected Delivery Date:", delivery_date)
print("Store No:", store_no)
print("Customer ID:", customer_ID)
print("Subteam:", subteam)
print("Buyer:", buyer)'''

rows = soup.find_all('tr')
data = []
for row in rows:
    cells = row.find_all('td')
    if len(cells) == 7:
        item = {
            'code': cells[1].text.strip(),
            'quantity': int(cells[2].text.strip().split("\xa0")[0]),
            'cost': float(cells[5].text.strip())
            # 'number': cells[0].text.strip(),
            # 'description': cells[3].text.strip(),
            # 'unit': cells[4].text.strip(),
            # 'UPC': cells[6].text.strip()
        }
        data.append(item)

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
    'Unit Price': [0] * 6 + costs
}

# Create a DataFrame
df = pd.DataFrame(data)

# Save the DataFrame as an Excel sheet
df.to_excel('order.xlsx', index=False)


