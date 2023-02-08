from bs4 import BeautifulSoup
import pandas as pd

# Read the Excel file into a DataFrame
df = pd.read_excel("Whole Foods Info.xlsx")


# Look up a value in the "Whole Foods Name" column and return the corresponding "Name to Use" value
def lookup_name_to_use(value):
    return df.loc[df["Whole Foods Name"] == value.upper()]["Name to Use"].values[0]


with open("order_122886246.html", "r") as file:
    html = file.read()
    soup = BeautifulSoup(html, "html.parser")
# print(soup.prettify())

# Extract Order Number
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
buyer = buyer_tds[0].text if buyer_tds else None

# Print extracted data
print("Order Number:", order_number)
print("Order Date:", order_date)
print("Expected Delivery Date:", delivery_date)
print("Store No:", store_no)
print("Customer ID:", customer_ID)
print("Subteam:", subteam)
print("Buyer:", buyer)

rows = soup.find_all('tr')
data = []
for row in rows:
    cells = row.find_all('td')
    if len(cells) == 7:
        item = {
            'code': cells[1].text.strip(),
            'quantity': int(cells[2].text.strip().split("\xa0")[0]),
            'cost': cells[5].text.strip()
            # 'number': cells[0].text.strip(),
            # 'description': cells[3].text.strip(),
            # 'unit': cells[4].text.strip(),
            # 'UPC': cells[6].text.strip()
        }
        data.append(item)
item_dictionary = {}
item_cost = {}
for item in data:
    code = lookup_name_to_use(item['code'])
    quantity = item['quantity']
    item_dictionary[code] = quantity
    item_cost[code] = item['cost']
print(item_dictionary.keys())
print(item_cost)

