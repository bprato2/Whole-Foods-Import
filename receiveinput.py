from bs4 import BeautifulSoup

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
account_no = account_no_tds[0].text if account_no_tds else None

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
print("Account No:", account_no)
print("Subteam:", subteam)
print("Buyer:", buyer)