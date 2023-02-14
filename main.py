import os
from receiveinput import process_html_file, add_lists_together, make_the_excel
import pandas as pd

folder_path = "C:/Users/blair/OneDrive/Documents/GitHub/Whole-Foods-Import/Whole Foods PDFs"

orders_compiled = []
for filename in os.listdir(folder_path):
    file_path = os.path.join(folder_path, filename)
    if os.path.isfile(file_path):
        sales_line = process_html_file(file_path)
        orders_compiled.append(sales_line)

# Read the Excel file into a DataFrame
df = pd.read_excel("Whole Foods Stores.xlsx")

sales_line_data, sales_header_data = add_lists_together(orders_compiled, df)

make_the_excel(sales_line_data, sales_header_data)

print("Your file has been saved!")
