import openpyxl
import pyxll


# Open the workbook you want to add the XLM macro to
wb = openpyxl.load_workbook("orders.xlsx")

# Import the XML map
xml_map_path = "header xml file.xml"
wb.vba_project.import_file(xml_map_path)

# Save the changes to the workbook
wb.save("orders.xlsx")
