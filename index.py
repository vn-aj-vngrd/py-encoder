# https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html

from logging import exception
import pandas as pd
from openpyxl import Workbook
from datetime import datetime

notIncluded = [
    "Main Menu",
    "Running Hours",
    "MECO Setting",
    "Sheet3",
    "Cylinder Liner Monitoring",
    "ME Exhaust Valve Monitoring",
    "FIVA VALVE Monitoring",
    "Fuel Valve Monitoring",
    "Sheet1",
]

header = (
    "vessel",
    "machinery",
    "code",
    "name",
    "description",
    "interval",
    "commissioning_date",
    "last_done_date",
    "last_done_running_hours",
)

# Get the location of the data
filename = input("Input filename: ")
path = "data/" + filename + ".xlsx"

# Read the data
try:
    data = pd.read_excel(path, sheet_name=None, index_col=None, header=None)

    # Get the keys
    xl = pd.ExcelFile(path)
    keys = xl.sheet_names

    # Iterate through the sheets
    for key in keys:
        if key not in notIncluded:
            print(key)

            # Vessel Name
            vessel = data[key].iloc[0, 2]

            # Machinery Name
            machinery = data[key].iloc[2, 2]

            # Start traversing the data on row 7
            row = 7
            isValid = True

            # Prepare the sheets
            book = Workbook()
            sheet = book.active

            sheet.append(header)

            while isValid:

                rowData = (
                    vessel,
                    machinery,
                )

                for col in range(7):
                    d = data[key].iloc[row, col]

                    if (pd.isna(d)) and (col == 0):
                        isValid = False
                        break

                    if ((col == 4) or (col == 5)) and isinstance(d, datetime):
                        d = d.strftime("%d-%b-%y")

                    tempTuple = (d,)
                    rowData += tempTuple

                if isValid:
                    sheet.append(rowData)
                    row += 1

            book.save("res/" + key + ".xlsx")

    print("Done...")
except Exception as e:
    print("Error: " + str(e))
