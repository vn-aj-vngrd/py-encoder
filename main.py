# https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html

import pandas as pd
from openpyxl import Workbook
from datetime import datetime
import os
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

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
    "Details",
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

while True:
    i = 0
    files = []

    for excel in os.listdir("./data"):
        if excel.endswith(".xlsx"):
            files.append(excel)
            print(i, "-", excel)
            i += 1

    # Get the location of the data
    try:
        file_key = input("\nInput file number: ")
        path = "data/" + files[int(file_key)]
    except Exception as e:
        print("Error: ", str(e))

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

                book.save("main_res/" + key + ".xlsx")

        print("Done...")
    except Exception as e:
        print("Error: " + str(e))

    isContinue = input("Input 1 to continue: ")
    if isContinue != "1":
        break
