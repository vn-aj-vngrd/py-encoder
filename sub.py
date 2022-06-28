# https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html

import pandas as pd
from openpyxl import Workbook
from datetime import datetime
import os
import warnings
from datetime import date

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
    "running_hours",
    "updating_date",
)

if not os.path.exists("./sub_res"):
    os.makedirs("./sub_res")

while True:
    if not os.path.exists("./src"):
        os.makedirs("./src")
        
    files = []
    i = 0
    for excel in os.listdir("./src"):
        if excel.endswith(".xlsx"):
            files.append(excel)
            print(i, "-", excel)
            i += 1

    if (len(files) == 0):
        print("No such data found in src directory.")
        exit()
        
    # Get the location of the data
    try:
        file_key = input("\nInput file number: ")
        file_name = files[int(file_key)]

        path = "src/" + file_name
    except Exception as e:
        print("Error: ", str(e))

    # Read the data
    try:
        data = pd.read_excel(path, sheet_name=None, index_col=None, header=None)

        # Get the keys
        xl = pd.ExcelFile(path)
        keys = xl.sheet_names

        # Get updated_at
        updating_date = data["Running Hours"].iloc[2, 3].strftime("%d-%b-%y")
        # print(updating_date)

        # Prepare the sheets
        book = Workbook()
        sheet = book.active

        # Append the dates
        sheet.append(header)

        # Iterate through the sheets
        for key in keys:
            if key not in notIncluded:
                print(key)

                # Vessel Name
                vessel = data[key].iloc[0, 2]

                # Machinery Name
                machinery = data[key].iloc[2, 2]

                # Running Hours
                running_hours = data[key].iloc[3, 5]

                rowData = (vessel, machinery, running_hours, updating_date)
                sheet.append(rowData)

        create_name = file_name[: len(file_name) - 4]
        creation_folder = "./sub_res/" + create_name
        if not os.path.exists(creation_folder):
            os.makedirs(creation_folder)
        book.save(creation_folder + "/" + file_name)

        print("Done...")
    except Exception as e:
        print("Error: " + str(e))

    isContinue = input("Input 1 to continue: ")
    if isContinue != "1":
        break
