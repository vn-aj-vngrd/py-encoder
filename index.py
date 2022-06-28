import pandas as pd
import numpy as np
import math

# https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html

# Get the location of the data
path = "./test.xlsx"

vessel = "Vessel_1"

# header = ["vessel", "machinery", "name", "description", "interval", "comissioning_date", "last_done", "running_hours"]

# Read the data
data = pd.read_excel(path, sheet_name=None, index_col=None, header=None)

# Get the keys
xl = pd.ExcelFile(path)
keys = xl.sheet_names
del keys[0]
del keys[0]

# Iterate through the sheets
for key in keys:
    # print(key)

    writer = pd.ExcelWriter(key + ".xlsx", engine="xlsxwriter")
    writer.save()

    # Machinery Name
    print(data[key].iloc[2, 2])
    row = 7

    isValid = True
    while isValid:
        for col in range(7):
            d = data[key].iloc[row, col]

            if (pd.isna(d)) and (col == 0):
                isValid = False
                break

            print(d)
        row += 1
    break
