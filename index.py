import pandas as pd
import numpy as np

# Get the location of the data
path = "./test.xlsx"

# Read the data
df_dict = pd.read_excel(path, sheet_name=None)

# Get sheet names
df_sheetNames = pd.ExcelFile(path)

# Locate the data
# newData = data.iloc[0,2]

