import pandas as pd
import numpy as np

# Get the location of the data
path = "./test.xlsx"

# Read the data
df_dict = pd.read_excel(path, sheet_name=None)

# Combine data from all worksheets as single DataFrame
df_all = pd.concat(df_dict.values(), ignore_index=True)

# Get sheet names
df_sheetNames = pd.ExcelFile(path)

# Locate the data
machinery = df_all.iloc[0, 2]

# for i in range(len(df_sheetNames.sheet_names)):
#     print(i)


data = pd.read_excel(path, index_col=None, header=None, sheet_name=1, nrows=2)
# data.fillna(" ", inplace=True)

print(data)

# print("Machinery: " + machinery)
# print("\n")
# print(df_all)
