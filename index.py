import pandas as pd
import numpy as np

path = "./test.xlsx"

df_dict = pd.read_excel(path, sheet_name=None)
df_sheetNames = pd.ExcelFile(path)

# newData = data.iloc[0,2]

# print(newData)

print(df_dict.sheet_names)
