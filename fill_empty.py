import pandas as pd

df = pd.read_excel("output.xlsx")
df.fillna(value="否",inplace=True)
df.to_excel("output_format.xlsx", index=False, sheet_name="Sheet1")