import pandas as pd

# TODO
writer = pd.ExcelWriter("xxx")
df = pd.read_excel("xx")
df.fillna(value="N/A",inplace=True)
df.to_excel("", index=False, sheet_name="xxx")
writer.save()