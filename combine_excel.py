import pandas as pd
import glob

location = "./data/重客群/*.xlsx"
excel_files = glob.glob(location)
print(excel_files)
pd.set_option("display.max_rows", 9999)
df = pd.DataFrame()
for excel_file in excel_files:
    tmp_df = pd.read_excel(excel_file)
    df = pd.concat([df, tmp_df])
df.to_excel("重客群.xlsx", index=False)