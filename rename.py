import os

baseDir = "./com"
files = os.listdir(baseDir)
for file in files:
    os.rename(f"{baseDir}/{file}", f"{baseDir}/{file}.xlsx")
print("success")