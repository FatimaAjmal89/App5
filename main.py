import pandas as pd
import glob

#pandas need openpyxl to read excel files
#xlsx is a used for reading excel sheets

filepaths = glob.glob("invoices/*.xlsx")
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)