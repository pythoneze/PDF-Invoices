import pandas as pd
import glob as gb

FILEPATHS = gb.glob("invoices/*xlsx*")

for filepath in FILEPATHS:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")