import numpy as np
import pandas as pd

# loading GL activity file
sap_df = pd.read_excel("../data/sap_export.xlsx", header=0, sheet_name=0)


# loading chart of account file
coa_df = pd.read_excel("../data/SAP_Chart_of_Accounts.xlsx", header=0, sheet_name=0)
coa_df[['Numbering','Category']] = coa_df['Hierarchy'].str.split(' - ', expand=True)
coa_df = coa_df.filter(items=["Account Number","Category","Description"])
coa_df = coa_df.dropna()
coa_df = coa_df.rename(columns={"Account Number": "GL_Account"})

# Vlookup with GL_Account with chart of account for category
merged = pd.merge(sap_df,coa_df[["GL_Account","Category"]],on='GL_Account',how='left')

# make pivot grouped by category and sum of Amount
pivot = merged.groupby('Category')['Amount'].sum().reset_index()

# adding total row
total_row = pivot.sum(numeric_only=True).to_frame().T
total_row.index = ['Total']
total_row['Category'] = 'Total' 
pivot = pd.concat([pivot, total_row])


# save to Excel file
with pd.ExcelWriter("../output/Processed JE Summary.xlsx") as writer:
    pivot.to_excel(writer, sheet_name="Summary Pivot",      index=False)
    merged.to_excel(writer, sheet_name="SAP GL Data",       index=False)
    coa_df.to_excel(writer, sheet_name="Chart of Accounts", index=False)
    