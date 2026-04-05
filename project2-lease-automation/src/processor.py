import numpy as np
import pandas as pd

prev_qtr = "Q3"
curr_qtr = "Q4"
fiscal_year = "FY26"


# loading GL activity file
prev_df = pd.read_excel("../data/lease_harbor_Q3.xlsx", header=0, sheet_name=0)
prev_df["CoCodeCurr"] = prev_df["Company_Code"] + prev_df["Currency"] 

# filtered_df = df[df['Age'] == 25]
prev_df_A = prev_df[prev_df["Portfolio"]=="Region A"]
prev_df_B = prev_df[prev_df["Portfolio"]=="Region B"]
prev_df_C = prev_df[prev_df["Portfolio"]=="Region C"]







# loading GL activity file
curr_df = pd.read_excel("../data/lease_harbor_Q4.xlsx", header=0, sheet_name=0)
curr_df["CoCodeCurr"] = curr_df["Company_Code"] + curr_df["Currency"]

# Filtered by region
curr_df_A = curr_df[curr_df["Portfolio"]=="Region A"]
curr_df_B = curr_df[curr_df["Portfolio"]=="Region B"]
curr_df_C = curr_df[curr_df["Portfolio"]=="Region C"]




# save to Excel file
with pd.ExcelWriter("../output/Processed JE Summary.xlsx") as writer:
    #pivot.to_excel(writer, sheet_name="Summary Pivot",      index=False)
    #merged.to_excel(writer, sheet_name="SAP GL Data",       index=False)
    prev_df.to_excel(writer, sheet_name="Chart of Accounts", index=False)
    