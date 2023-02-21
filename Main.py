import openpyxl
from data_creation.df_creation_counts import create_counts_df
import pandas as pd
from table_creation.table_creation_counts import add_dataframe_to_excel_sheet
# from remove_spaces import remove_extra_spaces
workbook = openpyxl.load_workbook('original.xlsx', data_only=True)

# remove_extra_spaces(filename="original.xlsx")


#code for Adding the Counts Table and Chart to the Page
new_df = create_counts_df(workbook)
last_row = add_dataframe_to_excel_sheet(workbook , "Trying" ,  new_df)
# code Ended to add the Counts Table and Chart to the Page


print(last_row)
workbook.save(filename="original.xlsx")