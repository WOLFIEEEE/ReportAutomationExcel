import openpyxl
from data_creation.df_creation_counts import create_counts_df
import pandas as pd
from table_creation.table_creation_counts import add_dataframe_to_excel_sheet
from table_creation.table_creation_severity import table_creation_for_severity
from table_creation.table_creation_conformance import table_creation_for_conformance
from table_creation.table_creation_issuetype import table_creation_for_issuetype
from table_creation.table_creation_status import table_creation_for_status
from table_creation.table_creation_status_counts import table_creation_for_status_counts
workbook = openpyxl.load_workbook('original.xlsx', data_only=True)

# remove_extra_spaces(filename="original.xlsx")


#code for Adding the Counts Table and Chart to the Page
df , new_df = create_counts_df(workbook)
last_row = add_dataframe_to_excel_sheet(workbook , "Trying" ,  new_df)
# code Ended to add the Counts Table and Chart to the Page

table_creation_for_severity(workbook, "Trying", last_row)
table_creation_for_conformance(workbook, "Trying", last_row + 15)
table_creation_for_issuetype(workbook, "Trying", last_row + 30)
table_name = table_creation_for_status(workbook, "Trying", df , last_row + 52)
table_creation_for_status_counts(workbook, "Trying", table_name, last_row + 45)




print(last_row)
workbook.save(filename="original.xlsx")