import openpyxl
from data_creation.df_creation_counts import create_counts_df
from table_creation.table_creation_counts import add_dataframe_to_excel_sheet
# from remove_spaces import remove_extra_spaces
workbook = openpyxl.load_workbook('original.xlsx', data_only=True)

# remove_extra_spaces(filename="original.xlsx")
new_df = create_counts_df(workbook)
add_dataframe_to_excel_sheet(workbook , "Trying" ,  new_df)

workbook.save(filename="original.xlsx")