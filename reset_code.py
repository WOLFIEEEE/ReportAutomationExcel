import openpyxl

# load the workbook
workbook = openpyxl.load_workbook('original.xlsx')

# select the sheet to delete
sheet_to_delete = workbook['Trying']

# delete the sheet
workbook.remove(sheet_to_delete)

# save the changes to the workbook
workbook.save('original.xlsx')
