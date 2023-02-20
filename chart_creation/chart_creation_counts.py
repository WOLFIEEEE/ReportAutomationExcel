from openpyxl.worksheet.table import Table
from openpyxl.chart import BarChart, Reference, Series

def create_column_chart_for_counts(sheet_name, chart_title, table_name, workbook):
    # get the worksheet by name
    sheet = workbook[sheet_name]

    # get the table by name
    table = sheet.tables[table_name]

    # get the range of the data in the table
    data_range = table.ref

    # create a new chart
    chart = BarChart()

    # set the title of the chart
    chart.title = chart_title

    # define the data range for the chart
    data = Reference(sheet, range_string=data_range)

    # define the x-axis categories for the chart
    categories = Reference(sheet, min_col=table.header_row_count+1, min_row=table.header_row_count, max_col=table.total_columns, max_row=table.header_row_count)

    # add the data and categories to the chart
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    # set the style of the chart
    chart.style = 10

    # add the chart to the worksheet
    sheet.add_chart(chart, "D1")

    return workbook
