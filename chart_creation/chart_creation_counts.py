from openpyxl.chart import BarChart, Reference
from openpyxl.utils.cell import range_boundaries
from openpyxl.chart import Series

def create_column_chart(workbook, sheetname, tablename):
    # Select the sheet by name
    ws = workbook[sheetname]

    # Get the table by name
    table = ws.tables[tablename]

    # Get the range of cells that the table covers
    table_range = table.ref

    # Get the column of data to plot in the chart (assuming it's the second column)
    col_index = 2

    # Define the chart object and set its properties
    

    # Define the data series for the chart
    table_start_row, table_start_col, table_end_row, table_end_col = range_boundaries(table_range)
    print(table_start_row, table_start_col, table_end_row, table_end_col)
    length = table_end_col - table_start_col;
    height = 10;
    chart = BarChart()
    chart.title = "WCAG 2.1 AA Success Criteria Distribution"
    chart.x_axis.title = "WCAG Success Criteria #"
    chart.y_axis.title = "Defect Count"
    chart.width = length
    chart.height = height
    
    data = Reference(ws, min_col=col_index, min_row=table_start_col, max_row=table_end_col)
    chart.add_data(data , titles_from_data=True)

    # Define the category axis for the chart
    categories = Reference(ws, min_col=col_index-1, min_row=table_start_col+1, max_row=table_end_col)
    chart.set_categories(categories )

    # Add the chart to the worksheet
    ws.add_chart(chart, "E10")

    # Save the workbook
