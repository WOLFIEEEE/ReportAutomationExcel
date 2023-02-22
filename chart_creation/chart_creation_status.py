from openpyxl.chart import BarChart, Reference
from openpyxl.utils.cell import range_boundaries
from openpyxl.chart.label import DataLabelList
from openpyxl.chart import PieChart, Reference, Series
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference, Series

def create_pie_chart_status(workbook, sheetname, tablename):
    ws = workbook[sheetname]
    table = ws.tables[tablename]

    # Get the range of cells that the table covers
    table_range = table.ref
    # Define the data series for the chart
    #keep in mind here rows and column are mismtached
    a,b,c,d = range_boundaries(table_range)
    print(a,b,c,d)
    # Define the data range for the chart
    data_start_row = d
    data_end_row = d
    table_end_col = c
    table_start_col = a
    data = Reference(ws, min_row=data_start_row, min_col=table_start_col, max_col=table_end_col, max_row=data_end_row)

    # Define the category range for the chart
    category_start_row = b
    category_end_row = b
    table_start_col = a+1
    table_end_col = c
    categories = Reference(ws, min_row=category_start_row, min_col=table_start_col, max_col=table_end_col, max_row=category_end_row)

    # Create the chart and add the data series to it
    chart = PieChart()
    chart.title = "Severity/Impact Wise Defect Distribution"
    chart.width = 16
    chart.height = 10
    chart.legend.position = 'r'
    #chart.legend.layout = 'vertical'
    # chart.legend.font = Font(size=12, bold=True)
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showLeaderLines = True
    chart.dataLabels.showPercent = True
    chart.dataLabels.showCatName = True
    chart.dataLabels.showSerName = False
    chart.dataLabels.showVal = False
    chart.dataLabels.dLblPos = 'outEnd'
    chart.dataLabels.dLblsOverlapping = 'noOverlap'
    chart.dataLabels.showLegendKey = False
    data_series = Series(data, title_from_data=True)
    chart.append(data_series)

    # Set the categories for the chart
    chart.set_categories(categories)
    pos = f"H{b+3}"
    # Add the chart to the worksheet
    ws.add_chart(chart, pos)