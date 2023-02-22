from openpyxl.chart import BarChart, Reference
from openpyxl.utils.cell import range_boundaries
from openpyxl.chart.label import DataLabelList
from openpyxl.chart import PieChart, Reference, Series
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.colors import ColorChoice
# from openpyxl.chart.label import DataLabel, LeaderLines

# leader_lines = LeaderLines()

# leader_lines.chartSpace = 5
# leader_lines.spPr = None

def create_column_chart_issuetype(workbook, sheetname, tablename):
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
    pos = f"K{b-3}"
    # Add the chart to the worksheet
    ws.add_chart(chart, pos)


# def create_column_chart_severity(workbook, sheetname, tablename, last_row):
#     # Select the sheet by name
#     ws = workbook[sheetname]
#     create_column_chart(ws)

#     # Get the table by name
#     table = ws.tables[tablename]

#     # Get the range of cells that the table covers
#     table_range = table.ref

#     # Get the column of data to plot in the chart (assuming it's the second column)
#     col_index = 2

#     # Define the chart object and set its properties
    

#     # Define the data series for the chart
#     table_start_row, table_start_col, table_end_row, table_end_col = range_boundaries(table_range)
#     print(table_start_row, table_start_col, table_end_row, table_end_col)
#     # length = table_end_col - table_start_col-2;
#     height = 9;
#     chart = BarChart()
#     chart.title = "Severity/Impact Wise Defect Distribution"
#     chart.x_axis.title = "Defect Impact Category"
#     chart.y_axis.title = "Defect Count"
#     chart.width = 15
#     chart.height = height
#     chart.dataLabels = DataLabelList()
#     chart.dataLabels.showVal = True
#     chart.dataLabels.showCatName = False
#     chart.dataLabels.showSerName = False
#     chart.dataLabels.showLegendKey = False
#     # chart.dataLabels.number_format = "0"
#     # chart.dataLabels.position = "t"
#     print(table_start_row, table_start_col, table_end_row, table_end_col+1)
#     data = Reference(ws , min_row = last_row, min_col = table_start_row , max_col= table_end_row+1 , max_row=last_row)
#     # data = Reference(ws, min_col=col_index, min_row=table_start_col, max_row=table_end_col)
#     chart.add_data(data , titles_from_data=True)

#     # Define the category axis for the chart
#     categories = Reference(ws , min_row = last_row-1, min_col = table_start_row , max_col= table_end_row+1 )
#     # categories = Reference(ws, min_col=col_index-1, min_row=table_start_col+1, max_row=table_end_col)
#     chart.set_categories(categories )

#     # create a cloustered column chart using the data and categories
    
#     # for series in chart.series:
#     #     if series.dLbls is not None:
#     #         for dLbl in series.dLbls:
#     #             dLbl.tx.rich.p[0].numFmtId = -1

#     nn = int(((table_end_col - table_start_col)/4) + 3);
#     # Add the chart to the worksheet
#     pos = f"F{nn}"
#     ws.add_chart(chart, "H30")