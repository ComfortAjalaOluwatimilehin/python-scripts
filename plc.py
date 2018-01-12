#USAGE python plc.py y <filepath> 
#plc - > plot line chart 

import sys, openpyxl,os

def plot_line_chart(filepath):
	wb = openpyxl.load_workbook(filepath)
	ws = wb.active 
	lineChart = openpyxl.chart.LineChart()
	lineChart.style = 13
	#if labels start from the second row 
	#else start from the first 
	row_start = 1
	
	if type(ws.cell(row=1, column=1).value) == str:
		lineChart.x_axis.title = ws.cell(row=1, column=1).value 
		lineChart.y_axis.title =ws.cell(row=1, column=2).value 
		row_start = 2
	data = openpyxl.chart.Reference(ws, min_col=2, min_row=row_start, max_col=2, max_row=ws.max_row)
	cat = openpyxl.chart.Reference(ws, min_col=1, min_row=row_start, max_row=ws.max_row) #categories
	lineChart.add_data(data,  titles_from_data=False)
		# Style the lines
	s1 = lineChart.series[0] 
	s1.title = None
	lineChart.y_axis.majorGridlines = None
	lineChart.x_axis.tickLblSkip = 100
	s1.graphicalProperties.line.solidFill = "848484"
	s1.smooth = True # Make the line smooth
	
	lineChart.set_categories(cat)
	
	ws.add_chart(lineChart, "D10")
	wb.save(filepath)
	
if __name__ == "__main__":
	if len( sys.argv ) == 2:
		path = sys.argv[1]
		filepath = os.path.abspath(path)
		if os.path.exists(filepath):
			plot_line_chart(filepath)
		else:
			print("Requires existing file")
	else:
		print("Requires file path")