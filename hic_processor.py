#script to process my hydrophobic interaction chromatography data
#USAGE: python hic_processor.py <filepath>

import os, sys, logging, openpyxl, re

logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")

filepath  = None

def text_to_rows(textfilepath: str):
	fh = open(textfilepath)
	data = []
	for line in fh:
		line = line.strip()
		values = re.split(r"\s+", line)
		data.append(values)
	fh.close()
	return german_to_us(data)

def german_to_us(rows: list):
	pattern = re.compile(r"(\d+),(\d+)")
	#has_num = re.compile(r"\d+")
	update = []
	for row in rows:
		sub_update = []
		for value in row:
			if pattern.search(value):
				value = re.sub(pattern, r"\1.\2", value)
				value = float(value)
			sub_update.append(value)
		update.append(sub_update)
		
	return rows_to_sheet(update)
def rows_to_sheet(rows: list):
	target_path = os.path.dirname(filepath)
	filename = filepath.split(".")[0]
	excel_path = os.path.join( target_path, filename + ".xlsx"  )
	wb = openpyxl.Workbook()
	ws = wb.active

	for row in range( len(  rows  )  ):
		for col in range(  len(  rows[row]  ) ):
				ws.cell(row=row + 1, column=col + 1, value=rows[row][col])
				
	wb.save(excel_path)
	logging.info("Excel Sheet saved")
	return draw_chart(excel_path)

#extra 

def draw_chart(filepath: str):
	#personalized 
	#row 4 col 2 --> categories
	#row r4c1 (280), r4c4(220), r4c6 (260)
	
	wb = openpyxl.load_workbook(filepath)
	ws = wb.active 
	lineChart = openpyxl.chart.LineChart()
	lineChart.style = 13
	lineChart.y_axis.title = 'Absorption (mAU)'
	lineChart.x_axis.title = 'Volume (ml)'
	lineChart.y_axis.scaling.min = 0
	lineChart.y_axis.scaling.max = 6000
	
	lineChart.x_axis.tickLblSkip = 2050
	lineChart.y_axis.tickLblSkip = 1500
	
	lineChart.x_axis.minorTickMark = None
	
	#lineChart.x_axis.scaling.max = 120
	titles = [{"c":1, "title":"280nm"}, {"c":6, "title":"260nm"}]
	for nm in range( len(titles) ):
		r = openpyxl.chart.Reference(ws, min_col=titles[nm]["c"], min_row=4, max_row=ws.max_row)
		s = openpyxl.chart.Series(r, title=titles[nm]["title"])
		lineChart.series.append(s)
	cat = openpyxl.chart.Reference(ws, min_col=2, min_row=4, max_col=2, max_row=ws.max_row) #categories
	#styling 
	#s280.marker.graphicalProperties.solidFill = "314474"
	colors = ("314474","BC4B4B","bbbbbb" )
	symbols = ("triangle", "dash", "circle")
	for pos  in  range(len(titles) ):
		s = lineChart.series[pos]
		s.graphicalProperties.line.solidFill  = colors[pos]
		#s.graphicalProperties.line.width = 0.5
		
		#s.smooth = True
		#s.marker.graphicalProperties.line.solidFill = colors[pos]
		#s.marker.symbol = symbols[pos]
		#s.marker.size = "2"
		#s.graphicalProperties.line.noFill=True

		
	lineChart.y_axis.majorGridlines = None
	
	
	lineChart.set_categories(cat)
	
	
	
	ws.add_chart(lineChart, "D10")
	wb.save(filepath)
	logging.info("Chart drawn")
	
	
	return
def task():
	global filepath
	if len( sys.argv) == 2:
		filepath = sys.argv[1]
		if os.path.exists(filepath):
			if filepath.endswith(".txt"):
				text_to_rows(filepath)
			else:
				logging.error("Text file required")
		else:
			logging.error("File does not exist")
	else:
		logging.error("#USAGE: python hic_processor.py <filepath>")
	return 
	
	#done see ya!
	
	
	

	
	
if __name__ == "__main__":
	task()