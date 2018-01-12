#USAGE python getdata.py 

		
from pyexcel_ods import get_data
import os , re, json, openpyxl, logging, typing, sys

logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")

	
def get_rows_xy(sheet: typing.List ):
	'''
		argument:
			sheet : a multidimensional list
		Gets all the xy data 
		returns : list 
	'''
	start = False #when true, starts appending to the data list
	data = []
	dictionary = {"y":None, "x":None}
	for sub_list in sheet:
		if len(sub_list) == 0:
			break
		if start:
			data.append(sub_list[:2])
		if "XUNITS" in sub_list:
			dictionary["x"] = sub_list[1]
		if "YUNITS" in sub_list:
			dictionary["y"] = sub_list[1]
		if "XYDATA" in sub_list: #searches for XYDATA string in a sublist
			start = True 
	dictionary["data"] = data
	return comma_to_dot(dictionary)

def comma_to_dot(dictionary: dict ):
	'''
		transform a German format number to US format e.g. 100, 36 to 100.36
		arguments : dictionary -> dict 
		returns:  data -> transformed rows (all values are numbers with "," replaced with ".")
	'''
	data = []
	pattern = re.compile(r"(\d+),(\d+)")
	for sub_row in dictionary["data"]:
		temp = []
		for value in sub_row:
			if type(value) == int:
				#u = float(value)  / 1000
				#u = str(u)
				#print(str(value) + " = " + u)
				value = float(value)  / 1000
				value = str(value)
			update = re.sub(pattern.pattern, r"\1.\2", str(value))
			update = float(update)
			temp.append(update)
		data.append(temp)
	dictionary["data"] = data
	return dictionary		

def add_rows_to_sheet(dictionary : dict, sheet):
	'''
		Adds data to worksheet 
		arguments :
			rows: list of data 
			sheet: worksheet 
		returns : 
			worksheet (updated with rows)
	'''
	rows = dictionary["data"]
	offset = 0
	if "x" in dictionary and "y" in dictionary:
		sheet.cell(row=1, column=1, value=dictionary["x"])
		sheet.cell(row=1, column=2, value=dictionary["y"])
		offset = 1
	for row in range(len( rows )):
		sub_row = rows[row]
		for col in range(len(sub_row) ):
			sheet.cell(row=row + 1 + offset, column=col + 1, value= sub_row[col] )
	return sheet
	
def ods_to_excel(filepath, target)->None:
	'''
		Generates an excel workbook with ods data 
		argument: 
			filepath --> a valid filepath to the directory containing ods documents 
			target -> a valid filepath to the directory for the generated excel sheets 
		returns:
			None
	'''
	filename = os.path.basename(filepath)
	filename = filename.split(".")[0]
	data = get_data(filepath)
	json_data = json.dumps(data)
	json_data = json.loads(json_data)
	workbook_path = os.path.join( target, filename +  ".xlsx"   )
	if not os.path.exists(workbook_path):
		wb = openpyxl.Workbook()
	else:
		wb = openpyxl.load_workbook(workbook_path)
	for sheetname in json_data:
		ws = wb.active
		ws.title = sheetname[:21]
		data_update = get_rows_xy(json_data[sheetname])
		ws = add_rows_to_sheet(data_update, ws)
		logging.info("Adding data to " + sheetname )
	wb.save(workbook_path)
	logging.info("Complete")
	

if __name__ == "__main__":
	if len(sys.argv) == 2:
		path = sys.argv[1]
		filepath = os.path.join(path, filename)
		if os.path.exists(filepath):
			ods_to_excel(filepath)