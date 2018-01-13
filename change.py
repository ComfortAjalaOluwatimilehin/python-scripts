#change data


import os, sys, openpyxl


def task(filepath):
	wb = openpyxl.load_workbook(filepath)
	ws = wb.active 









if __name__ == "__main__":
	if len( sys.argv ) == 2:
		path = sys.argv[1]
		filepath = os.path.abspath(path)
		if os.path.exists(filepath):
			task(filepath)
		else:
			print("Requires existing file")
	else:
		print("Requires file path")