

import os, sys, plc, logging
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")

def task(directory):
	
	for filename in os.listdir(directory):
		if filename.endswith(".xlsx"):
			filepath = os.path.join(directory, filename)
			plc.plot_line_chart(filepath)
			logging.info("Chart  added to : " + filename)

	logging.info("Completed")


if __name__ == "__main__":
	if len( sys.argv ) == 2:
		source = os.path.abspath(sys.argv[1])
		if os.path.exists(source):
			task(source)
		else:
			print("The source directory is invalid")
	else:
		print("Requires source directory and target directory")