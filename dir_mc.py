#USAGE python dir_mc.py sourcepath 

import os, sys, mc, logging, transition_mc

logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")


def task(sourcepath: str):
	files = []
	for file in os.listdir(sourcepath):
		if file.endswith(".xlsx"):
			files.append( os.path.join( sourcepath, file ) )
	if len(files) > 0:
		transition_mc.get_data_to_dict(files, sourcepath)
		logging.info("Chart created")
	return
	
	
if __name__ == "__main__":
	if len( sys.argv ) == 2:
		sourcepath = sys.argv[1]
		if os.path.exists(sourcepath):
			task(sourcepath)
		else:
			logging.error("Source path does not exist")
	else:
		logging.error("#USAGE python dir_mc.py sourcepath ")