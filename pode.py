
#USAGE : python pode.py <source path> <target path>
import getdata, os, sys 


def task(source , target ):
	for filename in os.listdir(source):
		if(filename.endswith(".ods")):
			filepath = os.path.join(source, filename)
			getdata.ods_to_excel(filepath, target)
		
		
		

if __name__ == "__main__":
	if len( sys.argv ) == 3:
		source = os.path.abspath(sys.argv[1])
		target = os.path.abspath(sys.argv[2])
		if not os.path.exists(target):
			os.mkdir(target)
		if os.path.exists(source):
			task(source, target)
		else:
			print("The source directory is invalid")
	else:
		print("Requires source directory and target directory")