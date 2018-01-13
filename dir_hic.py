
import hic_processor, sys, os 


def task():
	if len( sys.argv ) == 2:
		path = os.path.abspath(sys.argv[1])
		if os.path.exists(path):
			for filename in os.listdir(path):
				if filename.endswith(".txt"):
					
		else:
			print("Path does not exists")
	else:
		print("Provide Directory path of HIC Files")



if __name__ == "__main__":
	task()