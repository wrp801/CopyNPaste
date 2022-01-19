from pathlib import Path 

def main():
	xl_wings_path = Path.home().joinpath(".xlwings")
	conf_path  = Path.home().joinpath(".xlwings","xlwings.conf")
	module_path = xl_wings_path.joinpath("CopyNPaste","copy_n_paste")
	with open(conf_path,'a') as f:
		f.write(f"\"PYTHONPATH\", \"{module_path}\"")
		f.write('\n')
		f.write(f"\"UDF MODULES\", \"copy_n_paste \"")
	
	f.close()

if __name__ == '__main__':
	main()
	print("Config file has been updated")
