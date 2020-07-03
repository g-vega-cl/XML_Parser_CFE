import subprocess
import sys
import os 

directory = os.getcwd()
print("curr_dir = ", directory)
path = directory + "{}".format("\Install_dependencies.bat")
print(path)
subprocess.call([path])