import subprocess
import os

os.system('taskkill /f /im time_class.exe')
subprocess.Popen('time_class.exe')