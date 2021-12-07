import sys
import os
import glob
import winreg
import subprocess

# Launch excel on a given spreadsheet file
def launch_excel (file):
    excel_path = winreg.QueryValue(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe")
    subprocess.run ([excel_path,  file])

# Get the current users download directory
def get_download_dir ():
    return os.getenv("USERPROFILE") + r"\downloads"

# Get the latest file from a file path that may contain wildcards (e.g., "*.xlsx")
def get_latest (file):
    list_of_files = glob.glob(file)
    if len(list_of_files) == 0:
        return None
    return max(list_of_files, key=os.path.getctime)

# See if there is a file specified as parameter in command line
def get_argv_file ():
    if len(sys.argv) <= 1:
        return None
    input_file=sys.argv[1]
    if not os.path.exists(input_file):
        print ("File not found: " + input_file)
        return None
    return input_file

# Create a new filename from an existing one by appending a filename prefix
def get_output_file_name (existing_file, prefix):
    return os.path.dirname(existing_file) + "\\" + prefix + os.path.basename(existing_file)

# Check if output_file exists and is newer than input_file
def is_current (output_file, input_file):
    if not os.path.exists(output_file):
        return False
    if os.path.getctime (output_file) < os.path.getctime (input_file):
        return False
    return True
