"""
StemCsaAggregator.py - take an exported spreadsheet from Project STEM for the AP Comp Sci A class, and
aggregates it in a manner suitable for Synergy grading
* Input: Optional path grading .csv exported from STEM; if missing the the most recent .csv file in the download folder is used
* Output: A file named "Aggregate " + inputFileName, which is auto-launched into excel

The aggregation works as follows:
* For each student, aggregates are calculated per lesson and category
* For each student, overall grades are also calculated as sanity check that our calculation = STEM calculation
* Columns are dropped unless at least half the students have a recorded grade (non-na value)
* Missing values (na's) for non-dropped columns are imputed as avg values for each column to handle students
* who joined late or missed have a late assignment.

Prereqs for running this script: Python 3.5+ + pandas + openpyxl
E.g.; if Python is installed from software center, then navigating to the scripts directory and:
pip install pandas      - dataframes for data manipulation
"""

# Libraries
import pandas as pd
import re
import sys
import os
import glob
import winreg
import subprocess
import datetime

# Constants
trace_debugging = False
    # Category weights (towards overall grade)
exercise_weight = .1
quiz_weight = .3
exam_weight = .6
    # Category names (what they will be called in output aggregate file)
exercise_cat_name = "Exercises"
quiz_cat_name = "Quiz and assignment"
exam_cat_name = "Exam"

# Function for trace debugging
def trace(msg):
    if trace_debugging:
        print (msg)

# Get the input file
usage = """
SteamCsaAggregator.py [fileName]
Where fileName is an exported grading spreadsheet.
If filename is missing, default to finding the most recent .csv file in the download folder
"""
input_file=""
# If the file is explicitly specified as command line argument, just check that the file exists
if len(sys.argv) > 1:
    input_file=sys.argv[1]
    if not os.path.exists(input_file):
        print ("File not found: " + input_file)
        print (usage)
        sys.exit(1)
# Else use the latest CSP grades .csv file in the download folder (where STEM exports to)
else:
    year = str(datetime.datetime.now().year)
    list_of_files = glob.glob(os.getenv("USERPROFILE") + "\\downloads\\" + year + "*Grades-AP_CS_A*.csv")
    input_file = max(list_of_files, key=os.path.getctime)
    print ("Using " + input_file)

# Read the input file, drop the columns we don't need, and index on student name and section
df = pd.read_csv(input_file)                                # Read the file
df.drop(labels=["ID","SIS User ID", "SIS Login ID"], axis=1, inplace=True)   # Drop student ID columns
df = df.loc[:, df.iloc[0] != "(read only)"]                 # Drop STEM aggregates
df = df.loc[:, df.iloc[0] != "0"]                           # Drop unscored columns
df = df.set_index(["Student", "Section"])                   # Index on student name/section
df.dropna (axis=1, thresh=int(df.shape[0]/2), inplace=True) # Only keep columns if >1/2 students have submitted
#df = df.apply(lambda x: x.fillna(x.mean()),axis=0)          # Add imputed values for missing entries

# Change the column names to what we want to aggregate on: <unit #> <category>
# The category can be figured out based on a regular expression applied to the column name
# Unit number is always the first number in the text (for all categories)
cat_info = [
    (exam_cat_name, "^Unit \d+ Exam"),              # Unit N Exam (id)
    (exercise_cat_name, "^Unit \d+: Lesson"),       # Unit N: Lesson M - title (id)
    (quiz_cat_name, "^(Unit \d+ Quiz|Assignment)")  # Unit N Quiz (id) | Assignment N (id)
]                

col_names = df.columns.tolist()
for col_ix in range(len(col_names)):
    col_name = col_names[col_ix]
    match = re.search ("\d+", col_name)
    if match is None:
        print ("Warning: can't parse unit number from column " + col_name + ". Skipping....")
        continue
    unit_num = match[0]
    category = None
    for cat in cat_info:
        if re.search(cat[1], col_name) is not None:
            category = cat[0]
            break
    if category is None:
        print ("Warning: can't parse category from column " + col_name + ". Skipping....")
        continue
    
    new_col_name = unit_num + " " + category
    col_names[col_ix] = new_col_name
    trace (new_col_name + "\t\tWas: " + col_name)

df.columns = col_names

# Sum up all columns with the same name to produce aggregate points by unit and category
df = df.groupby(by=df.columns, axis=1).sum()

# Convert everything to an int, since imputed values create messy decimals
df = df.astype(int)

# Also create columns showing overall "score" (percent points) by category, across all units
df["  "] = [None] * df.shape[0] # Blank column as separator for the aggregates
for c in cat_info:
    cat = c[0]
    cols = [col for col in df.columns if cat in col]
    col_name = cat + " grade"
    df[col_name] = df[cols].sum(axis=1)
    df[col_name] = (.49 + 100*df[col_name]/df[col_name][0]).astype(int)

# Finally, create a column of overall score for sanity
df["   "] = [None] * df.shape[0] # Blank column as separator for the aggregates
df["Overall grade"] = df[exercise_cat_name + " grade"] * exercise_weight + \
    df[quiz_cat_name + " grade"] * quiz_weight + \
    df[exam_cat_name + " grade"] * exam_weight
df["Overall grade"] = df["Overall grade"].astype(int)

# Save the results to the output file and launch excel
output_file = os.path.dirname(input_file) + "\\Aggregated " + os.path.basename(input_file)
df.to_csv(output_file)
excel_path = winreg.QueryValue(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe")
subprocess.run ([excel_path,  output_file])
