"""
StemCsaAggregator.py - take an exported AP Comp Sci A spreadsheet from project STEM and aggregate it in a manner suitable for import to Synergy

This script provides an interface which is called from GradeAggregator, and is not meant to be called directly

The aggregation works as follows:
* For each student, aggregates are calculated per lesson and category
* For each student, overall grades are also calculated as sanity check that our calculation = STEM calculation
* Columns are dropped unless at least half the students have a recorded grade (non-na value)

Prereqs for running this script: Python 3.5+ + pandas + openpyxl
E.g.; if Python is installed from software center, then navigating to the scripts directory and:
pip install pandas      - dataframes for data manipulation
"""

# Libraries
import pandas as pd
import re
import datetime
import GradeUtils
from GradeUtils import trace, println

# Category names (what they will be called in output aggregate file)
exercise_cat_name = "Exercises"
quiz_cat_name = "Quiz and assignment"
exam_cat_name = "Exam"

# The name of this aggregator
def name ():
    return "STEM AP Comp Sci A Aggregator"

# The file pattern we check for files exported from STEM
def get_input_file_pattern ():
    year = str(datetime.datetime.now().year)
    download_dir = GradeUtils.get_download_dir()
    return download_dir + "\\" + year + "*Grades-*AP_CS_A*.csv"

# Get the default input file = the latest exported STEM CSP grade sheet in the download folder
def get_default_input_file ():
    year = str(datetime.datetime.now().year)
    download_dir = GradeUtils.get_download_dir()
    return GradeUtils.get_latest (get_input_file_pattern ())

# Aggregate an input file into an output file
def aggregate (input_file, output_file):
    # Read the input file, drop the columns we don't need, and index on student name and section
    df = pd.read_csv(input_file)                                # Read the file
    df.drop(labels=["ID","SIS User ID", "SIS Login ID"], axis=1, inplace=True)   # Drop student ID columns
    df = df.loc[:, df.iloc[0] != "(read only)"]                 # Drop STEM aggregates
    df = df.loc[:, df.iloc[0] != "0"]                           # Drop unscored columns
    df = df.set_index(["Student", "Section"])                   # Index on student name/section
    df.dropna (axis=1, thresh=int(df.shape[0]/4), inplace=True) # Only keep columns if >1/4 students have submitted
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
        if col_name.startswith("FRQ"):
            col_names[col_ix] = "drop me"
            continue
        match = re.search ("\d+", col_name)
        if match is None:
            println ("Warning: can't parse unit number from column " + col_name + ". Skipping....")
            col_names[col_ix] = "drop me"
            continue
        unit_num = match[0]
        category = None
        for cat in cat_info:
            if re.search(cat[1], col_name) is not None:
                category = cat[0]
                break
        if category is None:
            println ("Warning: can't parse category from column " + col_name + ". Skipping....")
            col_names[col_ix] = "drop me"
            continue
        
        new_col_name = unit_num + " " + category
        col_names[col_ix] = new_col_name
        trace (new_col_name + "\t\tWas: " + col_name)

    df.columns = col_names

    # drop all columns named "drop me"
    df = df.loc[:, df.columns != "drop me"]

    # Sum up all columns with the same name to produce aggregate points by unit and category
    df = df.groupby(by=df.columns, axis=1).sum()

    # Convert everything to an int, since imputed values create messy decimals
    df = df.astype(int)

    # Save the results to the output file and launch excel
    df.to_csv(output_file)

# Wrap the functions above into a standard interface for GradeAggregtor.py
class StemCsaAggregator:
    def __init__ (self):
        self.name = name
        self.aggregate = aggregate
        self.get_default_input_file = get_default_input_file
        self.get_input_file_pattern = get_input_file_pattern
