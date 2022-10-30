"""
StemCspAggregator.py - take an exported AP Comp Sci Principals spreadsheet from project STEM and and aggregate it in a manner suitable for import to Synergy

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
import GradeUtils
import datetime
from GradeUtils import trace, println

# Category names (what they will be called in output aggregate file)
exercise_cat_name = "Exercises"
project_cat_name = "Project"
performance_cat_name = "Create task"
quiz_cat_name = "Quizzes"
exam_cat_name = "Exam"

# The name of this aggregator
def name ():
    return "STEM AP Comp Sci Principals Aggregator"

# The file pattern we check for files exported from STEM
def get_input_file_pattern ():
    year = str(datetime.datetime.now().year)
    download_dir = GradeUtils.get_download_dir()
    return download_dir + "\\" + year + "*Grades-AP_CS_Principles*.csv"

# Get the default input file = the latest exported STEM CSP grade sheet in the download folder
def get_default_input_file ():
    return GradeUtils.get_latest (get_input_file_pattern ())

# Aggregate an input file into an output file
def aggregate (input_file, output_file):
    # Read the input file, drop the columns we don't need, and index on student name and section
    df = pd.read_csv(input_file)                                # Read the file
    df.drop(labels=["ID","SIS User ID", "SIS Login ID"], axis=1, inplace=True)   # Drop student ID columns
    df = df.loc[:, df.iloc[0] != "(read only)"]                 # Drop STEM aggregates
    df = df.loc[:, df.iloc[0] != "0"]                           # Drop unscored columns
    df = df.set_index(["Student", "Section"])                   # Index on student name/section

    # Change the column names to what we want to aggregate on: <unit #> <category>
    col_names = df.columns.tolist()
    for col_ix in range(len(col_names)):
        col_name = col_names[col_ix].lower().replace("unit ", "")
        new_col_name = None     # The new column name we will aggregate on
        
        # Most columns are exercises, quizzes, or exams starting with the unit number
        match = re.search ("^(unit )*\d", col_name)
        if match is not None:
            if "exercise" in col_name or "ap-style" in col_name or "review" in col_name or "additional practice" in col_name:
                new_col_name = match[0] + " " + exercise_cat_name
            elif "quiz" in col_name:
                new_col_name = match[0] + " " + quiz_cat_name
            elif "exam" in col_name:
                new_col_name = match[0] + " " + exam_cat_name
            elif "0.5 " in col_name or "0.6 " in col_name:
                new_col_name = "1 " + exercise_cat_name

        # big picture exercises don't include the unit number - must figure it out from the name
        elif "big picture" in col_name:
            if "collaboration" in col_name:
                new_col_name = "2 " + exercise_cat_name
            if "moore" in col_name:
                new_col_name = "2 " + exercise_cat_name
            elif "reselling" in col_name:
                new_col_name = "3 " + exercise_cat_name
            elif "ethics" in col_name or "intellectual" in col_name:
                new_col_name = "4 " + exercise_cat_name
            elif "data" in col_name:
                new_col_name = "5 " + exercise_cat_name  
            elif "innovation" in col_name or "divide" in col_name or "neutrality" in col_name:
                new_col_name = "5 " + exercise_cat_name

        # projects don't include the unit number - must figure it out from the name
        elif "milestone" in col_name or "final project submission" in col_name:
            if "password" in col_name:
                new_col_name = "2 " + project_cat_name
            elif "unintend" in col_name:
                new_col_name = "3 " + project_cat_name
            elif "image" in col_name:
                new_col_name = "4 " + project_cat_name
            elif "tedx" in col_name:
                new_col_name = "5 " + project_cat_name
            elif "exploring" in col_name:
                new_col_name = "6 " + project_cat_name

        # Some columns called "tedxkinda: xxx" that are just unit 5 exercises
        elif "tedxkinda: " in col_name:
            new_col_name = "5 " + exercise_cat_name

        # Columns wih the name "question type: " are unit 7 AP review exercises
        elif "question type: " in col_name or "ap cb practice" in col_name:
            new_col_name = "7 " + exercise_cat_name

        # And finally, the following are create-task assignments
        if "create task" in col_name or "mini create" in col_name or "peer review" in col_name:
            new_col_name = performance_cat_name

        # Some of the tedxkind
        if new_col_name is None:
            println ("Can't translate\t\t: " + col_name, file=sys.stderr)
        else:
            col_names[col_ix] = new_col_name
            trace (new_col_name + "\t\t<- " + col_name)

    df.columns = col_names

    # Keep those columns for which at least 1/4 the students have turned something in
    df.dropna (axis=1, thresh=int(df.shape[0]/4), inplace=True)

    # Fill in na's with imputed values, using the mean value of the columns. This handled students
    # who joined very late and didn't do initial work, or students who are a little late on recent assignments
    # There should be none of these long term, as teachers should enter "0" grade at some point
    #df = df.apply(lambda x: x.fillna(x.mean()),axis=0)

    # Sum up all columns with the same name to produce aggregate points by unit and category
    df = df.groupby(by=df.columns, axis=1).sum()

    # Convert everything to an int, since imputed values create messy decimals
    df = df.astype(int)
    
    # Save the results to the output file
    df.to_csv(output_file)

# Wrap the functions above into a standard interface for GradeAggregtor.py
class StemCspAggregator:
    def __init__ (self):
        self.name = name
        self.aggregate = aggregate
        self.get_default_input_file = get_default_input_file
        self.get_input_file_pattern = get_input_file_pattern
