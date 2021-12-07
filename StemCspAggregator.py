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

# Constants
trace_debugging = False
    # Category weights (towards overall grade)
exercise_weight = .1
project_weight = .1
performance_weight = .1
quiz_weight = .2
exam_weight = .5
    # Category names (what they will be called in output aggregate file)
exercise_cat_name = "Exercises"
project_cat_name = "Project"
performance_cat_name = "Performance"
quiz_cat_name = "Quizzes"
exam_cat_name = "Exam"

# Function for trace debugging
def trace(msg):
    if trace_debugging:
        print (msg)

# The name of this aggregator
def name ():
    return "Stem CS Aggregator"

# Get the default input file = the latest exported STEM CSP grade sheet in the download folder
def get_default_input_file ():
    year = str(datetime.datetime.now().year)
    download_dir = GradeUtils.get_download_dir()
    return GradeUtils.get_latest (download_dir + "\\" + year + "*Grades-AP_CS_Principles*.csv")

# Aggregate an input file into an output file
def aggregate (input_file, output_file):
    # Read the input file, drop the columns we don't need, and index on student name and section
    df = pd.read_csv(input_file)                                # Read the file
    df.drop(labels=["ID","SIS User ID", "SIS Login ID"], axis=1, inplace=True)   # Drop student ID columns
    df = df.loc[:, df.iloc[0] != "(read only)"]                 # Drop STEM aggregates
    df = df.loc[:, df.iloc[0] != "0"]                           # Drop unscored columns
    df = df.set_index(["Student", "Section"])                   # Index on student name/section

    # Change the column names to what we want to aggregate on: <unit #> <category>
    # Use the fact that STEM orders columns by category and then unit #
    cat_order = [exercise_cat_name, project_cat_name, quiz_cat_name, exam_cat_name, performance_cat_name]
    cat_num = 0     # Current index into category_order
    unit_num = 0    # Current unit number
    proj_name = None
    col_names = df.columns.tolist()
    for col_ix in range(len(col_names)):
        col_name = col_names[col_ix]
        col_name = col_name.replace ("Unit ", "")
        match = re.search ("^\d+", col_name)
        if match is not None:
            if int(match[0]) < unit_num:    # If unit number decreased, we must be at a new category
                cat_num += 1
                trace ("\nStarting category = " + cat_order[cat_num])
            if unit_num != int(match[0]):
                trace ("    Starting unit " + match[0]) # Else it's just a new unit in the same category
            unit_num = int(match[0])
        elif cat_num == 0 and col_name.find("Milestone") != -1: # First project is unit 2
            cat_num += 1
            unit_num = 2
            proj_name = col_name.split()[0]
            trace ("\nStarting category = " + cat_order[cat_num])
            trace ("    Starting unit 2: " + proj_name)
        if cat_num == 1:
            next_proj_name = col_name.split()[0].split(":")[0]
            if proj_name != next_proj_name:  # Each project gets a new unit number
                unit_num += 1
                proj_name = next_proj_name
                trace ("    Starting unit " + str(unit_num) + ": " + proj_name)
        if cat_num == 2:
            assert col_name.find("Quiz") != -1
        if cat_num == 3 and col_name.find("Exam") == -1:
            assert col_name.find("Create") != -1
            cat_num += 1
            trace ("\nStarting category = " + cat_order[cat_num])
        assert unit_num <= 6
        new_col_name = cat_order[cat_num]
        if cat_num != 4:
                new_col_name = str(unit_num) + " " + new_col_name
        col_names[col_ix] = new_col_name
        trace (new_col_name + "\t\tWas: " + col_name)

    df.columns = col_names

    # Keep those columns for which at least half the students have turned something in
    df.dropna (axis=1, thresh=int(df.shape[0]/2), inplace=True)

    # Fill in na's with imputed values, using the mean value of the columns. This handled students
    # who joined very late and didn't do initial work, or students who are a little late on recent assignments
    # There should be none of these long term, as teachers should enter "0" grade at some point
    #df = df.apply(lambda x: x.fillna(x.mean()),axis=0)

    # Sum up all columns with the same name to produce aggregate points by unit and category
    df = df.groupby(by=df.columns, axis=1).sum()

    # Convert everything to an int, since imputed values create messy decimals
    df = df.astype(int)

    # Also create columns showing overall "score" (percent points) by category, across all units
    df["  "] = [None] * df.shape[0] # Blank column as separator for the aggregates
    for cat in cat_order:
        cols = [col for col in df.columns if cat in col]
        col_name = cat + " grade"
        df[col_name] = df[cols].sum(axis=1)
        df[col_name] = (.49 + 100*df[col_name]/df[col_name][0]).astype(int)

    # Finally, create a column of overall score for sanity
    df["   "] = [None] * df.shape[0] # Blank column as separator for the aggregates
    df["Overall grade"] = df[exercise_cat_name + " grade"] * exercise_weight + \
        df[project_cat_name + " grade"] * project_weight + \
        df[performance_cat_name + " grade"] * performance_weight + \
        df[quiz_cat_name + " grade"] * quiz_weight + \
        df[exam_cat_name + " grade"] * exam_weight
    df["Overall grade"] = df["Overall grade"].astype(int)

    # Save the results to the output file
    df.to_csv(output_file)

# Wrap the functions above into a standard interface for GradeAggregtor.py
class StemCspAggregator:
    def __init__ (self):
        self.name = name
        self.aggregate = aggregate
        self.get_default_input_file = get_default_input_file