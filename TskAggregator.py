"""
TsKAggregator - take an exported spreadsheet from TSK and and aggregate it in a manner suitable for import to Synergy

This script provides an interface which is called from GradeAggregator, and is not meant to be called directly

The aggregation works as follows:
* For each student, aggregates are calculated per lesson:
    * An aggregate score is computed for assignments, quizzes/lesson checks, and the unit exam
    * Each assignment is scored 0 or 1: 1 point if the work turned in, has no syntax errors, and contains
      at half the number of expected lines of code (coding problems) or half the number of correct answers 
      (for warm-ups); otherwise it's a 0
    * Exams, quizzes, and lesson check scores are the number of questions answered correctly
* For each student, overall grades are also calculated:
    * Exam grade = % of points on unit exams
    * Quiz grade = % of points on quizzes and lesson checks
    * Assignment grade = % of points on assignments
    * Overall grade = 50% assignments + 30% unit exams + 20% quizzes and lesson checks (weights are configurable)

Prereqs for running this script: Python 3.5+ + pandas + openpyxl.
E.g., if Python is installed from software center, then navigating to the scripts directory and:
pip install pandas      - dataframes for data maniupulation
pip install openpyxl    - to parse .xlsx files
"""

# Libraries
import pandas as pd
import re
import GradeUtils
from GradeUtils import trace, println

# The name of this aggregator
def name ():
    return "Tech Smart Kids Aggregator"

# The file pattern we check for files exported from STEM
def get_input_file_pattern ():
    download_dir = GradeUtils.get_download_dir()
    return download_dir + r"\CS20*.xlsx"

# Get the default input file = the latest exported TSK grade sheet in the download folder
def get_default_input_file ():
    return GradeUtils.get_latest (get_input_file_pattern ())

# Aggregate an input file into an output file
def aggregate (input_file, output_file):
    # Read the input file
    df = pd.read_excel(input_file, header=None)

    # There are six header rows; we'll extract the info we need, create column names based on that, and then remove the header rows
    # In particular, we'll extract info from rows 0 (unit number), 1 (lesson number) and 3 (category)
    # We don't use info in rows  2 (exercise title), 4 (date in was done in class), or 5
    # We'll put the results (what we indend to be the column names) in a list called col_names
    unit_num="0"
    lesson=""
    col_names = ["Last name", "First name", "ID"]   # The first three columns are fixed
    # Parsing logic: A nan header value means "use the previous row value", and non-nan values needs parsing
    for i in range (3, df.shape[1]):
        if pd.isna(df.loc[0][i]) == False:
            unit_num = df.loc[0][i][5]                          # Row 0 format is "Unit #: xxx", so 5th element is the actual #
        if pd.isna(df.loc[1][i]) == False:
            l = df.loc[1][i]                                    # Row 1 format is "Lesson #: xxx" - let's extract the lesson #
            l = l.replace ("Lesson ", "")                       # Now we have just "#: xxx"
            l = l[:l.find(":")]                                 # Now we have just "#"
            if l == "Q":                                        # If a quiz, append Q to lesson number to lesson strings are ordered by due data
                lesson += "Q"
            elif l.isdigit() or l=="T":                         # If it's a numberic lesson number of test or quiz, create a new lesson name for aggregation 
                lesson = unit_num + "." + l                     # We don't create new lesson aggregate names for the little "P" "PLx, "RA" exercises
        if df.loc[3][i] == "Assessment":
            if df.loc[2][i].startswith("Practice Test"):
                col_names.append(lesson + " Assignment")        # Treat practice tests like assignments
            elif df.loc[2][i].endswith("Lesson Check"):
                col_names.append(lesson + " Assignment")        # Treat the lesson checks like assignments
            elif lesson.endswith("T"):
                col_names.append(lesson + " Exam")              # Lesson x.T assessments are exams
            else:
                col_names.append(lesson + " Quiz")              # Other assessments are lesson check or unit quizzes
        else:
            col_names.append(lesson + " Assignment")            # Everything else is an assignment

    # Set the column names and drop the no-longer needed header rows    
    df.columns = col_names
    df.drop(labels=range(0,5), axis=0, inplace=True)

    # Create an index column called "Student" of the form "last, first", drop the other three header columns
    df.insert(0, "Student", df["Last name"].str.cat(df["First name"], sep=", "))
    df.drop(columns=["Last name", "First name", "ID"], inplace=True)

    # Keep those columns for which at least 1/4 the students have turned something in
    df.dropna (axis=1, thresh=int(df.shape[0]/4), inplace=True) # Can change to how='all" instead of thresh to only drop if no one has turned something in

    # Now let's clean up the cells - replacing text with numbers we can use for scoring
    # Work columns are scored either 1 (submitted, no syntax errors, at least half the expected lines of code), else 0
    # Assessment columns are scored based on number of questions answered correctly
    df = df.fillna(0)                                                           # Not started = 0
    df = df.replace(to_replace ='In progress', value = 0, regex = True)         # In progress = 0
    df = df.replace(to_replace ='In Progress', value = 0, regex = True)         # In Progress = 0
    df = df.replace(to_replace ='.*Syntax error.*', value = 0, regex = True)    # Syntax error = 0
    df = df.replace(to_replace ='Turned In', value = 1, regex = True)           # Turned in = 1
    df = df.replace(to_replace =' lines of code.*', value = "", regex = True)   # Remove text we don't need
    df = df.replace(to_replace ='\n.*', value = "", regex = True)               # Remove text we don't need
    df = df.replace(to_replace =' \(.+\)', value = "", regex = True)            # Remove text we don't need
    
    # Now all cells have numeric values except for a few cells of the form a/b
    # For Work columns, transform to 1 if a/b > 1/2; else to 0
    # If Assessment columns, transform to a, but remember that column is worth b points
    # While doing this, keep track of the max points each row is work
    # During the iteration, temporarily rename the columns so they have unique names
    max_score = [1] * df.shape[1]   # By default, assume max points for a column is 1 (overridden below for assessment columns)
    max_score[0] = "Max score"
    columns = df.columns
    df.columns = range(0, df.shape[1])
    for c in df.columns[2:]:
        for r in df.index:
            val = df.loc[r][c]
            if type(val) == str:
                match = re.findall("\d+", val)
                a = int(match[0])
                b = int(match[1])
                if columns[c].find("Assignment") != -1:
                    if a/b >= .5:
                        df.at[r,c] = 1    # Was df.loc[r][c] = 1
                    else:
                        df.at[r, c] = 0    # As above
                else:
                    df.at[r, c] = a        # Ass above
                    max_score[c] = b
    df.columns = columns

    # Finally, add a header row that represents the max points that this row is worth
    df.loc[0] = max_score
    df = df.sort_index()
    df.index = range (df.shape[0])

    # Sum up all columns with the same name
    df = df.set_index("Student")
    df = df.apply(pd.to_numeric)    # Sanity check that all cells are numeric
    df = df.groupby(by=df.columns, axis=1).sum()

    # Add a dummy section column at the front
    df.insert(0, "Section", ["Default"] * df.shape[0])

    # Save the results to the output file
    df.to_csv(output_file)

class TskAggregator:
    def __init__ (self):
        self.name = name
        self.aggregate = aggregate
        self.get_default_input_file = get_default_input_file
        self.get_input_file_pattern = get_input_file_pattern
