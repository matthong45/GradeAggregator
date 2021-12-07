"""
TsKAggregator - take an exported spreadsheet from TSK and turn it into a grade
* Input: Optional path to a grading .xlsx exported from TSK; if missing the the most recent .xlsx file in the download folder is used
* Output: A file named "Aggregate " + inputFileName, which is auto-launched into excel

The aggregation works as follows:
* For each student, aggregates are calculated per lesson:
    * An aggregate score is computed for the "work" (homework + classwork) and the assessments
    * The work score is calculated by giving 1 point if the work turned in, has no syntax errors, and contains
      at half the number of expected lines of code (coding problems) or half the number of correct answers 
      (for warm-ups); otherwise it's a zero
    * The assessment score is the number of questions answered correctly
* For each student, overall grades are also calculated:
    * Assessment grade = % of assessment points scores
    * Work grade = % of work points scored
    * Overall grade = 50% work grade + 50% assessment grade (weights are configurable)

Prereqs for running this script: Python 3.5+ + pandas + openpyxl.
E.g., if Python is installed from software center, then navigating to the scripts directory and:
pip install pandas      - dataframes for data maniupulation
pip install openpyxl    - to parse .xlsx files
"""
# Configuration variables
work_score_weight = .5
assessment_score_weight = .5

# Libraries
import pandas as pd
import re
import sys
import GradeUtils

# Get the input file
usage = """
TskAggregator.py [fileName]
Where fileName is an exported TechSmartKids grading spreadsheet.
If filename is missing, default to finding the most recent .xlsx file in the download folder
"""

# Get the input file to process; either command line arg or latest export in download directory
def get_input_file ():
    input_file=GradeUtils.get_argv_file()
    if input_file is None:
        download_dir = GradeUtils.get_download_dir()
        input_file = GradeUtils.get_latest (download_dir + r"\CS20*.xlsx")
        if input_file is None:
            print ("Can't find CS20*.xlsx in " + download_dir)
    return input_file

#
def tsk_aggregate (input_file, output_file):
    # Read the input file
    df = pd.read_excel(input_file, header=None)

    # There are six header rows; we'll extract the info we need, create column names based on that, and then remove the header rows
    # In particular, we'll extract info from rows 0 (unit number), 1 (lesson number) and 3 (category)
    # We don't use info in rows  2 (exercise title), 4 (date in was done in class), or 5
    # We'll put the results (what we indend to be the column names) temporarily in row 0
    unit_num="0"
    lesson=""
    # Parsing logic: A nan header value means "use the previous row value", and non-nan values needs parsing
    for i in range (3, df.shape[1]):
        if pd.isna(df.loc[0][i]) == False:
            unit_num = df.loc[0][i][5]                          # Row 0 format is "Unit #: xxx", so 5th element is the actual #
        if pd.isna(df.loc[1][i]) == False:
            lesson = df.loc[1][i].replace ("Lesson ", "")       # Row 1 format is "Lesson #: xxx", so extract the lesson number
            lesson = unit_num + "." + lesson[:lesson.find(":")] # And transform to "u.l" (u = unit #, l = lesson #)
        if df.loc[3][i] == "Assessment":
            df.loc[0][i] = lesson + " Assessment"               # And append the category (if it's Work or Assessment)
        else:
            df.loc[0][i] = lesson + "  Work"
    # First three column names are fixed and don't need parsing; first name, last name, and student ID
    df.loc[0][0] = "Last name"
    df.loc[0][1] = "First name"
    df.loc[0][2] = "Student ID"
    # Now set the data frame column names to the above (format is "u.l category"), and then remove the no longer needed header rows
    df.columns = df.loc[0]
    df = df.drop(labels=range(0,5), axis=0)
    # Drop the student ID column - we do't need it
    df = df.drop(labels="Student ID", axis=1)
    # Drop na columns - these are things that were skipped or not yet reached
    df = df.dropna(axis=1, how='all')

    # Now let's clean up the cells - replacing text with numbers we can use for scoring
    # Work columns are scored either 1 (submitted, no syntax errors, at least half the expected lines of code), else 0
    # Assessment columns are scored based on number of questions answered correctly
    df = df.fillna(0)                                                           # Not started = 0
    df = df.replace(to_replace ='In progress', value = 0, regex = True)         # In progress = 0
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
    max_score[1] = ""
    columns = df.columns
    df.columns = range(0, df.shape[1])
    for c in df.columns[2:]:
        for r in df.index:
            val = df.loc[r][c]
            if type(val) == str:
                match = re.findall("\d+", val)
                a = int(match[0])
                b = int(match[1])
                if columns[c].find("Work") != -1:
                    if a/b >= .5:
                        df.loc[r][c] = 1
                    else:
                        df.loc[r][c] = 0
                else:
                    df.loc[r][c] = a
                    max_score[c] = b
    df.columns = columns

    # Finall, add a header row that represents the max points that this row is worth
    df.loc[0] = max_score
    df = df.sort_index()
    df.index = range (df.shape[0])

    # Sum up all columns with the same name
    df = df.set_index(["Last name", "First name"]) 
    df = df.groupby(by=df.columns, axis=1).sum()

    # Compute overall assessment grade (% of points scored across all lessons)
    assessment_cols = [col for col in df.columns if 'Assessment' in col]
    assessment_col_name = "Assessment grade"
    df[assessment_col_name] = df[assessment_cols].sum(axis=1)
    df[assessment_col_name] = (.49 + 100*df[assessment_col_name]/df[assessment_col_name][0]).astype(int)

    # Compute overall work grade (% of points scored across all lessons)
    work_cols = [col for col in df.columns if 'Work' in col]
    work_col_name = "Work grade"
    df[work_col_name] = df[work_cols].sum(axis=1)
    df[work_col_name] = (.49 + 100*df[work_col_name]/df[work_col_name][0]).astype(int)

    # Compute overall grade (weighted avg of assessment and work grades)
    df["Overall grade"] = df[work_col_name] * work_score_weight + df[assessment_col_name] * assessment_score_weight

    # Save the results to the output file and launch excel
    df.to_excel(output_file)
    GradeUtils.launch_excel (output_file)

# Get the input file to process, reaggregate if (derived) aggregate output file is not current
input_file = get_input_file()
if input_file is None:
    print ("Skipping TechSmart aggregation since there are no files to aggregate")
else:
    print ("Using " + input_file)
    output_file = GradeUtils.create_file_name(input_file, "Aggregated ")
    if GradeUtils.is_current (output_file, input_file):
        print ("Aggregate file is already current: " + output_file)
    else:
        tsk_aggregate (input_file, output_file)

