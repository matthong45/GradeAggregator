"""
GradeUtils.py - Support methods for the other grading scripts

Author: Marc Shepard
"""
from operator import index
import sys
import os
import glob
import winreg
import subprocess
import datetime
from typing import Callable, OrderedDict
from pandas import DataFrame, read_csv
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from collections import OrderedDict

# Configuration variables
roster_file_name = "Roster.csv"     # File to persist due dates for Synergy bulk import format files
print_func = print                  # Callback for printing (overridden by GUI when aggregation done with GUI)
trace_debugging = False             # For verbose debug output

# Println is a function that can be redirected to a GUI
def println(msg):
    print_func(msg)

# Function for trace debugging
def trace(msg):
    if trace_debugging:
        println (msg)

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
    dir = os.path.dirname(existing_file)
    file_name = os.path.basename(existing_file)
    file_name = os.path.splitext(file_name)[0]       # remove the extension
    return dir + os.path.sep + prefix + file_name + ".csv"

# Check if output_file exists and is newer than input_file
def is_current (output_file, input_file):
    if not os.path.exists(output_file):
        return False
    if os.path.getctime (output_file) < os.path.getctime (input_file):
        return False
    return True


# Check if the string is a date
def is_date (text):
    try:
        datetime.datetime.strptime(text, '%m/%d/%Y')
        return True
    except ValueError:
        return False


# Synergy requires due dates for each assignment, but they are not in the STEM export file
# We'll ask the teacher and persist the data in "Assignment_due_dates.csv"
# As a future enhancement, we should auto-generate this for TSK rather than ask (but for STEM we always need to ask)
def get_assignment_due_dates (course : str, assignments : list[str], callback : Callable) -> dict:
    # If the file "<course> due dates.csv" doesn't exist, create it
    file_name = "Assignment_due_dates.csv"
    if not os.path.exists(file_name):
        println ("Creating " + file_name)
        df = DataFrame()
        for col in ["COURSE", "ASSIGNMENT_NAME", "ASSIGNMENT_DATE"]:
            df[col] = ""
        df.to_csv(file_name, index=False)

    # Create a dictionary of assignment due dates from previously persisted dates + new assignments
    df = read_csv(file_name)
    assignment_dict = OrderedDict()
    for index, row in df.iterrows():
        if row["COURSE"] != course:
            continue
        assignment_dict[row["ASSIGNMENT_NAME"]] = row["ASSIGNMENT_DATE"]
    for assignment in assignments:
        if assignment not in assignment_dict:
            assignment_dict[assignment] = ""

    # Invoke the callback to let user change due dates. Callback returns true if there are changes that need to be saved
    if callback (assignment_dict):
        println ("Saving due dates to \"" + file_name + "\"")
        # Update df to have the latest dates from the dictionary; first update any existing rows
        assignments_in_df = []
        for index in df.index:
            row = df.loc[index]
            if row["COURSE"] != course:
                continue
            row["ASSIGNMENT_DATE"] = assignment_dict[row["ASSIGNMENT_NAME"]]
            assignments_in_df.append(row["ASSIGNMENT_NAME"])
        # Next add new rows as needed
        for assignment, due_date in assignment_dict.items():
            if assignment in assignments_in_df:
                continue
            df.loc[len(df.index)] = [course, assignment, due_date]

        df.to_csv(file_name, index=False)

    # Return the dictionary of assignments and due dates
    return assignment_dict

# Synergy requires due dates for each assignment
# We'll keep this data in a file called "<course> due dates.csv"
# We'll ask teachers to supply the dates when we see new assignments for the first time
# As a future enhancement, we should auto-generate this for TSK rather than ask (but for STEM we always need to ask)
def get_assignment_due_dates_old (course, assignments):
    # If the file "<course> due dates.csv" doesn't exist, create it
    file_name = course + " due dates.csv"
    if not os.path.exists(file_name):
        println ("Creating " + file_name)
        df = DataFrame()
        for col in ["ASSIGNMENT_NAME", "ASSIGNMENT_DATE"]:
            df[col] = ""
        df.to_csv(file_name, index=False)

    # Open "<course> due dates.csv" and ask the teacher to supply dates for any new assignments
    df = read_csv(file_name)
    df = df.set_index("ASSIGNMENT_NAME")
    new_dates = False
    for assignment in assignments:
        if assignment.strip() == "":
            break
        if not assignment in df.index:  # No due date entered - get one
            date = "invalid"
            while date not in ["X", "S"] and not is_date(date):
                date = input ("Enter due date for " + assignment + " (d/m/yyyy, X to permanently skip, S to temporarily skip)? ")
            df.loc[assignment] = [date]
            new_dates = True
        elif str(df.loc[assignment]["ASSIGNMENT_DATE"]) == "X":   # Due date set to perma-skip - ignore assignment
            continue
        else:                           # Any other date was entered; reconfirm for now (TODO - optimize this)
            date = str(df.loc[assignment]["ASSIGNMENT_DATE"])
            new_date = "invalid"
            while new_date not in ["X", "S", ""] and not is_date(new_date):
                new_date = input ("Change due date for " + assignment + " (" + date + ")? CR=keep, d/m/yyyy=new date, X=permanently skip, S=temporarily skip: ")
            if new_date != "" and new_date != date:
                df.loc[assignment] = [new_date]
                new_dates = True

    # If there were any changes, save them back
    if new_dates:
        println ("Saving due dates to \"" + file_name + "\"")
        df.to_csv(file_name, index=True)

    # create a dictionary of assignments -> due dates
    dict = {}
    for assignment in df.index:
        dict[assignment] = df.loc[assignment]["ASSIGNMENT_DATE"]

    return dict

# Student data parsed from Roster.csv
class student:
    def __init__(self, period, course, student_name, sis, alias):
        self.period = period
        self.course = course
        self.id = str(sis)
        self.last_name = student_name.split(",")[0]
        first_name = student_name.replace(self.last_name + ",", "")
        if first_name.endswith("."):
            first_name = first_name[:-3]
        self.first_name = first_name.strip()
        if alias is None or not isinstance (alias, str) or alias == "":
            self.alias = self.last_name + ", " + self.first_name
        else:
            self.alias = alias

    def __str__(self):
        s = self.id + ": " + self.last_name + ", " + self.first_name
        if self.alias != self.last_name + ", " + self.first_name:
            s += " (" + self.alias + ")"
        s += "\t" + "P" + self.period + " " + self.course
        return s

# Teacher must add Roster.csv (exported from Synergy) to this directory to enable Synergy import
def synergy_import_configured():
    return os.path.exists(roster_file_name)

# The synergy bulk import file name will be the same as the input file name, but with a prefix and xlsx extension
def get_synergy_output_dir(input_file):
    return os.path.dirname(input_file)

# Get a dictionary of students info from student name from the roster csv file
def get_roster_dict():
    roster_dict = {}

    df = read_csv(roster_file_name)
    for col in ["Period", "Course Title", "Student Name", "Sis Number"]:
        if col not in df.columns:
            println (roster_file_name + ": invalid format - no column named " + col)
            return None

    for r in range(df.shape[0]):
        row = df.iloc[r]
        if "Alias" in df.columns:
            s = student(row["Period"], row["Course Title"], row["Student Name"], row["Sis Number"], row["Alias"])
        else:
            s = student(row["Period"], row["Course Title"], row["Student Name"], row["Sis Number"], None)
        roster_dict[s.alias] = s

    return roster_dict

def get_assignment_type (course, assignment):
    type = assignment.split(maxsplit=1)[-1]
    if type in ["Assignment", "Exercises"]:
        return "Assignment"
    elif type in ["Quiz", "Quizzes", "Quiz and assignment"]:
        return "Formative Assessment"
    elif type == "Exam":
        return "Summative Assessment"
    elif type == "Project":
        return "Projects"
    elif type == "Performance":
        return "Performance task"
    else:
        return None

# Convert a grade aggregate spreadsheet into Synergy bulk import format
def agg_to_synergy (input_file : str, output_dir : str, due_date_callback : Callable):
    # Read in student roster info - we need to join this to the aggregated data
    roster_dict = get_roster_dict()
    if roster_dict is None:
        println (roster_file_name + " not found - skipping Synergy bulk import formatting")
        return None
    if len(roster_dict) == 0:
        println (roster_file_name + " has no student roster info - skipping Synergy bulk import formatting")
        return None

    # Synergy requires separate bulk import files for each period a class is taught. We'll use a dictionary whose key is the period to keep track separately.
    sdf_dict = {}   # Create a dictionary to hold a dataframe representing each 

    # Open the input file and convert each of it's rows (one per student) to multiple synergy rows (one per assigment per student)
    df = read_csv(input_file)
    course = None
    due_dates = None

    for r in range(1, df.shape[0]):
        row = df.iloc[r]
        student = row["Student"]
        if not student in roster_dict:
            println ("Warning: " + student + " not found in Roster.csv. Skipping them for now. Please add a row for them or an 'Alias' column entry to fix this issue")
            continue
        student_info = roster_dict[student]
        if student_info.course == "Audit":
            continue
        elif course is None:
            course = student_info.course
        elif course != student_info.course:
            println (roster_file_name + " maps students for this class into to multiple courses: " + course + " and " + student_info.course + " - skipping Synergy bulk import formatting")
            return None

        if due_dates is None:
            if due_date_callback is None:
                get_assignment_due_dates_old (course, df.columns[2:])
            else:
                due_dates = get_assignment_due_dates (course, df.columns[2:], due_date_callback)

        for column_name in df.columns[2:]:
            if column_name.strip() == "":       # A blank column is used to separate things that go in Synergy from aggregates of those things
                break                           # We never want to import the "aggregates of aggregates" that follow
            id = student_info.id
            first_name = student_info.first_name
            last_name = student_info.last_name
            assignment_name = column_name
            assignment_description = column_name
            points = row[column_name]
            if not str(points).isdigit():
                println ("Can't parse points for " + assignment_name + " for " + first_name + " " + last_name + " - skipping Synergy bulk import formatting")
                return None
            max_points = df.iloc[0][column_name]
            if not str(max_points).isdigit():
                println ("Can't parse max_points for " + assignment_name + " - skipping Synergy bulk import formatting")
                return None
            overall_score = str(points) + "/" + str(max_points)
            assignment_type = get_assignment_type (course, assignment_name)
            if assignment_type is None:
                println ("Can't parse assignment type for " + assignment_name + " - skipping Synergy bulk import formatting")
                return None
            assignment_date = due_dates[assignment_name]
            if assignment_date in ["S", "X", ""]:
                continue
            output_row = [id, first_name, last_name, assignment_name, assignment_description, overall_score, max_points, assignment_type, assignment_date]
            
            period = student_info.period
            if period not in sdf_dict:
                # Make sure we have a dataframe (initially empty) to hold rows for each period
                sdf = DataFrame()
                for col in ["STUDENT_PERM_ID", "STUDENT_FIRST_NAME", "STUDENT_LAST_NAME", "ASSIGNMENT_NAME", "ASSIGNMENT_DESCRIPTION",
                            "OVERALL_SCORE", "POINTS", "ASSIGNMENT_TYPE", "ASSIGNMENT_DATE"]:
                    sdf[col] = ""
                sdf_dict[period] = sdf
            sdf = sdf_dict[period]
            sdf.loc[len(sdf.index)] = output_row

    output_files = []
    for period in sdf_dict:
        sdf = sdf_dict[period]
        output_file = output_dir + os.path.sep + "Synergy bulk import for P" + str(period) + " " + course + ".xlsx"
        wb = Workbook()
        ws = wb.active
        for r in dataframe_to_rows(sdf, index=False, header=True):
            ws.append(r)
        wb.save(output_file)
        output_files.append(output_file)
    
    return output_files
