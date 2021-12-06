********************************************************************************
GradeAggregator.py - take an exported spreadsheets from one of our CS teaching
platforms and transform it to a format suitable for import to Synergy

This tool currently supports importing spreadsheets from the following platforms:
* TechSmartKids (TSK) Python 201/202
* Project STEM AP Comp Sci
* Project STEM AP Comp Sci principals

TSK and aggregates it in a manner suitable for Synergy grading
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

Prereqs for running this script: Python 3.5+ + pandas + openpyxl
Python can be installed from software center, but then you will need to also:
pip install pandas      - dataframes for data maniupulation
pip install openpyxl    - to parse .xlsx files (only needed for TSK)


StemCspAggregator.py - take an exported spreadsheet from Project STEM for the AP Comp Sci Principals class, and
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

StemCspAggregator.py - take an exported spreadsheet from Project STEM for the AP Comp Sci Principals class, and
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
