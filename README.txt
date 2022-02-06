GradeAggregator.py - take exported spreadsheets from our CS teaching
platforms and transform them to a format suitable for import to Synergy

GradeAggregator.py currently works on grading spreadsheets exported from the following learning platforms:
* TechSmartKids (TSK) Python 201/202
* Project STEM AP Comp Sci
* Project STEM AP Comp Sci Principals

These learning platforms have tons of assignment, and it makes no sense to have
records for each assignment in Synergy. Rather, we aggregate them by (unit, category) for STEM,
and by (lesson, category) for TSK.

GradeAggregator.py automatically looks for the latest exported spreadsheet from each platform and then
auto-generates two types aggregates;
1. One for viewing and/or manual import to Synergy that has a row per student and column per assigment
2. Another to support Synergy bulk import, each assignment for each student has a separate row
The output aggregate files have the same path and file name as the import files,
except that file name has either "Aggregate" or "Synergy Aggregate" prepended.

All you have to do is use TSK or STEM's "export" feature, and then run GradeAggregator.py. It will automatically
find the latest exported file and generate the aggregates. By default it only regenerates the aggregates if
the exported spreadsheet is newer, but will prompt you for options as it runs.

TSK aggregation assigns 1 point to each assignment if it was turned in with no
syntax errors and has at least half the expected lines of code and/or correct answers.

All the aggregators skip assignments unless at least 1/4 of students have submitted something.

The prereqs for running GradeAggregator.py: Python 3.5+ + pandas + openpyxl
Python can be installed from software center, but then you will need to also:
pip install pandas      - dataframes for data manipulation
pip install openpyxl    - to parse .xlsx files (only needed for TSK)
Plus of course sync'ing down these files (fork via git)

There are some additional one-time steps if you want to produce a spreadsheet in Synergy mass import format (the second format mentioned above):
1. Create a "Roster.csv" file and place it in the same directory as this script, like so:
    a. Log in to synergy and export a spreadsheet like so:
        * Go to home/seating chart/reports/class roster selecting the [] icon to "open report interface"
        * Under sort/output, select CSV format, click "+add" and add perm ID column
        * Under options, select Perm Id
        * Click "print", wait for it to download
    b. Do the following post-processing on Roster.csv:
        * Remove all columns except "Period", "Course Title", "Student Name", and "Sis Number"
        * Remove rows for all courses that don't need grade aggregation (e.g, Special Topics)
        * Save the file as "Roster.csv" in the same local directory as this GradeAggregator.py script
2. Configure your grading category weights in Synergy:
    a. For TSK and AP Comp Sci A, this tool just uses the build in "assignment types" of assignments, formative, and summative - so you don't need to create any additional categories
    b. AP Comp Sci: The recommended weights of these categories for Comp Sci A is 10, 30, 60 (that matches the defaults in STEM)
    c. Tech Smart Kids: I'm using 50, 20, 30 for TSK, but this weighting is up to you
    d. AP Comp Sci Principals: You need to add two additional Synergy categories of "Projects" and "Performance task". The recommended weighting to match STEM is 10 (assignments), 10 (performance), 10 (projects), 20 (quizzes), 50 (exams)
3. When you run grade aggregation, you will likely get a warning message like this: Warning: skipping <some name>. Add an "Alias" column entry to Roster.csv if you with to export their grade
    * There are three possible reasons for this, each with a different solution.
    * If the student enrolled late (after you create Roster.csv), then add a row to Roster.csv for them
    * If the student used a slightly different name in the learning platform, append a column called "Alias" to Roster.csv (if you have not already done so), and add <some name> as it appears in the above warning in the corresponding students row
    * If the student is not getting graded (e.g., STEM's "Test Student", or someone auditing your class), then just ignore the warning
3. When you run grade aggregation, it will also ask you for the due dates of assignments
    * It persists your answer into a courses.csv file, so you only have to answer this once/aggregated assignment
    * You must manually edit the file to change the value of answers you have previously given
    * Type "skip" as your answer if you never want to grade the assignment. E.g., during 2nd semester, you will want to do this for the first semester STEM assignments
4. When you are ready to do the bulk import to Synergy, here are a few tips:
    a. Read https://synergy.wesdschools.org/Help_USA/synergysismanuals/grade_book_user_guide_secondary.pdf, starting on page 82
    b. Test out import in the Synergy test environment first: https://wa-bsd405.edupoint.com/train. The username/password are the same as regular synergy.
    c. On the bulk import screen, check the right "Upload Import File" options; "add assignments not found in current class", "overwrite existing scores", and "show detailed error messages" are good ones
    d. Note: The "primary key" for the assignment appears to be the assignment name; so you change the due date or re-import after more assignments are done (so a higher max score), those will get updated automatically, those get updated

Backlog: I've got a number of things I still need to do to make this work well:
* Auto-populate assignment due dates for TSK (rather than asking)
* Add GUI front end
