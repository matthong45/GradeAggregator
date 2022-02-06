GradeAggregator.py - take exported spreadsheets from our CS teaching
platforms and transform them to a format suitable for import to Synergy

This script currently supports importing spreadsheets from the following platforms:
* TechSmartKids (TSK) Python 201/202
* Project STEM AP Comp Sci
* Project STEM AP Comp Sci Principals

These grading platforms have tons of assignment, and it makes no sense to have
records for all of them in Synergy. Rather, we aggregate the scores by (unit, category)
for STEM, and by (lesson, category) for TSK.

This tool automatically looks for the latest exported spreadsheet from each platform and then
auto-generates two types aggregates;
* One for viewing and/or manual import to Synergy that has a row per student and column per assigment
* Another to support Synergy bulk import, each assignment for each student has a separate row
The output aggregate files have the same path and file name as the import files,
except that file name has either "Aggregate" or "Synergy Aggregate" prepended.

All you have to do is use TSK or STEM's "export" feature, and then run this script. It will automatically
find the latest exported file and generate the aggregates. By default it only regenerates the aggregates if
the exported spreadsheet is newer, but will prompt you for options as it runs.

TSK aggregation assigns 1 point to each assignment if it was turned in with no
syntax errors and has at least half the expected lines of code and/or correct answers.

Prereqs for running this script: Python 3.5+ + pandas + openpyxl
Python can be installed from software center, but then you will need to also:
pip install pandas      - dataframes for data manipulation
pip install openpyxl    - to parse .xlsx files (only needed for TSK)

There are some additional steps you need to take if you wish to also export to Synergy mass import format:
1. Create a "Roster.csv" file and place it in the same directory as this script, like so:
    a. Log in to synergy and export a spreadsheet like so:
        * Go to home/seating chart/reports/class roster selecting the [] icon to "open report interface"
        * Under sort/output, select CSV format, click "+add" and add perm ID column
        * Under options, select Perm Id
        * Click "print", wait for it to download
    b. Do the following post-processing:
        * Remove all columns except "Period", "Course Title", "Student Name", and "Sis Number"
        * Remove rows for all courses that don't need grade aggregation
        * Save the file as "Roster.csv" in the same directory as is grade aggregator app
2. Configure your grading category weights in Synergy:
    a. For TSK and AP Comp Sci A, this tool just uses the build in "assignment types" of assignments, formative, and summative - so you don't need to create any additional categories
    b. AP Comp Sci: The recommended weights of these categories for Comp Sci A is 10, 30, 60 (that matches the defaults in STEM)
    c. Tech Smart Kids: I'm using 50, 20, 30 for TSK, but this weighting is up to you
    d. AP Comp Sci Principals: You need to add two additional Synergy categories of "Projects" and "Performance task". The recommended weighting to match STEM is 10 (assignments), 10 (performance), 10 (projects), 20 (quizzes), 50 (exams)
3. When you run grade aggregation, you will likely get a warning message like this: Warning: skipping <some name>. Add an "Alias" column entry to Roster.csv if you with to export their grade
    * There are three possible reasons for this, each with a different solution.
    * If the student enrolled late (after you create Roster.csv), then add a row to Roster.csv for them
    * If the student used a slightly different name in the learning platform, append a column called "Alias" to Roster.csv (if you have not already done so), and add <some name> as it appears above in the corresponding students row
    * If the student is not getting graded (e.g., STEM's "Test Student", or someone auditing your class), then just ignore the warning
4. When you are ready to do the bulk import to Synergy, here are a few tips:
    a. Read https://synergy.wesdschools.org/Help_USA/synergysismanuals/grade_book_user_guide_secondary.pdf, starting on page 82
    b. Test out import in the Synergy test environment first: https://wa-bsd405.edupoint.com/train. The username/password are the same as regular synergy.
    c. On the bulk import screen, check the right "Upload Import File" options; "add assignments not found in current class", "overwrite existing scores", and "show detailed error messages" are good ones
    d. Note: The "primary key" for the assignment appears to be the assignment name; so you change the due date or re-import after more assignments are done (so a higher max score), those will get updated automatically, those get updated

WATCH OUT:
Manually delete any columns you don't want to import to Synergy from the first aggregate - before the Synergy aggregates are created
You need to do this to avoid importing last semesters grades!
You can also remove columns for stuff "in progress" (not finished or fully due), although it's OK to import partial stuff because:
1) Assignments are filtered unless at least 25% of folks have submitted, so that will stop a few eager beavers from making folks look behind
2) Subsequent imports will fix/update

This is my last "must do" fix before I'll consider things to be solid - but it's good enough to use now for initial import

TODO: I've got a number of things I still need to do to make this work well:
* Create a more teacher-friendly way to not aggregate a unit before it is done
* Add per-class option to start a particular student at a given assignment
* Add per-class option to specify the assignment date - autopopulate this for TSK
* Add options for assignment type overrides
