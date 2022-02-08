GradeAggregator.py - take exported spreadsheets from our CS teaching
platforms and transform them to a format suitable for import to Synergy

OVERVIEW
GradeAggregator.py currently works on grading spreadsheets exported from the following learning platforms:
* TechSmartKids (TSK) Python 201/202
* Project STEM AP Comp Sci
* Project STEM AP Comp Sci Principals

These learning platforms have tons of assignment, and it makes no sense to have records for each assignment in Synergy.
Rather, we aggregate them by (unit, category) for STEM (e.g, "5 Exercises"), and by (lesson, category) for TSK (e.g, "4.1 Assignments").
Each TSK assignment is scored as 1 point if it was turned in with no syntax errors and has at least half the expected lines of code and/or correct answers; else 0.
Assignments for all three classes are skipped unless at least 1/4 of students have submitted something. This threshold can be tuned.

GradeAggregator.py looks in the download folder for the latest exported grades spreadsheet from each platform and generates two types aggregates:
1. Manual import format: For viewing and/or manual import to Synergy. Has one row per student and column per assigment
2. Synergy bulk import format: For bulk import to synergy. May have multiple rows per student (one per assignment). A separate file is output
   per period (as required by Synergy), so if the class is taught in multiple periods then multiple bulk import aggregates will be produced.
The aggregate files are output into the same directory as the import files, the download folder,  with either an "Aggregate" or "Synergy Aggregate"
prefix to the file name.

SETUP
To enable manual import format aggregates, you just need to install these Python scripts like so:
1. Sync these scripts to your PC by forking this repo
2. Install Python 3.5 or later. The district version in Software center works, or you can install your own.
3. Install these two additional packages: 
    pip install pandas      - dataframes for data manipulation
    pip install openpyxl    - to parse .xlsx files (only needed for TSK)
Once that's done, you can try it out per the "USAGE" section below

To enable synergy bulk import aggregates, the following additional one-time configuration is required:
1. Create a "Roster.csv" file and place it in the same directory as this script, like so:
    a. Log in to synergy and export a spreadsheet like so:
        * Go to home/seating chart/reports/class roster selecting the [] icon to "open report interface"
        * Under sort/output, select CSV format, click "+add" and add perm ID column
        * Under options, select Perm Id
        * Click "print", wait for it to download
    b. Do the following post-processing on Roster.csv:
        * Remove all columns except "Period", "Course Title", "Student Name", and "Sis Number"
        * Add a column at the end called "Alias"; we'll see how to use this later (in USAGE)
        * Remove rows for all courses that don't need grade aggregation. In particular, any course besides the three currently supported classes
        * Save the file as "Roster.csv" in the same local directory as this GradeAggregator.py script
2. In Synergy, configure your grading category weights for each class:
    a. For TSK and AP Comp Sci A, this tool outputs three built-in Synergy assignment types: assignments, formative, and summative
        i. AP Comp Sci: The recommended weights of these categories for Comp Sci A is 10, 30, 60 (that matches the defaults in STEM)
        ii. Tech Smart Kids: I'm using 50, 20, 30 for TSK, but this weighting is up to you
    d. AP Comp Sci Principals: You need to add two additional Synergy categories of "Projects" and "Performance task".
        i. The recommended weighting to match STEM is 10 (assignments), 10 (performance), 10 (projects), 20 (quizzes), 50 (exams)
3. When you run grade aggregation, you will likely get a warning message like this the first time: "Warning: <student> not found in Roster.csv." There are three possible reasons for this, each with a different solution.
    * If the student enrolled late (after you create Roster.csv), then add a row to Roster.csv for them
    * If the student used a slightly different name in the learning platform, add <student> (exactly as it appears in the warning) to the "Alias" column in the corresonding row for that student in Roster.csv.
    * If the student is not getting graded (e.g., STEM's "Test Student", or someone auditing your class), then make sure they have a row in Roster.csv and the course name is "AUDIT" (all caps)
Once you have run grade aggregation on all your classes without errors, you are ready for the USAGE section.

USAGE
Once configured, the script is pretty easy to run:
1. Export grade spreadsheets from the learning platform
2. Launch GradeAggregator.py. If you have file association set up, you can just click on the file from explorer.
For each class (TSK, CSA, Principals) - this script just skips the class if there is no exported grading spreadsheet for it in the download folder.
If there is an exported spreadsheet, it finds the latest one and compares it's timestamp to the timestamp of the latest aggregation file to determine if
reaggregation is needed. If it is needed, it does the reaggregation and then launches Excel so you can look at the results.
It will also ask you if you want to manually reaggregate anyhow (generally you won't want to - I put this option in to help with development).

Some tips/tricks:
1. It's easier to view the manual excel spreadsheet after doing these steps:
    a. Right-sizing the columns by selecting everything (upper left corner), then home/format/auto-fit column width
    b. Freezing the student name column when you scroll: View/freeze frame/freeze first column
2. If you see "Warning: <student> not found in Roster.csv.", please see SETUP instructions above for what to do
3. When you run grade aggregation, it will also ask you for the due dates of assignments
    * It persists your answer into a courses.csv file, so you only have to answer this once/aggregated assignment
    * Type "X" as your answer if you never want to grade the assignment (e.g., during 2nd semester, you will want to do this for the first semester STEM assignments), or S to temporarily skip (e.g, for assignmens not yet due)
4. When you are ready to do the bulk import to Synergy, here are a few tips:
    a. Read https://synergy.wesdschools.org/Help_USA/synergysismanuals/grade_book_user_guide_secondary.pdf, starting on page 82
    b. Test out import in the Synergy test environment first: https://wa-bsd405.edupoint.com/train. The username/password are the same as regular synergy.
    c. On the bulk import screen, check the right "Upload Import File" options; "add assignments not found in current class", "overwrite existing scores", and "show detailed error messages" are good ones
    d. Note: The "primary key" for the assignment appears to be the assignment name; so you change the due date or re-import after more assignments are done (so a higher max score), those will get updated automatically, those get updated

BACKLOG
* Auto-populate assignment due dates for TSK (rather than asking)
* Add GUI front end
