GradeAggregator.py - take exported spreadsheets from our CS teaching
platforms and transform them to a format suitable for import to Synergy

This script currently supports importing spreadsheets from the following platforms:
* TechSmartKids (TSK) Python 201/202
* Project STEM AP Comp Sci
* Project STEM AP Comp Sci principals
In fact, this script is just a wrapper for these separate scripts, and adds the
ability to reformat the output in a manner suitable for Synergy bulk import

These grading platforms have tons of assignment, and it makes no sense to have
records for all of them in Synergy. Rather, we aggregate the scores by (unit, category)
for STEM, and by (lesson, category) for TSK.

This tool looks for the latest exported spreadsheet from each platform and then
auto-generates two types aggregates; one for viewing and/or manual import to
Synergy in an easy-to read format, and the other in a harder-to-read format suitable
for Synergy bulk import

The output aggregate files have the same path and file name as the import files,
except that file name has either "Aggregate" or "Synergy Aggregate" prepended. By
default, aggregates are only generated if they don't already exist or are older
than the latest export file.

TSK aggregation assigns 1 point to each assignment if it was turned in with no
syntax errors and has at least half the expected lines of code and/or correct answers.

Prereqs for running this script: Python 3.5+ + pandas + openpyxl
Python can be installed from software center, but then you will need to also:
pip install pandas      - dataframes for data maniupulation
pip install openpyxl    - to parse .xlsx files (only needed for TSK)

A second prereq is <XXX - TODO; add info on how to prepare additional data for Synergy
import once final details are sorted>
