"""
GradeAggregator.py - take exported spreadsheets from our CS teaching
platforms and transform them to a format suitable for import to Synergy

Author: Marc Shepard

For details, please refer to the README.txt file
"""

import GradeUtils
from TskAggregator import TskAggregator
from StemCspAggregator import StemCspAggregator
from StemCsaAggregator import StemCsaAggregator
from sys import argv
from os import getcwd, path, chdir

# Double clicking the file or launching from another directory won't work unless we first "cd" to the app directory
if len(argv) >= 1:
    file = argv[0]
    cwd = getcwd()
    dir = path.dirname(file)
    if not path.isabs(dir):
        dir = cwd + "\\" + dir
    chdir(dir)

aggregators = [TskAggregator(), StemCspAggregator(), StemCsaAggregator()]

for aggregator in aggregators:
    print()
    input_file = aggregator.get_default_input_file()
    aggregator_name = aggregator.name()
    if input_file is None:
        print (aggregator_name + ": Can't find an export file in the default location")
        print ("Skipping aggregation for " + aggregator_name)
        continue
    print (aggregator_name + ": Processing " + input_file)

    agg_file = GradeUtils.get_output_file_name(input_file, "Aggregated ")
    if GradeUtils.is_current (agg_file, input_file):
        print ("Aggregate file is already current: " + agg_file)
        choice = input ("Do you want to reaggregate (r), launch excel (l), or skip (anything else)? ")
        if choice == "r":
            aggregator.aggregate(input_file, agg_file)
        elif choice != "l":
            continue
    else:
        aggregator.aggregate(input_file, agg_file)
    
    GradeUtils.launch_excel (agg_file)
    if GradeUtils.synergy_import_configured():
        output_dir = GradeUtils.get_synergy_output_dir(agg_file)
        files = GradeUtils.agg_to_synergy (agg_file, output_dir)
        if files is None or len(files) == 0:
            print ("Failed to create synergy import file")
        else:
            for file in files:
                GradeUtils.launch_excel (file)
