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

aggregators = [TskAggregator(), StemCspAggregator(), StemCsaAggregator()]

for aggregator in aggregators:
    input_file = aggregator.get_default_input_file()
    print()
    aggregator_name = aggregator.name()
    if input_file is None:
        print (aggregator_name + ": Can't find an export file in the default location")
        print ("Skipping aggregation for " + aggregator_name)
        continue
    print (aggregator_name + ": Processing " + input_file)

    output_file = GradeUtils.get_output_file_name(input_file, "Aggregated ")
    if GradeUtils.is_current (output_file, input_file):
        print ("Aggregate file is already current: " + output_file)
        choice = input ("Do you want to reaggregate (r), launch excel (l), or skip (anything else)? ")
        if choice == "r":
            aggregator.aggregate(input_file, output_file)
        elif choice != "l":
            continue
    else:
        aggregator.aggregate(input_file, output_file)
    
    GradeUtils.launch_excel (output_file)