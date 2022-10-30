"""
GradeAggregator.py - Take exported spreadsheets from our CS teaching
platforms and transform them to a format suitable for import to Synergy

Author: Marc Shepard

For details, please refer to the README.txt file
"""

import GradeUtils
from GradeUtils import println
from TskAggregator import TskAggregator
from StemCspAggregator import StemCspAggregator
from StemCsaAggregator import StemCsaAggregator
from sys import argv, exit
from os import getcwd, path, chdir
from tkinter import scrolledtext
import tkinter as tk
import threading
import traceback
import math

useGui = True   # Use GUI or old command-line interface?

# Double clicking the file or launching from another directory won't work unless we first "cd" to the app directory
if len(argv) >= 1:
    file = argv[0]
    cwd = getcwd()
    dir = path.dirname(file)
    if not path.isabs(dir):
        dir = cwd + "\\" + dir
    chdir(dir)

# A widget for read-only text output
class TextOutput(scrolledtext.ScrolledText):
    def __init__ (self, win):
        pass
        super().__init__(win, wrap = tk.WORD, width = 80, height = 10, padx=10, pady=10)
        self.configure(state=tk.DISABLED)

    def writeln (self, msg : str):
        self.configure(state=tk.NORMAL)
        self.insert("end", msg + "\n")
        self.see("end")
        self.configure(state=tk.DISABLED)

# Pop-up modal diaog to get/modify due dates for Synergy bulk import
text_boxes = {}
dates = {}
pop = None
def on_close ():
    global text_boxes, dates, pop
    for assignment_name, text_box in text_boxes.items():
        due_date = text_box.get("1.0", "end").strip()
        if due_date == "" or GradeUtils.is_date(due_date):
            dates[assignment_name] = due_date
        else:
            println ("Skipping assignment " + assignment_name + " due to malformed due date: " + due_date)
            dates[assignment_name] = ""
    pop.destroy()

def assignment_due_dates_callback (due_dates : dict) -> bool:
    global text_boxes, dates, pop
    text_boxes = {}
    dates = due_dates
    pop = tk.Toplevel()
    frame = tk.Frame(pop)
    tk.Label(frame, text="Enter m/d/y due dates for Synergy bulk import").grid(row=0, columnspan = 2)
    tk.Label(frame, text="Level blank to not bulk import that assignment").grid(row=1, columnspan = 2)
    grid_row = len(due_dates) + 2
    for assignment_name, due_date in due_dates.items():
        if not isinstance(due_date, str) or not GradeUtils.is_date(due_date):
            due_date = ""
        tk.Label(frame, text=assignment_name).grid(row=grid_row, column = 0)
        text_boxes[assignment_name] = tk.Text (frame, height=1, width = 10)
        text_boxes[assignment_name].insert("1.0", due_date)
        text_boxes[assignment_name].grid(row=grid_row, column = 1, sticky = "ew")
        grid_row -= 1
    frame.pack(padx=10, pady=10)

    pop.protocol("WM_DELETE_WINDOW", on_close)
    pop.wait_window(frame)

    return True

# Aggregator wrapper
def aggregate (aggregator):
    println ("\nRunning " + aggregator.name())
    aggregate_already_running = True
    input_file = aggregator.get_default_input_file()
    if input_file is None:
        println ("Error: no files to aggregate with path: \"" + aggregator.get_input_file_pattern () + "\".")
        return
    println ("Processing " + input_file)

    agg_file = GradeUtils.get_output_file_name(input_file, "Aggregated ")
    aggregator.aggregate(input_file, agg_file)
    
    # Launch excel on the output file, so teacher can have a look
    println ("Launching aggregation file \"" + agg_file + "\".")
    GradeUtils.launch_excel (agg_file)

    # If configured, also transform the aggregation in a manner suitable for Synergy bulk import, and show those files as well
    if GradeUtils.synergy_import_configured():
        output_dir = GradeUtils.get_synergy_output_dir(agg_file)
        files = GradeUtils.agg_to_synergy (agg_file, output_dir, assignment_due_dates_callback)
        if files is None or len(files) == 0:
            println ("Failed to create synergy bulk import file from aggregation file")
        else:
            for file in files:
                println ("Launching synergy bulk import file \"" + file + "\".")
                GradeUtils.launch_excel (file)
    else:
        println ("Synergy bulk import aggregation not configured")

# run_aggregator - run one of the aggregators asynchronously so the UI remains responsive
# wrap each run in a try-catch block so we can output any error messages
def async_wrapper (aggregator):
    try:
        aggregate (aggregator)
    except Exception:
        tb = traceback.format_exc()
        println (tb)
def run_aggregator (aggregator):
    threading.Thread(target = async_wrapper, args = [aggregator]).start()

# Button actions for the three aggregators + help
def help_btn_onclick():
    println ("\nSee https://github.com/marcshepard/GradeAggregator/blob/master/README.txt")

def python_btn_onclick():
    run_aggregator(TskAggregator())

def principals_btn_onclick():
    run_aggregator(StemCspAggregator())

def csa_btn_onclick():
    run_aggregator(StemCsaAggregator())

# GUI wrapper
if useGui:
    # Create the window
    window = tk.Tk()
    window.title("Grade aggregator")
    frame = tk.Frame(window, relief = tk.RAISED)

    # First row is a text lable telling folks what to do
    tk.Label(frame, text="Click on the class you wish to aggregate").grid(row=0, column = 0, columnspan=4, sticky="ew")

    # Second row are buttons for aggregator + a help button
    tk.Button(frame, text="Python",     command=python_btn_onclick)    .grid(row=1, column=0, pady = 10)
    tk.Button(frame, text="Principals", command=principals_btn_onclick).grid(row=1, column=1, pady = 10)
    tk.Button(frame, text="Comp Sci A", command=csa_btn_onclick)       .grid(row=1, column=2, pady = 10)
    tk.Button(frame, text="Help",       command=help_btn_onclick)      .grid(row=1, column=3, pady = 10)

    # Last rows are for text output
    text_widget = TextOutput (frame)
    text_widget.grid(row=2, columnspan = 4)
    frame.pack(padx=10, pady=10)
    GradeUtils.print_func = text_widget.writeln

    window.mainloop()
    exit (0)

# These aggregators do the heavy lifting, in a couse specific manner (since the exported spreadsheets are all quite different)
aggregators = [TskAggregator(), StemCspAggregator(), StemCsaAggregator()]

for aggregator in aggregators:
    # Let each aggregator do it's thing
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
    
    # Launch excel on the output file, so teacher can have a look
    GradeUtils.launch_excel (agg_file)

    # If configured, also transform the aggregation in a manner suitable for Synergy bulk import, and show those files as well
    if GradeUtils.synergy_import_configured():
        output_dir = GradeUtils.get_synergy_output_dir(agg_file)
        files = GradeUtils.agg_to_synergy (agg_file, output_dir, None)
        if files is None or len(files) == 0:
            print ("Failed to create synergy import file")
        else:
            for file in files:
                GradeUtils.launch_excel (file)
