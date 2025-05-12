This Python utility automates the end‐to‐end workflow for converting electrochemical cycling data into publication‐quality OriginLab plots and exporting both the processed data and Origin project files. Given one or more Excel workbooks containing a “Record Sheet,” it:

Extracts and concatenates only the specific capacity (“SpeCap/mAh/g”) and voltage columns for each cycle into a consolidated CSV.

Imports that CSV into Origin, applies a user‐chosen graph template, and plots every cycle in a single graph (naming each trace “Cycle 1,” “Cycle 2,” etc.).

Saves the resulting Origin project file alongside the original data, and logs input/output details in an Origin Notes window.

A simple Tkinter GUI lets users batch‐process multiple files, track progress via a status bar, and close the tool when finished.
