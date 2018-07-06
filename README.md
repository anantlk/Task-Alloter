# Task Alloter

## Description

This script automatically assigns time slots to members for specific tasks.
It also selects most suitable person for that time slot according to there distance 
from venue of task.

## Usage

1.Run the script:
 
    python desk_duty.py

2.Reads the time table for the day and duty is assigned to them

3.Maximum 2 duties for a person per day

4.input.xlsx shows the format of entering the duties which the document will parse and take as input

5.Name it as "time_table.xlsx" and place it in just outside the main directory of our script.The sheet in the workbook will be named as "Sheet1".

6.The duties will be generated in format of output.xls in the directory of our script by name of "desk_duty.xls"

7.Assigns 2 person with duties in each SJT and TT(venue can be selected according to the choice of the duty generator)

## NEW FEATURES

1.TAKES CARE OF YOUR CONVENIENCE REGARDING TRAVELLING FROM ONE BUILDING TO ANOTHER

2.TRIES TO DISTRIBUTE WORK AS MUCH AS POSSIBLE
