### Date created
June 10th, 2020

### Project Title
Python tools

### Description
Tools to be re-usable in the future, given some additional effort to make them
more generic and capable
  * structured selection of 1 or more choices by users, among several offered
  * table-format printing
  Dec 4th 2023: added Excel utilities into this repository, including a full explanation of how to use 
  Note that an __init__.py file is required in order to be able to import this properly from other code
  What you will require also in another module using this (assuming it sits in a different folder)
      import sys
      sys.path.append('..')
      from python_tools import excel_utilities as exc

### Files used
seek_choices.py
print_in_tab.py
Dec 4th, 2023: excel_utilities with its read me file and an __init__.py

### Credits
Developed as part of the US bikeshare data project which is an assignment for
the Udacity "Programming for Data Science" Nanodegree
Dec 4th, 2023: excel_utilities developed during the Sabca assignment
