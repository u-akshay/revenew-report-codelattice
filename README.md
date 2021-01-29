# Revenew report for codelattice
Akshay U

### Introduction
Its a flask webapp to calculate the revenew of employees and the cost of a project with the info of profit or loss.

### Libraries
  + Flask
  + openpyxl
  + fpdf
  
  <b>Python 3.9.1
  
#### Input example
We have two `.xlsx` files as `Project Details.xlsx` and `Employee Timesheet.xlsx`.<br>
This file have few errors:
  1. Project name have spelling error like spaces like "makemytrip" and "make my trip"
  2. Vinitha and Vinita are used as name to represent same employee in different files
  3. Used "make my trip" instead of "metro business" in "Employee Timesheet.xlsx"

I had change the errors manually inside the code.

#### Output example
The main output is the `pdf` file generating automatically as `Revenew.pdf`. <br>
The uploaded `.xlsx` files are storing in `static` folder as "pd.xlsx" and "et.xlsx".


### Working
Download the entire repository. <br>
Install `requirements.txt` using `pip install -r requirements.txt` or install each libraries using pip.<br>
Run `app.py`.
