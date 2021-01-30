# Revenue Report for codelattice
Akshay U

## Introduction
Its a flask webapp to calculate the revenue of employees and the cost of a project with the info of profit or loss.


## Libraries
  + Flask
  + openpyxl
  + fpdf
  
  <b>Python 3.9.1


### Input example
We have two `.xlsx` files as `Project Details.xlsx` and `Employee Timesheet.xlsx`.<br>
This file have few errors:
  1. Project name have spelling error like spaces like _"makemytrip"_ and _"make my trip"_
  2. _Vinitha_ and _Vinita_ are used as name to represent same employee in different files
  3. Used _"make my trip"_ instead of _"metro business"_ in _"Employee Timesheet.xlsx"_

I had change the errors manually inside the code.


### Output example
The main output is the `pdf` file generating automatically as `Revenue.pdf`. <br>
The uploaded `.xlsx` files are storing in `static` folder as _"pd.xlsx"_ and _"et.xlsx"_.<br>
You can delete those files but delete the folders.


## Method
We have to calculate the revenue for the employee, we have the _Rate/Day_ column in _Project Details.xlsx_ and _No. of hours worked_ column in _Employee Timesheet.xlsx_. We take 8 hour work equal to 1 day. Then we calculated the total days and the days multiplied with Rate/Day, hence we got the total rate for each employee.

We have to find the profit and total cost of each project. We find the total rate of employee for each project and add the _Other Expenses_ to it. So we got the total cost. Then find the difference between _Project Estimation_ and _total cost_ which we calculated.


### Limitations
 + Currency conversion is only for CAD to INR.
 + Columns should be in the order of the sample input files.


## Working
Follow the steps below
  1. Download the entire repository. <br>
  2. Install `requirements.txt` using `pip install -r requirements.txt` or install each libraries using pip.<br>
  3. Run `app.py`.
