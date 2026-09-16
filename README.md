# Automated Student Report

Python and Excel automation project using Pandas and OpenPyXL.

## About the Project

This project reads student marks from an Excel file, processes the data using Python, and generates an automated student report.

The program calculates totals, percentages, grades, pass/fail status, and ranks. It also formats the Excel report to make it easier to read.

## Features

- Reads student data from an Excel file
- Cleans column names automatically
- Calculates total marks
- Calculates percentage
- Assigns grades
- Determines Pass/Fail status
- Calculates student ranks
- Identifies the topper
- Sorts students by rank
- Rounds percentages to two decimal places
- Adds borders, filters, and frozen headers
- Highlights the topper and failed students
- Creates a report summary
- Automatically adjusts Excel column widths

## Technologies Used

- Python
- Pandas
- OpenPyXL
- Microsoft Excel / WPS Office

## Input File

The program uses `students.xlsx` containing:

- Name
- Physics
- Chemistry
- Biology

## How It Works

1. The program reads the student data from `students.xlsx`.
2. It cleans the column names.
3. It calculates total marks and percentage.
4. It assigns grades and Pass/Fail status.
5. It calculates ranks and identifies the topper.
6. It sorts the students by rank.
7. It creates and formats the final Excel report.

## How to Run

Install the required Python libraries:

```bash
pip install pandas openpyxl
