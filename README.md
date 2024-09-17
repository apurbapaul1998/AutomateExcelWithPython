Supermarket Sales Data Automation with Python
Project Overview
This project demonstrates how Python can be used to automate repetitive tasks in Excel for periodic reporting. The main goal is to streamline the process of creating pivot tables, generating bar charts, and formatting sales reports automatically. This is especially useful for industries such as supply chain, sales, and marketing, where periodic reports need to be updated frequently with similar data structures.

Key Features
Automated generation of pivot tables using pandas.
Automated creation of bar charts in Excel using openpyxl.
Automated formatting and calculation of totals in the Excel report.
User input to specify the month for report generation.
Full automation from raw CSV file to a finished Excel report ready for distribution.
Project Structure
The project consists of two main Python scripts:

Pivot Table Creation (pivot_table.py): This script reads a CSV file of supermarket sales data, processes it using pandas, and generates a pivot table summarizing sales based on gender and product line.
Report Automation (pivot_to_report.py): This script automates the process of generating bar charts, adding totals, and formatting the final Excel report using the openpyxl library.
Installation
Prerequisites
Ensure that you have Python installed on your system. The project also requires the following Python libraries:

pandas
openpyxl
You can install the necessary libraries using the following command:

Copy code
pip install pandas openpyxl
Dataset
Download the supermarket sales data from Kaggle and save it in the project directory.

How to Run the Project
1. Pivot Table Creation
Run the pivot_table.py script to create a pivot table from the supermarket sales data. This script reads a CSV file and generates a new Excel file (Pivot_table.xlsx) with a pivot table summarizing total sales by gender and product line.

Copy code
python pivot_table.py
2. Report Generation and Automation
After creating the pivot table, run the pivot_to_report.py script to generate a bar chart, add totals, and format the final report.

This script prompts you to specify the month for the report and generates the final report (Report_<Month>.xlsx) automatically.

Copy code
python pivot_to_report.py
Alternatively, you can pass the month as an argument directly while running the script:

Copy code
python pivot_to_report.py April
Example Output
An example of the output includes:

Pivot Table: A detailed summary of total sales by gender and product line.
Bar Chart: A visual representation of sales by product line.
Formatted Report: A complete, formatted sales report with totals and chart, saved as Report_<Month>.xlsx.
