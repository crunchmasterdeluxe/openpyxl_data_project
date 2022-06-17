# Overview

The data used to compile these reports is confidential, but will be described here. The data is stored in a mySQL database in tabular format. The data for these reports is pulled from an accounts table (customers, which have a foreign key to employees and offices), an employees table (which has a foreign key to offices), and an offices table (which has a foreign key to employees). 

The data is fetched using Python's mysql-connector library and then put into multiple dataframes and segmented into office-level reports.

[Demo Video](https://youtu.be/ZVm2NQtjN6k)

# Data Analysis Results

Questions:
How many sales did the office have?
How many Conversions (Sold->Installed) did the office have?
Which partitions of employees were responsible for the sales?
How many sales on average did each rep have?
How is the office trending this year?
Which product does the office primarily sell? How has that changed over time?
How do conversion rates compare to other offices?
How is each individual rep converting?
Which reps are responsible for recruiting? 
Who was recruited? 
How are those recruits performing?
How is each rep's production? Sales per day? Interactions per day? 
How does each employee rank across the company?
How many sales is each rep predicted to have next month?
How are managers performing?

# Development Environment

Python, and primarily Plotly, mysql-connector, and openpyxl

# Useful Websites

* [Plotly](https://plotly.com/python/)
* [MySql Connector](https://pypi.org/project/mysql-connector-python/)
* [OpenPyXl](https://openpyxl.readthedocs.io/en/stable/)
