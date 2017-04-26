# Metrics_reporting
Generates the security metrics reporting spreadsheet


# Install
Install the xlwings package:
pip install xlwings


# Directory structure
Make sure you create a directory called 'history'.  This will contain the output from previous runs


# Start Excel
xlwings requires Excel to be running


# export your JIRA password
export PASSWORD=xxxxxx


# Run the executable:
python generate_metrics_spreadsheet.py  [month] [year] [JIRAuser] [metrics_reporting.db]

where month is a three letter month in lower case (e.g. jan, feb, etc) and year is a four digit year (e.g. 2017).  JIRAuser is the user you login to JIRA as.  The command above will create the file IPA_metrics_mm_dd_yyyy.csv extracted from JIRA and create also a xlsx file with the same name


# Copy the following files to the 'history' directory:
IPA_metrics_mm_dd_yyyy.xlsx and trending_totals_by_month.csv.  The first file is just for backup puposes but the second one will be updated in each run so it will keep growing with historical data.
