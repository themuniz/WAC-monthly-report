---
title: WAC Monthly Report
author: José Muñiz  
date:  12 March 2017
---

# Purpose

This script generates the _WAC Monthly Student Interaction Report_, by parsing the contact log (downloaded from Google Sheets) and outputting a new excel sheet with the records that match any given dates. The script also performs some basic data testing and cleaning.

# Installation

This script is written in the python programming language; please contact José Muñiz to make sure that your workstation is setup with python and the required libraries.

Download the zip file from [GitHub](https://github.com/themuniz/WAC-monthly-report) and expand it in your user directory (i.e., not on the desktop or in your downloads directory). The script directory should be kept _in toto._

Additionally, you should download the contact log (as an excel file) from Google Sheets _after_ the end of the time frame you wish to report. Place this file in the data sub-directory of the script. _N.B.: You should remove the google sheet after the report is created._

# Usage

## Summary

    usage: wac_monthly_report.py [-h] start_date end_date

    Create an excel document with the WAC student interactions for a given period
    of time

    positional arguments:
      start_date  Date (inclusive) of the first interaction: YYYY-MM-DD
      end_date    Date (inclusive) of the last interaction: YYYY-MM-DD

    optional arguments:
      -h, --help  show this help message and exit

## Details

-   Open the windows command-line and move to the script directory
-   Type `python wac_monthly_report.py 2017-03-01 2017-03-31` this will run the script against the contact log in the data directory for the dates starting 2017-03-01 to 2017-03-31
    -   These dates are _inclusive_, and will pull records with a contact date between and _including_ 01 March and 31 March
    -   Once the script has completed running (it may take several minutes, the script will say that the report was written), then remove the contact log from the data directory
-   The script will display messages on screen. Messages marked 'INFO' detail the script's normal operation; messages marked 'WARNING', 'ERROR', or 'CRITICAL' detail problems that should be investigated. Non-info messages are also recorded in the report log for the date that the script is run, so you can check on the issue.
    -   WARNING messages are for data related issues, and will report the worksheet that the error occurred, and a record ID. Please review the issue on the Google sheet, and if changes are made, re-download the file and restart the process.
    -   ERROR and CRITICAL messages are for technical issues. If you can't resolve the issue, please contact José Muñiz for technical support.
-   Once the script is finished, it will display `Writing FILENAME to the output directory`, where FILENAME is the name of the excel file. Send the file to Susan Ko, and remember to remove the contact log from the data directory.
