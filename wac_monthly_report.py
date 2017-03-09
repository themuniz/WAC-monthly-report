#!/usr/bin/env python
"""
wac_monthly_report.py

Create an excel document with the WAC interactions for the month
"""
import glob
import os
import sys
from datetime import date

import openpyxl
import pandas as pd

ALL_DATA = pd.DataFrame()
CONTACT_FILE = glob.glob('./data/*.xlsx')
TODAY = date.today()


def setup(directories):
    """Check if necessary directories exist, and create them if needed"""
    for d in directories:
        d = os.path.join('./', d)
        if not os.path.exists(d):
            os.makedirs(os.path.join(d))


def clean_records(data_frame):
    """Remove records with missing/invalid dates, dates outside of month,
    and incorrect columns"""
    data_frame = data_frame[pd.notnull(data_frame['Contact Date'])]
    # TODO: check for and remove strings in the 'contact date' column
    data_frame['Contact Date'] = pd.to_datetime(
        data_frame['Contact Date'], infer_datetime_format=True)
    # TODO: remove hardcoding of dates
    start_date = '2017-01-01'
    end_date = TODAY
    data_frame = data_frame[(data_frame['Contact Date'] > start_date) &
                            (data_frame['Contact Date'] < end_date)]
    # TODO: remove excess columns
    return data_frame


setup(['data', 'output'])

if len(CONTACT_FILE) == 1:
    pass
elif len(CONTACT_FILE) == 0:
    print('WAC contact history file is missing from data folder.')
    sys.exit()
else:
    print('Multiple WAC contact history files are in the data folder.')
    sys.exit()

wb = openpyxl.load_workbook(CONTACT_FILE[0])

sheets = wb.sheetnames
sheets = [x for x in sheets if ',' in x]
print('{} students found in total.'.format(len(sheets)))

for s in sheets:
    df = pd.read_excel(CONTACT_FILE[0], sheetname=s)
    df['Student Name'] = s
    ALL_DATA = ALL_DATA.append(df, ignore_index=True)

print('{} interactions.'.format(len(ALL_DATA)))

ALL_DATA.to_excel('./output/wac_monthly_report-{}.xlsx'.format(TODAY))
