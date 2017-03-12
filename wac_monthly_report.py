#!/usr/bin/env python
"""
wac_monthly_report.py

Create an excel document with the WAC interactions for the month
"""
import glob
import logging
import os
import sys
from datetime import date

import openpyxl
import pandas as pd

ALL_DATA = pd.DataFrame()
CONTACT_FILE = glob.glob('./data/*.xlsx')
TODAY = date.today()


def setup(directories):
    """Check if necessary directories exist and create them if needed"""
    logging.info('Starting setup')
    for d in directories:
        d = os.path.join('./', d)
        if not os.path.exists(d):
            os.makedirs(os.path.join(d))

    if len(CONTACT_FILE) == 1:
        pass
    elif len(CONTACT_FILE) == 0:
        logging.critical(
            'WAC contact history file is missing from the data directory')
        sys.exit()
    else:
        logging.critical(
            'Multiple WAC contact history files are in the data directory')
        sys.exit()


def clean_records(data_frame, start_date, end_date):
    """Remove records with missing/invalid dates, dates outside of month,
    and incorrect columns"""
    null_dates = data_frame[data_frame['Contact Date'].isnull()]
    for index, row in null_dates.iterrows():
        logging.warning('Empty contact date in record {}/{}'
                        .format(index, row['Student Name']))
    data_frame = data_frame[pd.notnull(data_frame['Contact Date'])]
    string_pattern = r'[A-Z][a-z]'
    string_dates = data_frame[data_frame['Contact Date']
                              .str.match(string_pattern, na=False)]
    for index, row in string_dates.iterrows():
        logging.warning('Invalid information in record {}/{}'
                        .format(index, row['Student Name']))
        logging.info('Removing invalid record')
        data_frame = data_frame.drop(index, axis='rows')
    logging.info('Converting dates')
    data_frame['Contact Date'] = pd.to_datetime(
        data_frame['Contact Date'], infer_datetime_format=True)
    logging.info(
        'Selecting dates between {} and {}'.format(start_date, end_date))
    data_frame = data_frame[(data_frame['Contact Date'] > start_date) &
                            (data_frame['Contact Date'] < end_date)]
    logging.info(
        'Found {} interactions with {} students between {} and {}'.format(
            len(data_frame),
            len(data_frame['Student Name'].unique()), start_date, end_date))
    return data_frame


def format(data_frame):
    """Remove, rename, and re-order columns"""
    # TODO: remove duplicate and unreported columns
    # TODO: re-order columns
    # TODO: Simple text cleanup
    return data_frame


def main(start_date='2017-01-01', end_date=date.today()):
    """Create WAC student interaction monthly report"""
    setup(['data', 'output'])
    logging.info('Setup complete')

    logging.info('Opening {}'.format(CONTACT_FILE[0]))
    wb = openpyxl.load_workbook(CONTACT_FILE[0])

    sheets = wb.sheetnames
    sheets = [x for x in sheets if ',' in x]

    logging.info('{} student worksheets found in {}'.format(
        len(sheets), CONTACT_FILE[0]))

    for s in sheets:
        df = pd.read_excel(CONTACT_FILE[0], sheetname=s)
        logging.info('Reading worksheet {}'.format(s))
        df['Student Name'] = s
        global ALL_DATA
        ALL_DATA = ALL_DATA.append(df, ignore_index=True)

    logging.info(
        '{} interactions found in {}'.format(len(ALL_DATA), CONTACT_FILE[0]))

    logging.info('Starting to clean records')
    ALL_DATA = clean_records(ALL_DATA, start_date, end_date)
    report_name = 'wac_monthly_report-{}.xlsx'.format(end_date)
    logging.info('Writing {} to the output directory'.format(report_name))
    ALL_DATA.to_excel(os.path.join('./output/', report_name))


if __name__ == '__main__':
    import plac
    # TODO: Add file logger and simplify formatter
    logger = logging.getLogger()
    handler = logging.StreamHandler()
    formatter = logging.Formatter(
        '%(asctime)s %(name)-12s %(levelname)-8s %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    plac.call(main)
