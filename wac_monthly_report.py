#!/usr/bin/env python
"""
wac_monthly_report.py

Create an excel document with the WAC student interactions for a given period
of time
"""
import argparse
import datetime
import glob
import logging
import os
import sys

import openpyxl
import pandas as pd


def setup(directories):
    """Check if necessary directories exist and create them if needed"""
    logging.info('Starting setup')
    for directory in directories:
        directory = os.path.join('./', directory)
        if not os.path.exists(directory):
            os.makedirs(os.path.join(directory))

    contact_file = glob.glob('./data/*.xlsx')

    if len(contact_file) == 1:
        pass
    elif len(contact_file) == 0:
        logging.critical(
            'WAC contact history file is missing from the data directory')
        sys.exit()
    else:
        logging.critical(
            'Multiple WAC contact history files are in the data directory')
        sys.exit()

    return contact_file[0]


class InteractionData(object):
    """Represents a dataframe containing all interactions found in the contact
    log. InteractionData have the following properties:

    Attributes:
        start_date: A string representing the date of the first (inclusive)
        interaction to be collected
        end_date: A string representing the date of the last (inclusive)
        interaction to be collected
        contact_file: The excel file containing the contact log
        data: A dataframe containing interactions to be reported
        """

    def __init__(self, start_date, end_date, contact_file):
        """Return an InteractionData object"""
        self.data_generated = datetime.date.today()
        self.start_date = start_date
        self.end_date = end_date
        self.contact_file = contact_file
        self.data = pd.DataFrame()

    def collect_data(self):
        """Collect all interactions from the contact log"""
        logging.info('Opening {}'.format(self.contact_file))
        workbook = openpyxl.load_workbook(self.contact_file)
        sheets = workbook.sheetnames
        sheets = [x for x in sheets if ',' in x]
        logging.info('{} student worksheets found in {}'.format(
            len(sheets), self.contact_file))

        logging.info('Looping through student worksheets')
        for sheet in sheets:
            df = pd.read_excel(self.contact_file, sheetname=sheet)
            logging.info('Reading worksheet {}'.format(sheet))
            df['Student Name'] = sheet
            self.data = self.data.append(df, ignore_index=True)

        logging.info('{} interactions found in {}'.format(
            len(self.data), self.contact_file))

    def clean_records(self):
        """Remove records with missing/invalid dates, dates outside of month,
        and select/order columns"""

        logging.info('Starting to clean records')
        # Remove records with missing/invalid dates
        null_dates = self.data[self.data['Contact Date'].isnull()]
        for index, row in null_dates.iterrows():
            logging.warning('Empty contact date in record {}/{}'
                            .format(index, row['Student Name']))
        self.data = self.data[pd.notnull(self.data['Contact Date'])]
        string_pattern = r'[A-Z][a-z]'
        string_dates = self.data[self.data['Contact Date']
                                 .str.match(string_pattern, na=False)]
        for index, row in string_dates.iterrows():
            logging.warning('Invalid information in record {}/{}'
                            .format(index, row['Student Name']))
            logging.info('Removing invalid record')
            self.data = self.data.drop(index, axis='rows')

        # Convert strings to dates
        logging.info('Converting dates')
        self.data['Contact Date'] = pd.to_datetime(
            self.data['Contact Date'], infer_datetime_format=True)

        # Begin text processing
        self.data['Student Name'] = self.data[
            'Student Name'].str.strip().str.title()
        self.data[[
            'Content/Topic of the Exchange', 'Actions and/or Follow up'
        ]] = self.data[[
            'Content/Topic of the Exchange', 'Actions and/or Follow up'
        ]].fillna('')
        self.data['Content/Topic of the Exchange'] = self.data[
            'Content/Topic of the Exchange'].str.strip().str.replace(
                'same as above', '')
        self.data['Assigned to Writing Fellow'] = self.data[
            'Assigned to Writing Fellow'].str.strip().str.title()
        self.data['Correspondence Method'] = self.data[
            'Correspondence Method'].str.title().str.strip()
        # Filter by dates
        logging.info('Selecting dates between {} and {}'.format(
            self.start_date, self.end_date))
        self.data = self.data[(self.data['Contact Date'] >= self.start_date) &
                              (self.data['Contact Date'] <= self.end_date)]

        logging.info(
            'Found {} interactions with {} students between {} and {}'.format(
                len(self.data),
                len(self.data['Student Name'].unique()), self.start_date,
                self.end_date))


class Report(object):
    """Represents a WAC student interaction report.

    Attributes:
        end_date: A string representing the date of the last (inclusive)
        interaction to be collected
        data: A dataframe containing interactions to be reported
    """

    def __init__(self, end_date, data):
        """Return a Report object"""

        self.date_generated = datetime.datetime.today()
        self.end_date = end_date
        self.data = data

    def format_report(self):
        """Remove, rename, and re-order columns and remove time from date"""

        logging.info('Starting to format report')
        logging.info('Dropping duplicate columns')
        drop_cols = [x for x in self.data.columns if '.1' in x]
        self.data = self.data.drop(drop_cols, axis='columns')
        logging.info('Renaming columns')
        self.data = self.data.rename(columns={
            'Assigned to Writing Fellow':
            'Writing Fellow',
            'Correspondence Method':
            'Type of contact',
            'Course (only the abbreviated form, e.g. PSY240)':
            'Course',
            'Student Name':
            'Student',
            'Contact Info':
            'Student Contact Info',
            'Professor (only last name)':
            'Professor',
        })
        self.data['Contact Date'] = self.data['Contact Date'].dt.date
        logging.info('Selecting and ordering columns')
        self.data = self.data[[
            'Contact Date', 'Student', 'Student Contact Info', 'Major',
            'Course', 'Professor', 'Writing Fellow', 'Type of contact',
            'Content/Topic of the Exchange', 'Actions and/or Follow up'
        ]]

    def save_report(self, report_type='excel'):
        """Save data as report_type"""
        report_name = 'wac_monthly_report-{}'.format(self.end_date)
        if report_type == 'json':
            self.data.to_json(
                './output/{}.json'.format(report_name), date_format='iso')
        else:
            self.data.to_excel(
                './output/{}.xlsx'.format(report_name), index_label='ID')
        logging.info('Report written to the output directory')


def main(start_date, end_date):
    """Create WAC student interaction monthly report"""
    start_date = datetime.datetime.strptime(start_date, '%Y-%m-%d').date()
    end_date = datetime.datetime.strptime(end_date, '%Y-%m-%d').date()
    contact_file = setup(['data', 'output'])
    logging.info('Setup complete')

    interactions = InteractionData(
        start_date=start_date, end_date=end_date, contact_file=contact_file)
    interactions.collect_data()
    interactions.clean_records()
    report = Report(end_date=interactions.end_date, data=interactions.data)
    report.format_report()
    report.save_report()


if __name__ == '__main__':
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    log_format = '%(asctime)s - %(levelname)-8s %(message)s'

    s_handler = logging.StreamHandler()
    s_handler.setLevel(logging.INFO)
    s_formatter = logging.Formatter(log_format)
    s_handler.setFormatter(s_formatter)
    logger.addHandler(s_handler)

    f_handler = logging.FileHandler(
        'report_log-{}.txt'.format(datetime.date.today()),
        encoding='utf-8',
        delay='true')
    f_handler.setLevel(logging.WARNING)
    f_formatter = logging.Formatter(log_format)
    f_handler.setFormatter(f_formatter)
    logger.addHandler(f_handler)

    parser = argparse.ArgumentParser(
        description="""Create an excel document with the WAC student interactions
        for a given period of time""")
    parser.add_argument(
        '--start_date',
        default='2017-01-01',
        help='Date (inclusive) of the first interaction: YYYY-MM-DD')
    parser.add_argument(
        '--end_date',
        default='2017-03-01',
        help='Date (inclusive) of the last interaction: YYYY-MM-DD')
    args = parser.parse_args()
    main(args.start_date, args.end_date)
