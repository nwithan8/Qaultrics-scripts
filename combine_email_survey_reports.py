#!/usr/bin/python3
from dotenv import load_dotenv
import glob, os
from pyaultrics.qualtrics import Qualtrics
from email_connection import EmailSender
from datetime import date, datetime, timedelta
import pandas as pd
import numpy as np
import csv
import xlwt
from xlwt import Workbook

load_dotenv()

BASE = ''
TOKEN = ''
SURVEY_RESPONSE_FOLDER = 'email_survey_reports'

HOW_OFTEN_DAYS = 7

EMAIL_FROM = ''
EMAIL_TO = ['']
EMAIL_SMTP_SERVER = ''

q = Qualtrics(qualtricsUrl=BASE, qualtricsToken=TOKEN, surveyResponseFolder=SURVEY_RESPONSE_FOLDER,
                        skipAPICalls=True, verbose=True)


def filter_responses_since(survey, numberDays, quota=0):
    filterDate = str((datetime.now() - timedelta(days=numberDays)).replace(microsecond=0))
    filters = {
        'EndDate': [filterDate, 'after']
    }
    filtered_results = survey.filter_responses_by_date(filters=filters)
    if len(filtered_results) >= quota:
        return filtered_results
    return None


def get_filtered_results_for_survey(survey_name):
    try:
        survey = q.get_survey(survey_name=survey_name, skipAPICalls=True)
        if survey:
            # if survey.get_responses(): # can be done internally by filter
            return filter_responses_since(survey=survey, numberDays=HOW_OFTEN_DAYS)
        return None
    except Exception as e:
        print(f"{e}")
    return None


def get_surveys(survey_names_file):
    try:
        test_connection = q.who_am_i(skipAPICalls=True).data
        print(test_connection)
        if test_connection:  # doesn't actually do anything, since Qualtrics will send back JSON even if 200
            with open(survey_names_file) as f:
                csv_reader = csv.reader(f, delimiter=',')
                temp_list = [row[0] for row in csv_reader]
                temp_list = alphabetize(temp_list)
                for name in temp_list:
                    print(f'Downloading report for {name}...')
                    results = get_filtered_results_for_survey(survey_name=name)
                    if results:
                        print(f'Downloaded {name} report.')
        return True
    except IndexError:  # results downloaded properly, but there's no data to load into a dataframe because no responses
        print(f'Downloaded {name} report.')
        return True
    except Exception as e:
        print(f"{e}")
    return False


def create_metrics(dataframe, sheet):
    total_responses = 0
    res_count = {
        'Pos': 0,
        'Neut': 0,
        'Neg': 0
    }
    if not dataframe.empty:
        # print(dataframe)
        # count response types
        total_responses = dataframe.shape[0]
        for index, row in dataframe.iterrows():
            if row['QID4'] == 11:
                res_count['Pos'] += 1
            elif row['QID4'] == 12:
                res_count['Neut'] += 1
            elif row['QID4'] == 13:
                res_count['Neg'] += 1
    sheet.write(r=0, c=1, label='Response Count')
    sheet.write(1, 0, 'Positive')
    sheet.write(1, 1, res_count['Pos'])
    sheet.write(2, 0, 'Neutral')
    sheet.write(2, 1, res_count['Neut'])
    sheet.write(3, 0, 'Negative')
    sheet.write(3, 1, res_count['Neg'])
    sheet.write(5, 0, 'Total')
    sheet.write(5, 1, total_responses)
    sheet.write(r=0, c=2, label='Response %')
    sheet.write(5, 2, 1)
    if total_responses:
        sheet.write(1, 2, float("{:.2f}".format(res_count['Pos'] / total_responses)))
        sheet.write(2, 2, float("{:.2f}".format(res_count['Neut'] / total_responses)))
        sheet.write(3, 2, float("{:.2f}".format(res_count['Neg'] / total_responses)))
    else:
        sheet.write(1, 2, 0)
        sheet.write(2, 2, 0)
        sheet.write(3, 2, 0)


def alphabetize(unsorted_list):
    return sorted(unsorted_list)


def remove_responses_older_than(days: int, dataframe):
    filterDate = str((datetime.now() - timedelta(days=days)).replace(microsecond=0))
    return dataframe.loc[dataframe['StartDate'] > filterDate]


def process_offline(reports_folder,
                    combined_file_path):  # could have done this through pyaultrics library, but whatever
    try:
        wb = Workbook()
        for file in os.listdir(reports_folder):  # already alphabetized by file system
            if file.endswith('.csv'):
                sheet = wb.add_sheet(f'{os.path.splitext(file)[0]}')
                df = pd.read_csv(f"{reports_folder}/{file}", skiprows=[1, 2])
                df = df[df['QID4'].notnull()]  # remove empty responses, somehow that happened
                df = remove_responses_older_than(days=HOW_OFTEN_DAYS, dataframe=df)  # remove old responses
                create_metrics(dataframe=df, sheet=sheet)
                wb.save(combined_file_path)
        return True
    except Exception as e:
        print(f"{e}")
    return False


def main(survey_names_file, final_file, send_email=False):
    try:
        if get_surveys(survey_names_file=survey_names_file):
            if process_offline(reports_folder=SURVEY_RESPONSE_FOLDER, combined_file_path=final_file):
                if send_email:
                    emailer = EmailSender(smtp_server=EMAIL_SMTP_SERVER)
                    if emailer.send_email(from_address=EMAIL_FROM, to_addresses=EMAIL_TO,
                                          subject=final_file.split(".")[0],
                                          body="Weekly report of customer service email surveys",
                                          attachments_paths=[final_file]):
                        return True
                    return False  # if email failed
                return True
    except Exception as e:
        print(f"{e}")
    return False


today = date.today().strftime('%m-%d-%y')
start_day = (date.today() - timedelta(days=HOW_OFTEN_DAYS)).strftime('%m-%d-%y')
final_filename = f'Customer Service Email Survey Results - {start_day} to {today}.xls'
# process_offline(reports_folder=SURVEY_RESPONSE_FOLDER, combined_file_path=final_filename)
if main(survey_names_file='Get_Reports.csv', final_file=final_filename, send_email=True):
    print("Done.")
else:
    print("Error.")
