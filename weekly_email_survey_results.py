#!/usr/bin/python3
from dotenv import load_dotenv
import glob, os
from pyualtrics.qualtrics import Qualtrics
from email_connection import EmailSender
from datetime import date, datetime, timedelta
import pandas as pd
import numpy as np
import csv

load_dotenv()

BASE = ''
AUSTIN_TOKEN = ''
SURVEY_RESPONSE_FOLDER = ''

PAST_X_DAYS = 7
CUSTOMER_SERVICE_SURVEY_ID = ""

EMAIL_FROM = ''
EMAIL_TO = ['']
EMAIL_SMTP_SERVER = ''
EMAIL_SMTP_PORT = 25

q = Qualtrics(qualtricsUrl=BASE, qualtricsToken=AUSTIN_TOKEN, surveyResponseFolder=SURVEY_RESPONSE_FOLDER,
                        skipAPICalls=True, verbose=True)


def create_metrics(dataframe, filename):
    total_responses = 0
    employee_list = {}
    if not dataframe.empty:
        d = {'Employee Email': [], 'Happy': [], 'Neutral': [], 'Unhappy': [], 'Total': []}
        dataframe = dataframe[dataframe['Q2'].notnull()]  # remove empty responses, somehow that happened
        for index, row in dataframe.iterrows():
            if row['employeeEmail']:
                if row['employeeEmail'] not in employee_list.keys():
                    employee_list[row['employeeEmail']] = {
                        'Happy': 0,
                        'Neutral': 0,
                        'Unhappy': 0,
                        'Total': 0
                    }
                if row['Q2'] == 1:
                    employee_list[row['employeeEmail']]['Happy'] +=  1
                    employee_list[row['employeeEmail']]['Total'] +=  1
                elif row['Q2'] == 2:
                    employee_list[row['employeeEmail']]['Neutral'] +=  1
                    employee_list[row['employeeEmail']]['Total'] +=  1
                elif row['Q2'] == 3:
                    employee_list[row['employeeEmail']]['Unhappy'] +=  1
                    employee_list[row['employeeEmail']]['Total'] +=  1
                else:
                    pass
        for email, values in employee_list.items():
            d['Employee Email'].append(email)
            d['Happy'].append(values['Happy'])
            d['Neutral'].append(values['Neutral'])
            d['Unhappy'].append(values['Unhappy'])
            d['Total'].append(values['Total'])
        new_df = pd.DataFrame(data = d)
        print(new_df)
        new_df.to_excel(filename, sheet_name="Summary")


def alphabetize(unsorted_list):
    return sorted(unsorted_list)


def remove_responses_older_than(days: int, dataframe):
    filterDate = str((datetime.now() - timedelta(days=days)).replace(microsecond=0))
    return dataframe.loc[dataframe['EndDate'] > filterDate]

def filter_responses_since(survey, numberDays, quota=0):
    if survey.get_responses():
        df = pd.read_csv(f"{SURVEY_RESPONSE_FOLDER}/{survey.name}.csv", skiprows=[1, 2])
        df = remove_responses_older_than(days=numberDays, dataframe=df)  # remove old responses
        return df
    return None

def get_filtered_results_for_survey(survey_id):
    try:
        survey = q.get_survey(survey_id=survey_id, skipAPICalls=True)
        if survey:
            # if survey.get_responses(): # can be done internally by filter
            return filter_responses_since(survey=survey, numberDays=PAST_X_DAYS)
        return None
    except Exception as e:
        print(f"{e}")
    return None


def process_offline(survey_id,
                    combined_file_path):  # could have done this through pyaultrics library, but whatever
    try:
        df = get_filtered_results_for_survey(survey_id=survey_id)
        if df is not None:
            create_metrics(dataframe=df, filename=combined_file_path)
        return True
    except Exception as e:
        print(f"{e}")
    return False


def main(survey_id, final_file, send_email=False):
    try:
        if process_offline(survey_id=survey_id, combined_file_path=final_file):
            if send_email:
                emailer = EmailSender(smtp_server=EMAIL_SMTP_SERVER, smtp_port=EMAIL_SMTP_PORT)
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
start_day = (date.today() - timedelta(days=PAST_X_DAYS)).strftime('%m-%d-%y')
final_filename = f'Customer Service Email Survey Results - {start_day} to {today}.xls'
# process_offline(reports_folder=SURVEY_RESPONSE_FOLDER, combined_file_path=final_filename)
if main(survey_id=CUSTOMER_SERVICE_SURVEY_ID, final_file=final_filename, send_email=True):
    print("Done.")
else:
    print("Error.")
