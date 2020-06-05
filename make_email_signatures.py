#!/usr/bin/python3

from dotenv import load_dotenv
from pyualtrics.qualtrics import Qualtrics
from email_connection import EmailSender
import csv
from datetime import datetime
from datetime import date, datetime, timedelta
import pandas as pd
import numpy as np
import xlwt
from xlwt import Workbook
import os, glob, sys

load_dotenv()

BASE = ''
AUSTIN_TOKEN = ''
SURVEY_RESPONSE_FOLDER = ''

CUSTOMER_SERVICE_SURVEY_ID = ''

EMAIL_FROM = ''
EMAIL_SMTP_SERVER = ''
EMAIL_SMTP_PORT = 25
PASSWORD = ''

SMILEY_EMBED_TEMPLATE = """
<h2>Old signature goes here</h2>
<p></p>
<br>
<table align="center" border="0" cellpadding="0" style="padding:10px;background-color:#f2f4f8;border:1px solid #d9dbde;border-radius:3px;font-family:arial">
   <tbody>
    <tr>
     <td style="padding-bottom:12px;font-size:15px;text-align:center"><img height="75" width="70" src="https://fultoncountyga.co1.qualtrics.com/CP/Graphic.php?IM=IM_2rCaQX3IteDDj7L" style="display:block;margin-left:auto;margin-right:auto"><div dir="ltr"><div title="Download"><div></div></div></div> <span style="font-size:13px"><span style="color:#3498db">How satisfied were you with services you received from {name} today?</span></span></td>
    </tr>
    <tr>
     <td>
      <table border="0" cellpadding="0" cellspacing="5" style="text-align:center" width="100%">
       <tbody>
        <tr>
         <td style="background:#ffffff;border-radius:3px"><a href="https://fultoncountyga.co1.qualtrics.com/jfe/form/{survey_id}?employeeEmail={email}&reaction=1" style="display:block;font-size:12px;text-decoration:none;color:#666;border:12px solid #ffffff;border-radius:3px;background:#ffffff" target="_blank"><img src="https://fultoncountyga.co1.qualtrics.com/CP/Graphic.php?IM=IM_0NvUTeoG7vc8UDP" style="width:50px;height:47px"></a></td>
         <td style="background:#ffffff;border-radius:3px"><a href="https://fultoncountyga.co1.qualtrics.com/jfe/form/{survey_id}?employeeEmail={email}&reaction=2" style="display:block;font-size:12px;text-decoration:none;color:#666;border:12px solid #ffffff;border-radius:3px;background:#ffffff" target="_blank"><img src="https://fultoncountyga.co1.qualtrics.com/CP/Graphic.php?IM=IM_2sjkLG0CdozbOBf" style="width:50px;height:48px"></a></td>
         <td style="background:#ffffff;border-radius:3px"><a href="https://fultoncountyga.co1.qualtrics.com/jfe/form/{survey_id}?employeeEmail={email}&reaction=3" style="display:block;font-size:12px;text-decoration:none;color:#666;border:12px solid #ffffff;border-radius:3px;background:#ffffff" target="_blank"><img src="https://fultoncountyga.co1.qualtrics.com/CP/Graphic.php?IM=IM_3OTLl031Brb6ojb" style="width:50px;height:47px"></a></td>
        </tr>
       </tbody>
      </table> </td>
    </tr>
   </tbody>
  </table>
"""

INSTRUCTIONS = r"""Follow these instructions to set up your Customer Service survey as your email signature (Note: This signature will only be attached to emails sent from your Outlook desktop application; it will not work on emails sent through mobile devices, Fulton County's webmail system or Office 365 online)

VIDEO TUTORIAL: https://youtu.be/icuhRpS4V-8

1. Download the .htm attachment from this email. If your computer asks if you would like to open or save the document, select "Save" and save the file to somewhere on your computer. You will need to remember where you saved this file.

2. In the Outlook application, in the top-left corner, select "File -> Options". An "Outlook Options" window should appear.

3. In the left column of the window, select "Mail".

4. On the right side of the window, locate the "Signatures" button. It should be the third button from the top.

5. While holding down the CTRL key, click the "Signatures" button. A File Explorer window should appear.

6. Navigate to where you saved the .htm file in Step 1. Right-click on the file and select "Copy".

7. Use the back arrow in the top-left area of the screen to return to the folder that opened originally in Step 5. You can also navigate to this folder by typing this address in the File Explorer address bar (do not copy-paste this, or else you will lose the file you copied earlier): C:\Users\<USERNAME HERE>\AppData\Roaming\Microsoft\Signatures

8. Paste the .htm file you copied into the Signatures folder, then close the File Explorer window with the X button in the top-right corner.

9. You should still see the "Outlook Options" window. Without holding the CTRL key this time, click the "Signatures" button.

10. A "Signatures and Stationary" window should appear. Here you can see a list of all your email signatures.

11. In the list in the top-left portion of the screen, select your current email signature. The text should appear in the editing area on the bottom portion of the window.

12. Copy your current email signature, including all text and/or images you would like to keep on your new signature.

13. Select "<YOUR NAME>" from the list of signatures from Step 11. The survey template should appear in the editing area.

14. Replace "Old signature goes here" with your old email signature that you copied in Step 12.

15. Feel free to adjust formatting accordingly. CAREFUL about deleting lines of text; the survey signature will have to be re-imported (back to Step 4) if the formatting is messed up. Remember, you can use CTRL+Z to undo your edits if you make a mistake.

16. Make this your default signature by selecting "<YOUR NAME>" in the "New messages" and "Replies/forwards" sections in the top-right section of the window.

17. Select "OK" in the bottom-right corner of the window when you are done. You can then close the "Outlook Options" window with the X button in the top-right corner of the window.

18. Test your new signature by drafting a new email (or a reply to an existing email).

19. Your new signature should automatically be added to your email. If it is not, locate and click the "Signature" button towards the center of the top bar in your email window (next to "Attach File"). Select "<YOUR NAME>" from the dropdown menu. Your new email signature should appear on your email.

20. Users will now be able to easily select one of the three smiley faces in your email signature, which will briefly open a new browser tab, automatically submit their feedback and then automatically close the browser tab. Responses will be stored on your Qualtrics survey. Contact Justyna Grinholc to gain access to your survey results.

21. Send Justyna Grinholc an email with your new signature when you have completed setup to check that your email signature works properly.
"""

emailer = EmailSender(smtp_server=EMAIL_SMTP_SERVER, smtp_port=EMAIL_SMTP_PORT, username=EMAIL_FROM, password=(PASSWORD if PASSWORD else None))

q = Qualtrics(qualtricsUrl=BASE, qualtricsToken=AUSTIN_TOKEN, surveyResponseFolder=SURVEY_RESPONSE_FOLDER, skipAPICalls=True,
              verbose=True)


def make_embed_code(employee_name, employee_email_address, survey_id):
    try:
        return SMILEY_EMBED_TEMPLATE.format(survey_id=survey_id, email=employee_email_address, name=employee_name)
    except Exception as e:
        print(e)
    return None


def save_embed(embed_html_code, filename):
    try:
        with open(filename, 'w+') as f:
            f.write(embed_html_code)
        return True
    except Exception as e:
        print(e)
    return False


def email_file(filename, to_address, recipient_name):
    return emailer.send_email(from_address=EMAIL_FROM, to_addresses=[to_address],
                              subject=f"Your customer service email signature - {recipient_name}", body=INSTRUCTIONS,
                              attachments_paths=[filename])


def remove_responses_older_than(dataframe, filterDate):
    return dataframe.loc[dataframe['EndDate'] > filterDate]

def filter_responses_since(survey, filter_date, quota=0):
    if survey.get_responses():
        df = pd.read_csv(f"{SURVEY_RESPONSE_FOLDER}/{survey.name}.csv", skiprows=[1, 2])
        df = remove_responses_older_than(dataframe=df, filterDate=filter_date)  # remove old responses
        return df
    return None


def get_filtered_results_for_survey(survey_name):
    try:
        survey = q.get_survey(survey_name=survey_name, skipAPICalls=True)
        if survey:
            if survey.get_responses(): # can be done internally by filter
                last_time = read_from_file("last_time.txt")
                if last_time:
                    return filter_responses_since(survey=survey, filter_date=last_time)
        return None
    except Exception as e:
        print(f"{e}")
    return None

def read_from_file(filename):
    with open(filename, 'r') as f:
        text = f.readline()
    if text:
        return text
    raise Exception(f"{filename} is empty")

def write_to_file(filename, text):
    if text:
        with open(filename, 'w+') as f:
            f.write(text)
    else:
        raise Exception("Couldn't store the new last time!")

def main(q, survey_name):
    test_connection = q.who_am_i(skipAPICalls=True).data
    print(test_connection)
    if test_connection:  # doesn't actually do anything, since Qualtrics will send back JSON even if 200
        results_df = get_filtered_results_for_survey(survey_name=survey_name)
        print(results_df)
        last_time = None
        if results_df is not None:
            past_email_addresses = []
            for index, row in results_df.iterrows():
                email_address = row['Q1']
                last_time = row['EndDate']
                if email_address not in past_email_addresses:
                    first_name = row['Q2']
                    last_name = row['Q3']
                    embed_code = make_embed_code(employee_name=f"{first_name} {last_name}", employee_email_address=email_address, survey_id=CUSTOMER_SERVICE_SURVEY_ID)
                    if embed_code:
                        if save_embed(embed_html_code=embed_code, filename=f"{first_name} {last_name}.htm"):
                            if email_file(filename=f"{first_name} {last_name}.htm", to_address=email_address, recipient_name=f"{first_name} {last_name}"):
                                print(f"Signature created and emailed to {first_name} {last_name} ({email_address})")
                                past_email_addresses.append(email_address)
            write_to_file(filename="last_time.txt", text=last_time)


main(q, 'Get Your Signature!')
