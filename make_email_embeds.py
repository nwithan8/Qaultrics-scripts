#!/usr/bin/python3

from dotenv import load_dotenv
from pyualtrics.qualtrics import Qualtrics
from email_connection import EmailSender
import csv

load_dotenv()

BASE = ''
TOKEN = ''
SURVEY_RESPONSE_FOLDER = 'survey_responses'

EMAIL_FROM = ''
EMAIL_TO = ['']
EMAIL_SMTP_SERVER = ''

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
         <td style="background:#ffffff;border-radius:3px"><a href="https://fultoncountyga.co1.qualtrics.com/jfe/form/{survey_id}?RID=MLRP_egQGW0nFKcN6JbD&amp;Q_CHL=email&amp;Q_PopulateResponse=%7B%22QID4%22:%2211%22%7D&amp;Q_PopulateValidate=1" style="display:block;font-size:12px;text-decoration:none;color:#666;border:12px solid #ffffff;border-radius:3px;background:#ffffff" target="_blank"><img src="https://fultoncountyga.co1.qualtrics.com/CP/Graphic.php?IM=IM_0NvUTeoG7vc8UDP" style="width:50px;height:47px"></a></td>
         <td style="background:#ffffff;border-radius:3px"><a href="https://fultoncountyga.co1.qualtrics.com/jfe/form/{survey_id}?RID=MLRP_egQGW0nFKcN6JbD&amp;Q_CHL=email&amp;Q_PopulateResponse=%7B%22QID4%22:%2212%22%7D&amp;Q_PopulateValidate=1" style="display:block;font-size:12px;text-decoration:none;color:#666;border:12px solid #ffffff;border-radius:3px;background:#ffffff" target="_blank"><img src="https://fultoncountyga.co1.qualtrics.com/CP/Graphic.php?IM=IM_2sjkLG0CdozbOBf" style="width:50px;height:48px"></a></td>
         <td style="background:#ffffff;border-radius:3px"><a href="https://fultoncountyga.co1.qualtrics.com/jfe/form/{survey_id}?RID=MLRP_egQGW0nFKcN6JbD&amp;Q_CHL=email&amp;Q_PopulateResponse=%7B%22QID4%22:%2213%22%7D&amp;Q_PopulateValidate=1" style="display:block;font-size:12px;text-decoration:none;color:#666;border:12px solid #ffffff;border-radius:3px;background:#ffffff" target="_blank"><img src="https://fultoncountyga.co1.qualtrics.com/CP/Graphic.php?IM=IM_3OTLl031Brb6ojb" style="width:50px;height:47px"></a></td>
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

emailer = EmailSender(smtp_server=EMAIL_SMTP_SERVER, smtp_port=25)

q = Qualtrics(qualtricsUrl=BASE, qualtricsToken=TOKEN, surveyResponseFolder=SURVEY_RESPONSE_FOLDER, skipAPICalls=True,
              verbose=True)


def make_smiley_survey(q, template_name, new_name, use_existing: bool = False):
    old_survey = q.get_survey(survey_name=new_name, skipAPICalls=True)
    if old_survey:
        if use_existing:
            print("Existing survey found. Using this one")
            return old_survey
        success = old_survey.delete(skipAPICalls=True)
        if success:
            print("Old survey deleted")
    template_survey = q.get_survey(survey_name=template_name, skipAPICalls=True)
    if template_survey:
        new_survey = template_survey.copy(new_name=new_name, activateNow=True, skipAPICalls=True)
        if new_survey:
            print("New survey created")
            return new_survey
        return None


def make_embed_code(survey):
    try:
        return SMILEY_EMBED_TEMPLATE.format(survey_id=survey.id, name=survey.name)
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


def make_smiley_survey_make_email_html_embed(employee_name, employee_email):
    new_survey = make_smiley_survey(q=q, template_name='Smiley Email Template', new_name=employee_name,
                                    use_existing=True)
    if new_survey:
        embed_code = make_embed_code(new_survey)
        if embed_code:
            if save_embed(embed_html_code=embed_code, filename='{}.htm'.format(employee_name)):
                if email_file(filename='{}.htm'.format(employee_name), to_address=employee_email,
                              recipient_name=employee_name):
                    print("Made survey for {name} and sent embed code to {email}".format(name=employee_name,
                                                                                         email=employee_email))
                else:
                    print("Could not send {file} to {email}".format(file='{}.htm'.format(employee_name),
                                                                    email=employee_email))
            else:
                print("Could not save embed HTML code.")
        else:
            print("Couldn't make embed HTML code.")
    else:
        print("Couldn't make new survey.")


def main(q, employee_list_filename):
    test_connection = q.who_am_i(skipAPICalls=True).data
    print(test_connection)
    if test_connection:  # doesn't actually do anything, since Qualtrics will send back JSON even if 200
        with open(employee_list_filename) as f:
            csv_reader = csv.reader(f, delimiter=',')
            for row in csv_reader:
                print(f'{row[0]} with email {row[1]}')
                make_smiley_survey_make_email_html_embed(employee_name=row[0], employee_email=row[1])


main(q, 'Emails.csv')
