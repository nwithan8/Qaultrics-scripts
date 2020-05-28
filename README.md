# Qualtrics-scripts
Random Qualtrics scripts

These scripts are highly specialized for internal use in Fulton County (GA) Government. The underlying Qualtrics API library ``pyualtrics`` is available at https://github.com/nwithan8/pyualtrics and can be used to make your own custom scripts

# Installation
1. Clone this repo with ``git clone  https://github.com/nwithan8/Qualtrics-scripts.git``
2. Navigate into the downloaded repo with ``cd Qualtrics-scripts``
3. Install dependencies with ``python3 -m pip install -r requirements.txt``
4. Open each script and complete the missing sensitive information (look for ALL CAPS variables towards the top of each file)
5. Execute a script with ``python3 script_name.py``

# Explanations
## make_email_embeds.py
This script will make a new Qualtrics customer service email survey for each indicated employee and automatically email the generated .htm file to the employee.
Employee names and email addresses should be noted in Emails.csv.

## combine_email_survey_reports.py
This script will collect survey result reports for each employee listed in Get_Reports.csv (technically, it's looking for a survey with the matching title, but if you used ``make_email_embeds.py``, Employee A's customer service survey will be titled 'Employee A'). The reports will then be combined into one .xls spreadsheet file (one sheet for each employee with basic results information), and this file will be automatically emailed to the indicated repicient (TO_EMAIL)
