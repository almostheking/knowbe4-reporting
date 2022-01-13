from sys import exit, argv, version
print(version)
from datetime import datetime, timezone
#import dateutil
from dateutil.relativedelta import relativedelta
import os
import time
import getopt
import requests
import json
import csv
import smtplib
import ssl
import mimetypes
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText

def _Opt_Help ():
    help = r"""This script sends a training progress or phish simulation failure CSV report to a recipient of your choice.
    You must provide this script with the six parameters laid out in the top example below.
           Example Usages:
                Generate and send a weekly training report that only includes training status for active campaigns and, optionally, "New Hire" campaigns
                    python send-report2.py -a <API Key> -t wt -c ACME -r recipient@company.com -s sender@company.com -p <sender's app password>
                Generate and send a full KnowBe4 raw data report for the past three months of activity (AKA, the last quarter)
                    python send-report2.py -a <API Key> -t a -f quarter -c ACME -r recipient@company.com -s sender@company.com -p <sender's app password>

            -h : print this help text.
            -e : exclude reporting on training campaigns that have "New Hire" in their names.
            -i : indicates the script will be interactive on the CLI; without this, the script will fail with error messages if missing required variables.
            -a : the API key belonging to the KB4 admin account associated with the desired client.
            -c : the client name or acronym to be used when naming the report file. Has no bearing on the targeted KB4 client.
            -r : the recipient that will receive the email with the attached report.
            -s : the address you'll use for sending the report. This should be an organization approved account. The script is written to support O365 and an app password.
            -p : the email sender password, so that the client can authenticate in order to send the report on your behalf. The script is written to support O365 and an app password. Will not support MFA.
            -t : the type of report; acceptable inputs include "wt" for weekly training report, "t" for training report requiring timeframe, "p" for phishing report requiring timeframe, and "a" for full report requiring timeframe.
            -f : the time period to include in the report(s); acceptable inputs include "year," "last_year," "year_to_date," "quarter," "month," and "week."
    """

    print(help)
    exit(2)

def _Choose_Type ():
    prmpt = r"""Choose a report type
    """

def _Calc_Date (datenow, frequency):
    if frequency == "quarter":
        subtrct = 3
    datepast = datenow - relativedelta(months=subtrct)
    print("datepast: " + str(datepast))
    return datepast

def _Get_T_Campaigns (header):
    print("Getting list of campaigns via KnowBe4 API...")
    campaign_status_resp = requests.get(f"https://us.api.knowbe4.com/v1/training/campaigns",headers=header)
    return json.loads(campaign_status_resp.text)

def _Fetch_WT_Report (header, exclude_newhire):
    custom_enrollments = []# this is the list of report data we'll be returning from this function

    campaign_data = _Get_T_Campaigns(header)

    print("Generating list of In Progress campaigns...")
    active_campaigns = []
    for campaign in campaign_data:# cycle through all campaigns and grab ones that are "In Progress"
        if "New Hire" in campaign.get('name') and exclude_newhire == True:
            print("Skipping New Hire campaign...")
            continue
        elif campaign.get('status') == "In Progress":
            print(campaign['name'] + " is In Progress. Adding to list...")
            active_campaigns.append(campaign)

    print("Getting list of enrollments for each In Progress campaign via KnowBe4 API...")
    for campaign in active_campaigns:# cycle through "In Progress" campaigns and gather completion status for each user enrolled

        campaign_id = campaign['campaign_id']
        enrollments = json.loads(requests.get(f"https://us.api.knowbe4.com/v1/training/enrollments",params={"campaign_id": campaign_id,"per_page": 500},headers=header).text)

        campaign_name = campaign['name']
        for enrollment in enrollments:
            if "New Hire" in campaign_name and enrollment['status'] == "Passed":# do not report on new hires who have completed their new hire training
                continue
            else:
                enroll = {}
                enroll['name'] = enrollment['user']['first_name'] + " " + enrollment['user']['last_name']
                enroll['email'] = enrollment['user']['email']

                user_id = str(enrollment['user']['id'])
                time.sleep(1)
                get_user = json.loads(requests.get(f"https://us.api.knowbe4.com/v1/users/"+user_id,headers=header).text)
                if get_user['status'] == 'archived':
                    continue
                else:
                    enroll['manager'] = get_user['manager_name']
                    enroll['campaign'] = campaign_name
                    enroll['module'] = enrollment['module_name']
                    enroll['status'] = enrollment['status']
                    custom_enrollments.append(enroll)

    return custom_enrollments

def _Fetch_T_Report (header, frequency, datenow):
    custom_enrollments = []# this is the list of report data we'll be returning from this function

    campaign_data = _Get_T_Campaigns(header)

    print("Generating list of campaigns that fit the given timeframe: " + frequency + " ...")
    recent_campaigns = []
    datepast = _Calc_Date(datenow, frequency)


    for campaign in campaign_data:# cycle through all campaigns and grab ones that fit within the given frequency
        print(campaign['name'])
        if campaign['end_date'] == None:
            print("Campaign detected with Null end_date. Probably New Hire!")
            exit()
        if datetime.strptime(campaign['end_date'], "%Y-%m-%dT%H:%M:%S.%f%z") > datepast:# if campaign's end date is newer than the oldest point for which this report is looking...
            print("You got a campaign match, buddy!")
        #
        #
        #
        # if "New Hire" in campaign.get('name') and exclude_newhire == True:
        #     print("Skipping New Hire campaign...")
        #     continue
        # elif campaign.get('status') == "In Progress":
        #     print(campaign['name'] + " is In Progress. Adding to list...")
        #     active_campaigns.append(campaign)
    exit()

def _Create_CSV (enrollments, client):

    print("Generating CSV report to send...")
    report_name = client+" Training Completion Status.csv"

    with open(report_name, "w", newline="") as csv_file:# build csv file using data dict generated in API call function
        columns = ["name","email","manager","campaign","module","status"]
        w = csv.DictWriter(csv_file, fieldnames=columns)
        w.writeheader()
        w.writerows(enrollments)

    return report_name

# construct email to send that contains csv attachment
def _Send_Email (recipient, sender, password, report_name):
    print("Sending CSV report as " + sender + "...")
    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = report_name.split(".")[0]
    body = f"""This report shows the status of current training campaigns broken down by user training completion status."""
    body = MIMEText(body)
    msg.attach(body)
    msg.preamble = report_name

    ctype, encoding = mimetypes.guess_type(report_name)
    if ctype is None or encoding is not None:
        ctype = "application/octet-stream"

    maintype, subtype = ctype.split("/", 1)

    if maintype == "text":
        fp = open(report_name)
        # Note: we should handle calculating the charset
        attachment = MIMEText(fp.read(), _subtype=subtype)
        fp.close()
    elif maintype == "image":
        fp = open(report_name, "rb")
        attachment = MIMEImage(fp.read(), _subtype=subtype)
        fp.close()
    elif maintype == "audio":
        fp = open(report_name, "rb")
        attachment = MIMEAudio(fp.read(), _subtype=subtype)
        fp.close()
    else:
        fp = open(report_name, "rb")
        attachment = MIMEBase(maintype, subtype)
        attachment.set_payload(fp.read())
        fp.close()
        encoders.encode_base64(attachment)
    attachment.add_header("Content-Disposition", "attachment", filename=report_name)
    msg.attach(attachment)

    serv_address = "smtp.office365.com"
    serv_port = 587
    context = ssl.create_default_context()
    with smtplib.SMTP(serv_address, serv_port) as smtp:
        smtp.starttls(context=context)
        smtp.login(sender, password)
        smtp.sendmail(sender, recipient, msg.as_string())
        smtp.quit()

def main (argv):
    try:
        opts, args = getopt.getopt(argv,"eia:c:r:s:p:t:f:",["api=","client=","recipient=","sender","password","type=","frequency="])
    except getopt.GetoptError:
        _Opt_Help()

    if len(argv) == 0:
        _Opt_Help()

    exclude_newhire = False
    interactive = False
    type_list = ["wt", "t", "p", "a"]
    frequency_list = ["year", "last_year", "year_to_date", "quarter", "month", "week"]
    for opt, arg in opts:
        if opt == "-e":
            exclude_newhire = True
        elif opt == "-i":
            interactive = True
        elif opt == "-a":
            api = arg
        elif opt == "-c":
            client = arg
        elif opt == "-r":
            recipient = arg
        elif opt == "-s":
            sender = arg
        elif opt == "-p":
            password = arg
        elif opt == "-t":
            type = arg
        elif opt == "-f":
            frequency = arg
        else:
            _Opt_Help()

    if type in type_list:
        header = {"Authorization": "Bearer "+api}
        if type == "wt":
            print("\nGenerating weekly training report...")
            enrollments = _Fetch_WT_Report(api, header, exclude_newhire)# fetch enrollments via API
            if enrollments:# check if enrollments is empty, doesn't send report if so
                report_name = _Create_CSV(enrollments, client)
                _Send_Email(recipient, sender, password, report_name)
                os.remove(report_name)# delete report file after sent
            else:
                exit()
        else:
            if frequency in frequency_list:
                datenow = datetime.now(timezone.utc)
                if type == "t":
                    print("\nGenerating training report for the following relative timeframe: " + frequency + " ...")
                    enrollments = _Fetch_T_Report(header, frequency, datenow)
                elif type == "p":
                    exit()
                elif type == "a":
                    exit()
            else:
                print("The frequency specified is not compatible.")
                _Opt_Help()
    else:
        print("The type specified is not compatible. Please use the -t option and try again.\n")
        _Opt_Help()

main(argv[1:])
