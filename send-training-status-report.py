from sys import exit, argv
from datetime import datetime, timezone
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
                Generate and send a phishing report for the past three months of activity (AKA, the last quarter)
                    python send-report2.py -a <API Key> -t p -f quarter -c ACME -r recipient@company.com -s sender@company.com -p <sender's app password>

            -h : print this help text.
            -e : exclude reporting on training campaigns that have "New Hire" in their names.
            -a : the API key belonging to the KB4 admin account associated with the desired client.
            -c : the client name or acronym to be used when naming the report file. Has no bearing on the targeted KB4 client.
            -r : the recipient that will receive the email with the attached report.
            -s : the address you'll use for sending the report. This should be an organization approved account. The script is written to support O365 and an app password.
            -p : the email sender password, so that the client can authenticate in order to send the report on your behalf. The script is written to support O365 and an app password. Will not support MFA.
            -t : the type of report; acceptable inputs include "wt" for weekly training report, "t" for training report requiring timeframe, and "p" for phishing report requiring timeframe.
            -f : the time period to include in the report(s); acceptable inputs include "year," "quarter," "month," and "week."
    """
    print(help)
    exit(2)

# def _Choose_Type ():
#     prmpt = r"""Choose a report type
#     """

# calculates the timeframe from which to draw data from the API
def _Calc_Date (datenow, frequency):
    nm = 0# no. of months to subtract
    nd = 0# no. of days to subtract
    if frequency == "quarter":
        nm = 3
    elif frequency == "year":
        nm = 12
    elif frequency == "month":
        nm = 1
    elif frequency == "week":
        nd = 7
    datepast = datenow - relativedelta(months=nm)# this is a datetime object
    datepast-=relativedelta(days=nd)
    return datepast

# get a list of training or phishing campaign ids
def _Get_Campaigns (header, tp):
    print("Getting list of campaigns via KnowBe4 API...")
    if tp == "t":
        campaign_status_resp = requests.get(f"https://us.api.knowbe4.com/v1/training/campaigns",headers=header)
    elif tp == "p":
        campaign_status_resp = requests.get(f"https://us.api.knowbe4.com/v1/phishing/campaigns",params={"campaign_id": "*","per_page": 500},headers=header)
    else:
        print("Could not generate list of campaigns, valid campaign type not given.")
        exit(2)
    return json.loads(campaign_status_resp.text)

# get data for weekly training status report from KB4 API
def _Fetch_WT_Report (header, exclude_newhire):
    custom_enrollments = []# this is the list of report data we'll be returning from this function

    campaign_data = _Get_Campaigns(header, "t")

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

# get data for a training report from KB4 API
def _Fetch_T_Report (header, frequency, datenow):
    custom_enrollments = []# this is the list of report data we'll be returning from this function

    campaign_data = _Get_Campaigns(header, "t")

    print("Generating list of campaigns that fit the given timeframe: " + frequency + " ...")
    recent_campaigns = []
    datepast = _Calc_Date(datenow, frequency)# returns a datetime object that represents the earliest time to track to

    for campaign in campaign_data:# cycle through all campaigns and grab ones that fit within the given frequency
        datecampaign_start = datetime.strptime(campaign['start_date'], "%Y-%m-%dT%H:%M:%S.%f%z")
        if campaign['end_date'] == None and "New Hire" in campaign.get('name') and campaign.get('status') == "In Progress":
            print("An active New Hire campaign has been detected. Adding to list...")
            recent_campaigns.append(campaign)
        elif datecampaign_start > datepast and datecampaign_start < datenow:# if campaign's start date is newer than the oldest point for which this report is looking AND the campaign is not set in the future from now...
            print("The campaign \"" + campaign['name'] + "\" falls within the desired report timeframe. Adding to list...")
            recent_campaigns.append(campaign)
        elif campaign.get('status') == "In Progress":
            print("The campaign \"" + campaign['name'] + "\" does not fall within the desired report timeframe, but it is currently In Progress. Adding to list...")
            recent_campaigns.append(campaign)
        else:
            print("The campaign \"" + campaign['name'] + "\" does not fall within the desired report timeframe. Skipping...")

    print("Getting list of enrollments for each relevant campaign via KnowBe4 API...")
    for campaign in recent_campaigns:# cycle through campaigns and gather completion status for each user enrolled

        campaign_id = campaign['campaign_id']
        enrollments = json.loads(requests.get(f"https://us.api.knowbe4.com/v1/training/enrollments",params={"campaign_id": campaign_id,"per_page": 500},headers=header).text)

        campaign_name = campaign['name']
        for enrollment in enrollments:# reports on all New Hire enrollments, regardless of status, if New Hire is in the recent_campaigns list
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

# get data for a phishing report from KB4 API
def _Fetch_P_Report (header, frequency, datenow):
    custom_recipients = []# this is the list of report data we'll be returning from this function

    campaign_data = _Get_Campaigns(header, "p")

    print("Generating list of campaigns that fit the given timeframe: " + frequency + " ...")
    recent_campaigns = []
    recent_psts = []
    datepast = _Calc_Date(datenow, frequency)# returns a datetime object that represents the earliest time to track to

    for campaign in campaign_data:# cycle through all campaigns and grab ones that fit within the given frequency
        datecampaign_start = datetime.strptime(campaign['last_run'], "%Y-%m-%dT%H:%M:%S.%f%z")
        if datecampaign_start > datepast and datecampaign_start < datenow:# if campaign's start date is newer than the oldest point for which this report is looking AND the campaign is not set in the future from now...
            print("The campaign \"" + campaign['name'] + "\" falls within the desired report timeframe. Adding to list...")
            recent_campaigns.append(campaign)
        elif campaign.get('status') == "In Progress":
            print("The campaign \"" + campaign['name'] + "\" does not fall within the desired report timeframe, but it is currently In Progress. Adding to list...")
            recent_campaigns.append(campaign)
        else:
            print("The campaign \"" + campaign['name'] + "\" does not fall within the desired report timeframe. Skipping...")

    print("Getting list of enrollments for each relevant campaign via KnowBe4 API...")
    for campaign in recent_campaigns:# cycle through campaigns and gather valid phishing tests (PSTs)

        campaign_id = str(campaign['campaign_id'])
        #print(campaign_id)
        url = f"https://us.api.knowbe4.com/v1/phishing/campaigns/" + campaign_id + "/security_tests"
        #print(url)
        psts = json.loads(requests.get(url,params={"per_page": 500},headers=header).text)
        for pst in psts:
            datepst_start = datetime.strptime(pst['started_at'], "%Y-%m-%dT%H:%M:%S.%f%z")
            if datepst_start > datepast and datepst_start < datenow:
                print("The phishing test \"" + str(pst['pst_id']) + "\" falls within the desired report timeframe. Adding to list...")
                recent_psts.append(pst)
            elif pst.get('status') == "Active":
                print("The phishing test \"" + str(pst['pst_id']) + "\" does not fall within the desired report timeframe, but it is currently Active. Adding to list...")
                recent_psts.append(pst)
            else:
                print("The phishing test \"" + str(pst['pst_id']) + "\" does not fall within the desired report timeframe. Skipping...")

        for pst in recent_psts:# cycle through PSTs and gather phishing test status for each enrolled recipient part of the test/campaign
            pst_id = str(pst['pst_id'])
            recipients = json.loads(requests.get(f"https://us.api.knowbe4.com/v1/phishing/security_tests/" + pst_id + "/recipients",params={"per_page": 500},headers=header).text)
            for recipient in recipients:
                recpt = {}
                recpt['name'] = recipient['user']['first_name'] + " " + recipient['user']['last_name']
                recpt['email'] = recipient['user']['email']

                user_id = str(recipient['user']['id'])
                time.sleep(1)# 400 milliseconds
                get_user = json.loads(requests.get(f"https://us.api.knowbe4.com/v1/users/"+user_id,headers=header).text)
                if get_user['status'] == 'archived':
                    continue
                else:
                    recpt['manager'] = get_user['manager_name']
                    recpt['campaign'] = campaign['name']
                    recpt['template_name'] = recipient['template']['name']
                    recpt['delivered_at'] = recipient['delivered_at']
                    recpt['opened_at'] = recipient['opened_at']
                    recpt['clicked_at'] = recipient['clicked_at']
                    recpt['replied_at'] = recipient['replied_at']
                    recpt['attachment_opened_at'] = recipient['attachment_opened_at']
                    recpt['macro_enabled_at'] = recipient['macro_enabled_at']
                    recpt['data_entered_at'] = recipient['data_entered_at']
                    recpt['reported_at'] = recipient['reported_at']
                    custom_recipients.append(recpt)

    return custom_recipients

# def _Fetch_A_Report (header, frequency, datenow):
#     exit()

# pass data into a CSV file based on type of report
def _Create_CSV (data, client, type):

    print("Generating CSV report to send...")
    if type == "t":
        report_name = client+" Training Completion Status.csv"
        columns = ["name","email","manager","campaign","module","status"]
    elif type == "p":
        report_name = client+" Phish Test Status.csv"
        columns = ["name","email","manager","campaign","template_name","delivered_at","opened_at","clicked_at","replied_at","attachment_opened_at","macro_enabled_at","data_entered_at","reported_at"]
    elif type == "a":
        exit()

    with open(report_name, "w", newline="") as csv_file:# build csv file using data dict generated in API call function
        w = csv.DictWriter(csv_file, fieldnames=columns)
        w.writeheader()
        w.writerows(data)

    return report_name

# construct email to send that contains csv attachment
def _Send_Email (recipient, sender, password, report_name):
    print("Sending CSV report as " + sender + " to " + recipient + "...")
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
    #interactive = False
    type_list = ["wt", "t", "p", "a"]
    frequency_list = ["year", "quarter", "month", "week"]
    for opt, arg in opts:
        if opt == "-e":
            exclude_newhire = True
        # elif opt == "-i":
        #     interactive = True
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
            enrollments = _Fetch_WT_Report(header, exclude_newhire)# fetch enrollments via API
            if enrollments:# check if enrollments is empty, doesn't send report if so
                report_name = _Create_CSV(enrollments, client, "t")
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
                    if enrollments:# check if enrollments is empty, doesn't send report if so
                        report_name = _Create_CSV(enrollments, client, type)
                        _Send_Email(recipient, sender, password, report_name)
                        os.remove(report_name)# delete report file after sent
                    else:
                        exit()
                elif type == "p":
                    print("\nGenerating phishing failures report for the following relative timeframe: " + frequency + " ...")
                    recipients = _Fetch_P_Report(header, frequency, datenow)
                    if recipients:
                        report_name = _Create_CSV(recipients, client, type)
                        _Send_Email(recipient, sender, password, report_name)
                        os.remove(report_name)
                    else:
                        exit()
                # elif type == "a":
                #     print("\nGenerating a full KnowBe4 report for the following relative timeframe: " + frequency + " ...")
                #     enrollments =
                #     exit()
            else:
                print("The frequency specified is not compatible.")
                _Opt_Help()
    else:
        print("The type specified is not compatible. Please use the -t option and try again.\n")
        _Opt_Help()

main(argv[1:])
