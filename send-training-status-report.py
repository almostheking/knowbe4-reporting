from sys import exit, argv
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
    help = r"""This script sends a training progress CSV report to a recipient of your choice.
    You must provide this script with five parameters.
           Example Usage:
                python send-report2.py -a <API Key> -c CSCP -r somebody@company.com -s me@company.com -p <sender's app password>

            -h : print this help text.
            -e : exclude reporting on training campaigns that have "New Hire" in their names.
            -a : the API key belonging to the KB4 admin account associated with the desired client.
            -c : the client name or acronym to be used when naming the report file. Has no bearing on the targeted KB4 client.
            -r : the recipient that will receive the email with the attached report.
            -s : the address you'll use for sending the report. This should be an organization approved account. The script is written to support O365 and an app password.
            -p : the email sender password, so that the client can authenticate in order to send the report on your behalf. The script is written to support O365 and an app password. Will not support MFA.
    """

    print(help)

def _Fetch_Report (api, exclude_newhire):
    header = {"Authorization": "Bearer "+api}

    campaign_status_resp = requests.get(f"https://us.api.knowbe4.com/v1/training/campaigns",headers=header)

    campaign_data = json.loads(campaign_status_resp.text)

    # cycle through all campaigns and grab ones that are "In Progress"
    active_campaigns = []
    for campaign in campaign_data:
        print(campaign.get('name'))
        if "New Hire" in campaign.get('name') and exclude_newhire == True:
            continue
        elif campaign.get('status') == "In Progress":
            active_campaigns.append(campaign)

    # cycle through "In Progress" camapaigns and gather completion status for each user enrolled
    for campaign in active_campaigns:
        campaign_id = campaign['campaign_id']
        campaign_name = campaign['name']

        enrollments = json.loads(requests.get(f"https://us.api.knowbe4.com/v1/training/enrollments",params={"campaign_id": campaign_id},headers=header).text)

        custom_enrollments = []
        for enrollment in enrollments:
            if "New Hire" in campaign_name and enrollment['status'] == "Passed": # do not report on new hires who have completed their new hire training
                continue
            else:
                enroll = {}
                enroll['name'] = enrollment['user']['first_name'] + " " + enrollment['user']['last_name']
                enroll['email'] = enrollment['user']['email']

                user_id = str(enrollment['user']['id'])
                time.sleep(1)
                get_user = json.loads(requests.get(f"https://us.api.knowbe4.com/v1/users/"+user_id,headers=header).text)
                enroll['manager'] = get_user['manager_name']
                enroll['campaign'] = campaign_name
                enroll['module'] = enrollment['module_name']
                enroll['status'] = enrollment['status']
                custom_enrollments.append(enroll)

        return custom_enrollments

def _Create_CSV (enrollments, client):
    report_name = client+" Training Completion Status.csv"

    # build csv file using enrollments dict generated in _Fetch_Report()
    with open(report_name, "w", newline="") as csv_file:
        columns = ["name","email","manager","campaign","module","status"]
        w = csv.DictWriter(csv_file, fieldnames=columns)
        w.writeheader()
        w.writerows(enrollments)

    return report_name

# construct email to send that contains csv attachment
def _Send_Email (recipient, sender, password, report_name):
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
        opts, args = getopt.getopt(argv,"ea:c:r:s:p:",["api=","client=","recipient=","sender","password"])
    except getopt.GetoptError:
        _Opt_Help()
        exit(2)

    if len(argv) == 0:
        _Opt_Help()
        exit(2)

    exclude_newhire = False
    for opt, arg in opts:
        if opt == "-e":
            exclude_newhire = True
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
        else:
            _Opt_Help()
            exit(2)

    # fetch enrollments via API
    enrollments = _Fetch_Report(api, exclude_newhire)

    # check if enrollments is empty, doesn't send report if so
    if enrollments:
        report_name = _Create_CSV(enrollments, client)
        _Send_Email(recipient, sender, password, report_name)
        os.remove(report_name) # delete report file after sent
    else:
        exit()

main(argv[1:])
