import email
from email.header import decode_header
import imaplib

from config import CONFIGURATION

from utility import get_quotation_number, create_new_folder, save, config_state, config_log
from app_log import AppLog

import os
import json
import time
from datetime import datetime


# username = "yee_test@outlook.com"
# password = "Zero0929"
# imap_server = "outlook.office365.com"


def email_server(app=None):
    username = "bridge@pcen.com.au"
    password = "PcE$yD2023"
    imap_server = "mail.pcen.com.au"

    log = AppLog()

    allow_email = [
        "Yee Zhuang <yeezhuang@gmail.com>",
        "<felix@pcen.com.au>",
        "Felix YE <felixyeqing@gmail.com>",
        "<admin@pcen.com.au>"
    ]
    conf = CONFIGURATION


    imap = imaplib.IMAP4_SSL(imap_server)
    imap.login(username, password)
    imap.select("INBOX")


    while True:
        status, messages = imap.search(None, "(UNSEEN)")
        messages = messages[0].split(b' ') if len(messages[0]) != 0 else []
        if status == "OK":
            for mail in messages:
                res, msg = imap.fetch(mail, "(RFC822)")
                for response in msg:
                    if isinstance(response, tuple):
                        # parse a bytes email into a message object
                        msg = email.message_from_bytes(response[1])
                        # decode the email subject
                        subject, encoding = decode_header(msg["Subject"])[0]
                        if isinstance(subject, bytes):
                            # if it's a bytes, decode to str
                            subject = subject.decode(encoding)
                        # decode email sender
                        From, encoding = decode_header(msg.get("From"))[0]
                        if isinstance(From, bytes):
                            From = From.decode(encoding)

                        Date, encoding = decode_header(msg.get("Date"))[0]
                        if isinstance(Date, bytes):
                            Date = Date.decode(encoding)

                        # if From not in allow_email:
                        #     continue
                        print("Subject:", subject)
                        print("From:", From)
                        print("Date:", Date)
                        if subject.startswith("Fwd: ") or subject.startswith("FW: "):
                            project_name = subject.split("Fwd: ")[-1].split("FW: ")[-1]
                            current_quotation = get_quotation_number()
                            folder_name = current_quotation + "-" + project_name
                            folder_path = os.path.join(conf["working_dir"], folder_name)
                            create_new_folder(folder_name, conf)

                            database_dir = os.path.join(conf["database_dir"], current_quotation)
                            data_json = json.load(open(os.path.join(conf["database_dir"], "data_template.json")))
                            data_json["Project Info"]["Project"]["Project Name"] = project_name
                            data_json["Project Info"]["Project"]["Quotation Number"] = current_quotation
                            data_json["Fee Proposal"]["Reference"]["Date"] = datetime.today().strftime("%d-%b-%Y")
                            os.makedirs(database_dir)
                            with open(os.path.join(database_dir, "data.json"), "w") as f:
                                json_object = json.dumps(data_json, indent=4)
                                f.write(json_object)
                            log.log_create_folder(From.split("<")[-1].split(">")[0], current_quotation)

                            if msg.is_multipart():
                                # iterate over email parts
                                for part in msg.walk():
                                    # extract content type of email
                                    content_type = part.get_content_type()
                                    content_disposition = str(part.get("Content-Disposition"))
                                    # get the email body
                                    try:
                                        body = part.get_payload(decode=True).decode()
                                    except:
                                        pass
                                    if content_type == "text/plain" and "attachment" not in content_disposition:
                                        # print text/plain emails and skip attachments
                                        # print(body)
                                        pass
                                    elif "attachment" in content_disposition:
                                        # download attachment
                                        filename = part.get_filename()
                                        if filename:
                                            folder_name = os.path.join(folder_path, "External")
                                            filepath = os.path.join(folder_name, filename)
                                            # download attachment and save it
                                            open(filepath, "wb").write(part.get_payload(decode=True))
                            else:
                                # extract content type of email
                                content_type = msg.get_content_type()
                                # get the email body
                                body = msg.get_payload(decode=True).decode()
                                if content_type == "text/plain":
                                    # print only text email parts
                                    print(body)
                            print("Create New Project")
                            # if content_type == "text/html":
                            #     # if it's HTML, create a new HTML file and open it in browser
                            #     folder_name = clean(subject)
                            #     if not os.path.isdir(folder_name):
                            #         # make a folder for this email (named after the subject)
                            #         os.mkdir(folder_name)
                            #     filename = "index.html"
                            #     filepath = os.path.join(folder_name, filename)
                            #     # write the file
                            #     open(filepath, "w").write(body)
                            #     # open in the default browser
                            # print("=" * 100)
                        elif subject[0:6].isdigit():
                            if app is None:
                                quotation_number = subject[0:8]
                                database_dir = os.path.join(conf["database_dir"], quotation_number, "data.json")
                                data_json = json.load(open(database_dir))
                                data_json["State"]["Email to Client"] = True
                                data_json["Email"]["Fee Proposal"] = datetime.strptime(Date, "%a, %d %b %Y %H:%M:%S %z").strftime("%Y-%m-%d")
                                data_json["Email"]["First Chase"] = ""
                                data_json["Email"]["Second Chase"] = ""
                                data_json["Email"]["Third Chase"] = ""
                                with open(database_dir, "w") as f:
                                    json.dump(data_json, f, indent=4)
                                log.log_email_to_client(From.split("<")[-1].split(">")[0], quotation_number)
                            else:
                                app.data["State"]["Email to Client"].set(True)
                                app.data["Email"]["Fee Proposal"].set(datetime.strptime(Date, "%a, %d %b %Y %H:%M:%S %z").strftime("%Y-%m-%d"))
                                save(app)
                                config_state(app)
                                config_log(app)
                            print("Update State")
                        elif subject.startswith("Re: "):
                            if app is None:
                                quotation_number = subject[4:12]
                                database_dir = os.path.join(conf["database_dir"], quotation_number, "data.json")
                                data_json = json.load(open(database_dir))
                                if len(data_json["Email"]["First Chase"])==0:
                                    data_json["Email"]["First Chase"] = datetime.strptime(Date, "%a, %d %b %Y %H:%M:%S %z").strftime("%Y-%m-%d")
                                elif len(data_json["Email"]["Second Chase"])==0:
                                    data_json["Email"]["Second Chase"] = datetime.strptime(Date, "%a, %d %b %Y %H:%M:%S %z").strftime("%Y-%m-%d")
                                elif len(data_json["Email"]["Third Chase"])==0:
                                    data_json["Email"]["Third Chase"] = datetime.strptime(Date, "%a, %d %b %Y %H:%M:%S %z").strftime("%Y-%m-%d")
                                with open(database_dir, "w") as f:
                                    json.dump(data_json, f, indent=4)
                                log.log_chase_client(From.split("<")[-1].split(">")[0], quotation_number)
                            else:
                                if len(app.data["Email"]["First Chase"].get())==0:
                                    app.data["Email"]["First Chase"].set(datetime.strptime(Date, "%a, %d %b %Y %H:%M:%S %z").strftime("%Y-%m-%d"))
                                elif len(app.data["Email"]["Second Chase"].get())==0:
                                    app.data["Email"]["Second Chase"].set(datetime.strptime(Date, "%a, %d %b %Y %H:%M:%S %z").strftime("%Y-%m-%d"))
                                elif len(app.data["Email"]["Thrid Chase"].get())==0:
                                    app.data["Email"]["Third Chase"].set(datetime.strptime(Date, "%a, %d %b %Y %H:%M:%S %z").strftime("%Y-%m-%d"))
                                save(app)
                                config_state(app)
                                config_log(app)

                            print("Chase Client")
        print("Process Sleep")
        time.sleep(10)
        print("Process Start")
    imap.close()
    imap.logout()

if __name__ == '__main__':
    email_server()

