import email
import tkinter.messagebox
from email.header import decode_header
import imaplib
import smtplib

from config import CONFIGURATION as conf

from utility import get_quotation_number, create_new_folder, save, config_state, config_log, save_to_mp
from app_log import AppLog

import os
import json
import time
from datetime import datetime
from PyPDF2 import PdfReader
from win32com import client as win32client
from asana_function import update_asana
import pythoncom
# username = "yee_test@outlook.com"
# password = "Zero0929"
# imap_server = "outlook.office365.com"
username = conf["email_username"]
password = conf["email_password"]
imap_server = conf["imap_server"]
smap_server = conf["smap_server"]
smap_port = conf["smap_port"]

from_addr = "bridge@pcen.com.au"
admin_addr = "admin@pcen.com.au"
# admin_addr = "yee.zhuang@gmail.com"

def sent_email_to_admin(msg_data):
    message = email.message_from_bytes(msg_data[0][1])
    message.replace_header("Subject", "Error when processing Bridge")
    message.replace_header("From", from_addr)
    message.replace_header("To", admin_addr)

    smtp = smtplib.SMTP(smap_server, smap_port)
    smtp.starttls()
    smtp.login(username, password)
    smtp.sendmail(from_addr, admin_addr, message.as_string())
    print("Sent Email to admin")
    smtp.quit()


def check_valid_project_name(project_name):
    working_dir = conf["working_dir"]
    for folder in os.listdir(working_dir):
        if "-" in folder:
            quotation, folder_project_name = folder.split("-", 1)
            current_quotation = datetime.today().strftime("%Y%m")[3:]+"000"
            if project_name == folder_project_name and quotation.startswith(current_quotation):
                return False
    return True

def email_server(app=None):

    log = AppLog()

    allow_email = [
        "Yee Zhuang <yeezhuang@gmail.com>",
        "<felix@pcen.com.au>",
        "Felix YE <felixyeqing@gmail.com>",
        "<admin@pcen.com.au>"
    ]


    imap = imaplib.IMAP4_SSL(imap_server)
    imap.login(username, password)
    imap.select("INBOX")
    pythoncom.CoInitialize()

    while True:
        try:
            status, messages = imap.search(None, "(UNSEEN)")
            messages = messages[0].split(b' ') if len(messages[0]) != 0 else []
        except Exception as e:
            time.sleep(10)
            print(e)
            if "socket" in str(e):
                if not app is None:
                    app.messagebox.show_error("Internet Error, Please close and restart python")
                    app.status_label.config(bg="Red")
                    return
            continue
        if status == "OK":
            for mail in messages:
                res, msg_data = imap.fetch(mail, "(RFC822)")
                for response in msg_data:
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
                        email_content = []
                        print("Subject:", subject.strip())
                        email_content.append(f"Subject: {subject.strip()}")
                        print("From:", From)
                        email_content.append(f"From: {From}")
                        print("Date:", Date)
                        email_content.append(f"Date: {Date}")
                        email_content.append("Content: ")

                        if subject.startswith("Fwd: ") or subject.startswith("FW: "):
                            # if _fee_acceptance(msg):
                            #     database_dir = conf["database_dir"]
                            #     found = False
                            #     for part in msg.walk():
                            #         # extract content type of email
                            #         content_disposition = str(part.get("Content-Disposition"))
                            #         # get the email body
                            #         if "attachment" in content_disposition:
                            #             # download attachment
                            #             filename = part.get_filename()
                            #             if filename:
                            #                 folder_name = os.path.join(conf["working_dir"], "app", "cache")
                            #                 filepath = os.path.join(folder_name, filename)
                            #                 # download attachment and save it
                            #                 open(filepath, "wb").write(part.get_payload(decode=True))
                            #                 parts = []
                            #                 def visitor_body(text, cm, tm, fontDict, fontSize):
                            #                     x = tm[4]
                            #                     y = tm[5]
                            #                     if y > 650 and y < 750 and x > 200:
                            #                         parts.append(text)
                            #                 reader = PdfReader(filepath)
                            #                 page = reader.pages[0]
                            #                 page.extract_text(visitor_text=visitor_body)
                            #                 if "Reference:" in parts or " Reference:" in parts:
                            #                     found = True
                            #                     quotation_number = parts[-1]
                            #                     revision = parts[-3]
                            #                     fee_acceptance_name = f"Fee Acceptance Rev {revision}.pdf"
                            #                     open(os.path.join(database_dir, quotation_number, fee_acceptance_name),
                            #                          "wb").write(part.get_payload(decode=True))
                            #                     app.log.log_fee_accept_file(From.split("<")[-1].split(">")[0],
                            #                                                 quotation_number)
                            #                     if not app is None and app.data["Project Info"]["Project"]["Quotation Number"].get() == quotation_number:
                            #                         app.data["State"]["Fee Accepted"].set(True)
                            #                         save(app)
                            #                         update_asana(app)
                            #                         config_state(app)
                            #                         config_log(app)
                            #                     else:
                            #                         data_json_file_name = os.path.join(database_dir, quotation_number, "data.json")
                            #                         data_json = json.load(open(data_json_file_name))
                            #                         data_json["State"]["Fee Accepted"] = True
                            #                         with open(data_json_file_name, "w") as f:
                            #                             json.dump(data_json, f, indent=4)
                            #                     print("Fee Accepted")
                            #                     break
                            #     if not found:
                            #         message = email.message_from_bytes(msg_data[0][1])
                            #         message.replace_header("Subject", "Error when processing Bridge")
                            #         message.replace_header("From", from_addr)
                            #         message.replace_header("To", to_addr)
                            #
                            #         smtp = smtplib.SMTP(smap_server, smap_port)
                            #         smtp.starttls()
                            #         smtp.login(username, password)
                            #         smtp.sendmail(from_addr, to_addr, message.as_string())
                            #         print("Sent Email to admin")
                            #         smtp.quit()
                            # else:
                                # iterate over email parts
                            try:
                                project_name = abstract_project_name(subject)

                                if not check_valid_project_name(project_name):
                                    continue
                                # all_project_name = get_all_project_name(conf["working_dir"])
                                # if project_name in all_project_name:
                                #     print("project exist, skip ")
                                #     continue
                                current_quotation = get_quotation_number()
                                folder_name = current_quotation + "-" + project_name
                                folder_path = os.path.join(conf["working_dir"], folder_name)
                                create_new_folder(folder_name, conf, current_quotation)

                                data_json = json.load(open(os.path.join(conf["database_dir"], "data_template.json")))
                                data_json["Project Info"]["Project"]["Project Name"] = project_name
                                data_json["Project Info"]["Project"]["Quotation Number"] = current_quotation
                                data_json["Current_folder_address"] = folder_name

                                data_json["Fee Proposal"]["Reference"]["Date"] = datetime.today().strftime("%d-%b-%Y")
                                data_json["Email"]["Fee Coming"] = datetime.today().strftime("%Y-%m-%d")
                                # with open(os.path.join(database_dir, "data.json"), "w") as f:
                                #     json_object = json.dumps(data_json, indent=4)
                                #     f.write(json_object)
                                log.log_create_folder(From.split("<")[-1].split(">")[0], current_quotation)
                                for part in msg.walk():
                                    # extract content type of email
                                    content_type = part.get_content_type()
                                    content_disposition = str(part.get("Content-Disposition"))
                                    # get the email body
                                    if content_type == "text/plain":
                                        email_content += part.get_payload().replace("=20", "").split("\n")
                                    elif "attachment" in content_disposition:
                                        # download attachment
                                        try:

                                            filename = part.get_filename()
                                            if filename:
                                                folder_name = os.path.join(folder_path, "External")
                                                filepath = os.path.join(folder_name, filename)
                                                # download attachment and save it
                                                open(filepath, "wb").write(part.get_payload(decode=True))
                                        except Exception as e:
                                            print(e)
                                data_json["Email_Content"] = "\n".join([e for e in email_content if len(e.replace(" ", "").replace("\r", ""))!=0])

                                with open(os.path.join(conf["database_dir"], current_quotation, "data.json"), "w") as f:
                                    json_object = json.dumps(data_json, indent=4)
                                    f.write(json_object)

                                if not app is None:
                                    config_state(app)
                                    config_log(app)
                                    save_to_mp(app, data_json)
                                f = open(os.path.join(folder_path, "External", "email.eml"), "wb")
                                f.write(response[1])
                                f.close()
                                print("Create New Project")
                            except Exception as e:
                                print(e)
                                # sent_email_to_admin(msg_data)
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
                            try:
                                if not app is None and app.data["Project Info"]["Project"]["Quotation Number"].get() == subject[0:8]:
                                    app.data["State"]["Email to Client"].set(True)
                                    app.data["Email"]["Fee Proposal"].set(datetime.strptime(Date, "%a, %d %b %Y %H:%M:%S %z").strftime("%Y-%m-%d"))
                                    quotation_number = subject[0:8]
                                    app.log.log_email_to_client(From.split("<")[-1].split(">")[0], quotation_number)
                                    save(app)
                                    config_state(app)
                                    config_log(app)
                                else:
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
                                print("Update Email")
                            except Exception as e:
                                print(e)
                                sent_email_to_admin(msg_data)
                        elif subject.startswith("Re: ") or subject.startswith("RE: "):
                            try:
                                quotation_number = subject.split(":", 1)[-1].split("-")[0].strip()
                                if not app is None and app.data["Project Info"]["Project"]["Quotation Number"].get() == quotation_number:
                                    if len(app.data["Email"]["First Chase"].get())==0:
                                        app.data["Email"]["First Chase"].set(datetime.strptime(Date, "%a, %d %b %Y %H:%M:%S %z").strftime("%Y-%m-%d"))
                                    elif len(app.data["Email"]["Second Chase"].get())==0:
                                        app.data["Email"]["Second Chase"].set(datetime.strptime(Date, "%a, %d %b %Y %H:%M:%S %z").strftime("%Y-%m-%d"))
                                    elif len(app.data["Email"]["Third Chase"].get())==0:
                                        app.data["Email"]["Third Chase"].set(datetime.strptime(Date, "%a, %d %b %Y %H:%M:%S %z").strftime("%Y-%m-%d"))
                                    app.log.log_chase_client(From.split("<")[-1].split(">")[0], quotation_number)
                                    save(app)
                                    config_state(app)
                                    config_log(app)
                                else:
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

                                print("Chase Client")
                            except Exception as e:
                                print(e)
                                sent_email_to_admin(msg_data)
                        elif subject.startswith("PCE INV"):
                            try:
                                inv_number = subject.split(" ")[2].split("-")[0]
                                online = False
                                if not app is None:
                                    app_invoice = [value["Number"].get() for value in
                                                   app.data["Invoices Number"] if
                                                   len(value["Number"].get()) != 0]
                                    if inv_number in app_invoice:
                                        online = True
                                if online:
                                    for value in app.data["Invoices Number"]:
                                        if value["Number"].get() == inv_number:
                                            value["State"].set("Sent")
                                            break
                                database_dir = os.path.join(conf["database_dir"], "invoices.json")
                                data_json = json.load(open(database_dir))
                                data_json[inv_number] = "Sent"
                                with open(database_dir, "w") as f:
                                    json.dump(data_json, f, indent=4)
                                print("Invoice Sent")
                            except Exception as e:
                                sent_email_to_admin(msg_data)
                        else:
                            sent_email_to_admin(msg_data)
            print("Process Sleep")
            time.sleep(10)
            print("Process Start")
    imap.close()
    imap.logout()

# def _fee_acceptance(msg):
#     for part in msg.walk():
#         content_type = part.get_content_type()
#         # get the email body
#         if content_type == "text/plain":
#             if "$$$" in part.get_payload():
#                 return True
#     return False


def abstract_project_name(subject):
    project_name = subject.split("Fwd: ")[-1].split("FW: ")[-1].strip()
    special_character_list = ['<', '>', ':', '"', '\\', '/', '|', '?', '*', "\n", "\r"]
    for chr in special_character_list:
        project_name = project_name.replace(chr, "_")
    return project_name


if __name__ == '__main__':
    email_server()

