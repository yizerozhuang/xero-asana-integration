from utility import convert_to_json

import os
import json
import jsondiff
from datetime import datetime
from config import CONFIGURATION



class AppLog:
    def __init__(self):
        self.conf = CONFIGURATION

    # def log_data(self):
    #     cur_data = convert_to_json(self.app.data)
    #     pre_data = json.load(open(os.path.join(self.database_dir, "data.json")))
    #     diff = jsondiff(pre_data, cur_data, syntax="symmetric")

    def log_create_folder(self, user, quotation):
        self.log_to_file(self.format("Create project from email", user), quotation)

    def log_rename_folder(self, app):
        self.log_to_file(self.format("Rename the Folder", app.user),
                         app.data["Project Info"]["Project"]["Quotation Number"].get())

    def log_fee_proposal(self, app):
        self.log_to_file(self.format(f"Create Fee Proposal Revision {app.data['Fee Proposal']['Reference']['Revision'].get()}", app.user),
                         app.data["Project Info"]["Project"]["Quotation Number"].get())

    def log_email_to_client(self, user, quotation):
        self.log_to_file(self.format("Sent an email to Client", user), quotation)

    def log_chase_client(self, user, quotation):
        self.log_to_file(self.format("Sent a chase email to Client", user), quotation)

    def log_finish_set_up(self, app):
        self.log_to_file(self.format("Finish Set Up", app.user),
                         app.data["Project Info"]["Project"]["Quotation Number"].get())

    def log_unsuccessful(self, app):
        self.log_to_file(self.format("Quote Unsuccessful", app.user),
                         app.data["Project Info"]["Project"]["Quotation Number"].get())

    def log_restore(self, app):
        self.log_to_file(self.format("Restore Project", app.user),
                         app.data["Project Info"]["Project"]["Quotation Number"].get())

    def log_delete(self, app):
        self.log_to_file(self.format("Delete Project", app.user),
                         app.data["Project Info"]["Project"]["Quotation Number"].get())

    def log_update_asana(self, app):
        self.log_to_file(self.format("Update Asana", app.user),
                         app.data["Project Info"]["Project"]["Quotation Number"].get())

    def log_fee_accept_file(self, app):
        self.log_to_file(self.format("log fee accept file", app.user),
                         app.data["Project Info"]["Project"]["Quotation Number"].get())

    def log_generate_invoices(self, app, inv):
        self.log_to_file(self.format(f"Generate Invoice {app.data['Financial Panel']['Invoice Details'][inv]['Number'].get()}", app.user),
                         app.data["Project Info"]["Project"]["Quotation Number"].get())

    def format(self, text, user):
        return f'{datetime.now().strftime("%Y-%m-%d, %H:%M:%S")} {user} {text}'

    def log_to_file(self, text, quotation):
        database_dir = os.path.join(self.conf["database_dir"], quotation)
        if isinstance(text, str):
            text = [text]
        with open(os.path.join(database_dir, "data.log"), "a") as f:
            for t in text:
                f.write(t + '\n')