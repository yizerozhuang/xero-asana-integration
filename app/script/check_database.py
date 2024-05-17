from config import CONFIGURATION as conf
from asana_function import clean_response

import asana
import os
import json
from win32com import client as win32client
import time
import shutil

asana_configuration = asana.Configuration()
asana_configuration.access_token = '2/1203283895754383/1206354773081941:c116d68430be7b2832bf5d7ea2a0a415'
asana_api_client = asana.ApiClient(asana_configuration)
task_api_instance = asana.TasksApi(asana_api_client)


working_dir = "P:"
# working_dir = conf["working_dir"]
database_dir = "P:\\app\\database"
# database_dir = conf["database_dir"]
accounting_dir = conf["accounting_dir"]
scope_dir = os.path.join(database_dir, "scope_of_work.json")
scope_json = json.load(open(scope_dir))

stage_dir = os.path.join(database_dir, "general_scope_of_staging.json")
stage_json = json.load(open(stage_dir))
count=0
all_project_asana_id_map = {}
all_invoice_asana_id_map = {}
for dir in os.listdir(database_dir):
    if os.path.isdir(os.path.join(database_dir, dir)):
        data_dir = os.path.join(database_dir, dir, "data.json")
        data_json = json.load(open(data_dir))
        quotation_number = data_json['Project Info']['Project']['Quotation Number']
        print(f"Checking database for {quotation_number}")
        assert data_json["Login_user"]==""
        asana_id = data_json["Asana_id"]
        if asana_id != "":
            if asana_id in all_project_asana_id_map.keys():
                raise Exception(f"Asana ID {asana_id} duplicate with {all_project_asana_id_map[asana_id]}")
            else:
                all_project_asana_id_map[asana_id] = quotation_number
        for inv in data_json["Invoices Number"]:
            if inv["Asana_id"] != "":
                if inv["Asana_id"] in all_invoice_asana_id_map.keys():
                    raise Exception(f"Invoice with invoice number {inv['Number']} has asana_id {inv['Asana_id']} which is duplicate with {all_invoice_asana_id_map[inv['Asana_id']]}")
                else:
                    all_invoice_asana_id_map[inv['Asana_id']] = inv["Number"]


        project_number = data_json["Project Info"]["Project"]["Project Number"]
        first_invoice = data_json["Invoices Number"][0]
        if project_number != first_invoice["Number"] and not first_invoice["State"] in ["Draft", "Voided"]:
            raise Exception(f"Project {quotation_number} with project number {project_number} has first invoice {first_invoice['Number']}")
        current_folder_address = data_json["Current_folder_address"]
        assert os.path.exists(os.path.join(working_dir, current_folder_address))

        project_name = data_json["Project Info"]["Project"]["Project Name"]
        if project_number != "":
            bridge_project_address = project_number+"-"+project_name
        else:
            bridge_project_address = quotation_number+"-"+project_name

        assert bridge_project_address==current_folder_address