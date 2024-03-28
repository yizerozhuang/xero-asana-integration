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

# database_dir = "P:\\app\\database"
working_dir = conf["working_dir"]
database_dir = conf["database_dir"]
# accounting_dir = conf["accounting_dir"]
scope_dir = os.path.join(database_dir, "scope_of_work.json")
scope_json = json.load(open(scope_dir))

stage_dir = os.path.join(database_dir, "general_scope_of_staging.json")
stage_json = json.load(open(stage_dir))
count=0
for dir in os.listdir(database_dir):
    if os.path.isdir(os.path.join(database_dir, dir)):
        data_dir = os.path.join(database_dir, dir, "data.json")
        data_json = json.load(open(data_dir))

        # data_json["Invoices"]["Overdue Fee"] = data_json["Invoices"].pop("Over Due Fee")
        #
        # for service in conf["invoice_list"]:
        #     if not data_json["Invoices"]["Details"][service]["Expand"] and len(data_json["Invoices"]["Details"][service]["Fee"])!=0:
        #         data_json["Invoices"]["Details"][service]["Content"][0]["Service"] = f"{service} Design and Documentation Full Amount"
        #         data_json["Invoices"]["Details"][service]["Content"][0]["Fee"] = data_json["Invoices"]["Details"][service]["Fee"]
        #         data_json["Invoices"]["Details"][service]["Content"][0]["in.GST"] = data_json["Invoices"]["Details"][service]["in.GST"]
        #         data_json["Invoices"]["Details"][service]["Content"][0]["Number"] = data_json["Invoices"]["Details"][service]["Number"]
        #         for i in range(1, 4):
        #             data_json["Invoices"]["Details"][service]["Content"][i]["Service"] = ""
        #             data_json["Invoices"]["Details"][service]["Content"][i]["Fee"] = ""
        #             data_json["Invoices"]["Details"][service]["Content"][i]["in.GST"] = ""
        #             data_json["Invoices"]["Details"][service]["Content"][i]["Number"] = "None"
        #     data_json["Invoices"]["Details"][service].pop("Expand")
        #     data_json["Invoices"]["Details"][service].pop("Number")

        # folder_dir = data_json["Current_folder_address"]
        # if not os.path.exists(os.path.join(conf["working_dir"], folder_dir)):
        #     print()
        # shortcut_dir = os.path.join(conf["working_dir"], folder_dir, "Database Shortcut.lnk")
        # shortcut_working_dir = conf["accounting_dir"]
        # shell = win32client.Dispatch("WScript.Shell")
        # shortcut = shell.CreateShortcut(shortcut_dir)
        # accounting_dir = os.path.join(conf["accounting_dir"],  data_json["Project Info"]["Project"]["Quotation Number"])
        # if not os.path.exists(accounting_dir):
        #     print()
        # shortcut.Targetpath = accounting_dir
        # shortcut.WorkingDirectory = shortcut_working_dir
        # # shortcut.IconLocation = os.path.join(conf["resource_dir"], "jpg", "logo.jpg")
        # shortcut.save()


        with open(data_dir, "w") as f:
            json_object = json.dumps(data_json, indent=4)
            f.write(json_object)
        print(count)
        count+=1
        # if not os.path.exists(os.path.join(database_dir, dir)):
        #     os.makedirs(os.path.join(database_dir, dir))
        # shutil.move(data_dir, os.path.join(database_dir, dir, "data.json"))
        # shutil.move(os.path.join(database_dir, dir, "data.log"), os.path.join(database_dir, dir, "data.log"))

