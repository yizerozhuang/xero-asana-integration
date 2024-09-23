from config import CONFIGURATION as conf
from asana_function import clean_response

from asana_function import flatter_custom_fields

import asana
import os
import json
from search_bar_page import return_self_mp
from app import App
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
accounting_dir = conf["accounting_dir"]

mp_dir = os.path.join(database_dir, "mp.json")
mp_json = json.load(open(mp_dir))

# scope_dir = os.path.join(database_dir, "scope_of_work.json")
# scope_json = json.load(open(scope_dir))
#
# stage_dir = os.path.join(database_dir, "general_scope_of_staging.json")
# stage_json = json.load(open(stage_dir))
# count=0
for dir in os.listdir(database_dir):
    if os.path.isdir(os.path.join(database_dir, dir)):
        print(dir)
        data_dir = os.path.join(database_dir, dir, "data.json")
        data_json = json.load(open(data_dir))
        data_log_dir = os.path.join(database_dir, dir, "data.log")
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

        # if "Over Due Fee" in data_json['Invoices'].keys():
        #     data_json['Invoices'].pop("Over Due Fee")

        # for service, value in data_json["Bills"]["Details"].items():
        #     if service in ["Mechanical Service", "Mechanical Review"]:
        #         for i in range(conf["n_bills"]):
        #             value["Content"][i]["Type"] = "Mechanical"
        #     elif service == "CFD Service":
        #         for i in range(conf["n_bills"]):
        #             value["Content"][i]["Type"] = "CFD"
        #     elif service == "Electrical Service":
        #         for i in range(conf["n_bills"]):
        #             value["Content"][i]["Type"] = "Electrical"
        #     elif service == "Hydraulic Service":
        #         for i in range(conf["n_bills"]):
        #             value["Content"][i]["Type"] = "Hydraulic"
        #     elif service == "Fire Service":
        #         for i in range(conf["n_bills"]):
        #             value["Content"][i]["Type"] = "Fire"
        #     elif service == "Installation":
        #         for i in range(conf["n_bills"]):
        #             value["Content"][i]["Type"] = "Installation"
        #     elif service in ["Variation", "Miscellaneous"]:
        #         for i in range(conf["n_bills"]):
        #             value["Content"][i]["Type"] = "Others"
        # for i in range(6):
        #     if "Payment Amount" in data_json["Invoices Number"][i].keys():
        #         data_json["Invoices Number"][i].pop("Payment Amount")
        #     if "Payment Date" in data_json["Invoices Number"][i].keys():
        #         data_json["Invoices Number"][i].pop("Payment Date")
        # total_fee = 0
        # total_ingst = 0
        #
        # data_json["Invoices"]["in.GST"] = str(round(float(data_json["Invoices"]["in.GST"]), 2)) if data_json["Invoices"]["in.GST"]!="" else "0"
        # mp_json[data_json["Project Info"]["Project"]["Quotation Number"]]["Total Fee InGST"] = str(round(float(data_json["Invoices"]["in.GST"]), 2)) if data_json["Invoices"]["in.GST"]!="" else "0"
        # for service in conf["invoice_list"]:
        #
        #     service_fee = float(data_json["Invoices"]["Details"][service]["Fee"]) if data_json["Invoices"]["Details"][service]["Fee"] !=""else 0
        #     service_fee_in_gst = float(data_json["Invoices"]["Details"][service]["in.GST"]) if data_json["Invoices"]["Details"][service]["in.GST"] != "" else 0
        #
        #     service_bill = float(data_json["Bills"]["Details"][service]["Fee"]) if data_json["Bills"]["Details"][service]["Fee"] != "" else 0
        #     service_bill_in_gst = float(data_json["Bills"]["Details"][service]["in.GST"]) if data_json["Bills"]["Details"][service]["in.GST"] != "" else 0
        #
        #     profit = round(service_fee - service_bill, 1)
        #     profit_in_gst = round(service_fee_in_gst - service_bill_in_gst, 1)
        #
        #     data_json["Profits"]["Details"][service]["Fee"] = str(profit)
        #     data_json["Profits"]["Details"][service]["in.GST"] = str(profit_in_gst)
        #
        #     total_fee+=profit
        #     total_ingst+=profit_in_gst
        #
        #     for i in range(3):
        #         data_json["Bills"]["Details"][service]["Content"][i]["Contact"] = ""



        # data_json["Profits"]["Fee"] = str(total_fee)
        # data_json["Profits"]["in.GST"] = str(total_ingst)

        for car in data_json["Fee Proposal"]["Calculation Part"]["Car Park"]:
            if "Project" in car.keys():
                car.pop("Project")

        # if len(data_json["Asana_id"]) !=0:
        #     asana_task = flatter_custom_fields(task_api_instance.get_task(data_json["Asana_id"]).to_dict()["data"])
        #     data_json["State"]["Asana State"] = asana_task["Status"]
        # else:
        #     data_json["State"]["Asana State"] = "Fee Proposal"
        #
        #
        # data_log = open(data_log_dir, "r")
        # lines = data_log.readlines()
        #
        # data_json["Email"]["Fee Coming"] = ""
        # data_json["Email"]["Fee Accepted"] = ""
        # for line in lines:
        #     if "log fee accept file" in line:
        #         data_json["Email"]["Fee Accepted"] = line[:10]
        #     elif "Create project from email" in line:
        #         data_json["Email"]["Fee Coming"] = line[:10]
        mp_convert_map = return_self_mp()
        mp_json[dir] = {k: v(data_json) for k, v in mp_convert_map.items()}



        with open(data_dir, "w") as f:
            json_object = json.dumps(data_json, indent=4)
            f.write(json_object)
        # print(count)
        # count+=1
        # if not os.path.exists(os.path.join(database_dir, dir)):
        #     os.makedirs(os.path.join(database_dir, dir))
        # shutil.move(data_dir, os.path.join(database_dir, dir, "data.json"))
        # shutil.move(os.path.join(database_dir, dir, "data.log"), os.path.join(database_dir, dir, "data.log"))

# with open(mp_dir, "w") as f:
#     json_object = json.dumps(mp_json, indent=4)
#     f.write(json_object)
