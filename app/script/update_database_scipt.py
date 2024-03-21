from config import CONFIGURATION as conf
from asana_function import clean_response

import asana
import os
import json
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

        # data_json["Login_user"] = ""
        # data_json["Invoices"]["Paid Fee"] = ""
        # data_json["Fee Proposal"]["Calculation Part"] = {
        #     "Car Park":[
        #         {
        #             "Project": "",
        #             "Car park Level": "",
        #             "Number of Carports": "",
        #             "Level Factor": "0",
        #             "Carport Factor": "0",
        #             "Complex Factor": "0",
        #             "CFD Cost": "0"
        #         } for _ in range(conf["car_park_row"])
        #     ],
        #     "Apt": "",
        #     "Custom Apt": "",
        #     "Area": "",
        #     "Custom Area": ""
        # }
        if len(data_json["Project Info"]["Project"]["Project Number"])!=0:
            address = os.path.join(data_json["Project Info"]["Project"]["Project Number"]+"-"+data_json["Project Info"]["Project"]["Project Name"])
        else:
            address = os.path.join(data_json["Project Info"]["Project"]["Quotation Number"]+"-"+data_json["Project Info"]["Project"]["Project Name"])
        data_json["Current_folder_address"] = address
        print(count)
        count+=1


        with open(data_dir, "w") as f:
            json_object = json.dumps(data_json, indent=4)
            f.write(json_object)

        # if not os.path.exists(os.path.join(database_dir, dir)):
        #     os.makedirs(os.path.join(database_dir, dir))
        # shutil.move(data_dir, os.path.join(database_dir, dir, "data.json"))
        # shutil.move(os.path.join(database_dir, dir, "data.log"), os.path.join(database_dir, dir, "data.log"))

