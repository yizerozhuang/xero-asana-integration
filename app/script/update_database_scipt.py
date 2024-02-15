from config import CONFIGURATION as conf
from asana_function import clean_response

import asana
import os
import json
import time

asana_configuration = asana.Configuration()
asana_configuration.access_token = '2/1203283895754383/1206354773081941:c116d68430be7b2832bf5d7ea2a0a415'
asana_api_client = asana.ApiClient(asana_configuration)
task_api_instance = asana.TasksApi(asana_api_client)

# database_dir = "./database"
database_dir = conf["database_dir"]
scope_dir = os.path.join(database_dir, "scope_of_work.json")
scope_json = json.load(open(scope_dir))

stage_dir = os.path.join(database_dir, "general_scope_of_staging.json")
stage_json = json.load(open(stage_dir))
count=0
for dir in os.listdir(database_dir):
    if os.path.isdir(os.path.join(database_dir, dir)):
        data_dir = os.path.join(database_dir, dir, "data.json")
        data_json = json.load(open(data_dir))

        # if len(data_json["Asana_id"])!=0:
        #     asana_task = clean_response(task_api_instance.get_task(data_json["Asana_id"]))
        #     data_json["Asana_url"] = asana_task["permalink_url"]
        # else:
        #     data_json["Asana_url"] = ""
        # if len(data_json["Project Info"]["Drawing"]) != 12:
        #     for i in range(len(data_json["Project Info"]["Drawing"]), 12):
        #         data_json["Project Info"]["Drawing"].append(
        #             {
        #                 "Drawing Number": "",
        #                 "Drawing Name": "",
        #                 "Revision": ""
        #             }
        #         )
        #
        # data_json["Fee Proposal"]["Installation Reference"] = {
        #     "Date": "",
        #     "Revision": "1",
        #     "Program":None
        # }
        # if len(data_json["Project Info"]["Building Features"]["Major"]["Block"]) != 20:
        #     for i in range(len(data_json["Project Info"]["Building Features"]["Major"]["Block"]), 20):
        #         data_json["Project Info"]["Building Features"]["Major"]["Block"].append([
        #                 "",
        #                 "",
        #                 "",
        #                 "",
        #                 "",
        #                 "",
        #                 "",
        #                 ""
        #             ]
        #         )
        # data_json["Project Info"]["Building Features"]["Major"]["Total Car spot"] = "0"
        # data_json["Project Info"]["Building Features"]["Major"]["Total Others"] = "0"
        # data_json["Project Info"]["Building Features"]["Major"]["Total Apt"] = "0"
        # data_json["Project Info"]["Building Features"]["Major"]["Total Commercial"] = "0"
        # data_json["Project Info"]["Building Features"]["Apt"] = ""
        # data_json["Project Info"]["Building Features"]["Basement"] = ""

        # for key, value in data_json["Fee Proposal"]["Time"].items():
        #     if type(value) is dict:
        #         new = value["Start"] + "-" + value["End"]
        #         data_json["Fee Proposal"]["Time"][key]=new
        # if not "Program" in data_json["Fee Proposal"]["Installation Reference"]:
        #     data_json["Fee Proposal"]["Installation Reference"]["Program"] = None
        # if not "Asana State" in data_json["State"]:
        #     data_json["State"]["Asana State"] = ""
        # data_json["Lock"] = {
        #     "Proposal": False,
        #     "Invoices": False
        # }
        # for service in data_json["Bills"]["Details"].values():
        #     for content in service["Content"]:
        #         if "Description" in content.keys():
        #             content.pop("Description")
        # # print(list(data_json.keys()))
        # design_order_list = ['Asana_id', 'Asana_url', 'Lock', 'State', 'Email', 'Email_Content', 'Address_to', 'Project Info', 'Fee Proposal', 'Invoices', 'Invoices Number', 'Remittances', 'Bills', 'Profits', 'Verbal Acceptance Note', 'Fee_Acceptance_Upload', 'Verbal_Acceptance_Upload']
        # new_dic = {k: data_json[k] for k in design_order_list}
        # data_json = new_dic
        # data_json["Invoices"]["Paid Fee"] = ""
        data_json["Login_user"] = ""
        print(count)
        count+=1


        with open(data_dir, "w") as f:
            json_object = json.dumps(data_json, indent=4)
            f.write(json_object)


