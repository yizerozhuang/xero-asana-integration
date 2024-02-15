import asana
from config import CONFIGURATION as conf
from app_log import AppLog

from win32com import client as win32client
import os
import json
import time


asana_configuration = asana.Configuration()
asana_configuration.access_token = '2/1203283895754383/1206354773081941:c116d68430be7b2832bf5d7ea2a0a415'
asana_api_client = asana.ApiClient(asana_configuration)
project_api_instance = asana.ProjectsApi(asana_api_client)
task_api_instance = asana.TasksApi(asana_api_client)
stories_api_instance = asana.StoriesApi(asana_api_client)
workspace_gid = '1198726743417674'
database_dir = conf["database_dir"]

inv_dir = os.path.join(conf["database_dir"], "invoices.json")
inv_json = json.load(open(inv_dir))

log=AppLog()

MP_project_gid = "1203405141297991"
project_types = ["Restaurant", "Office", "Commercial", "Group House", "Apartment", "Mixed-use complex", "School", "Others"]
excel = win32client.Dispatch("Excel.Application")


def check_if_quotation_exist(quotation_number):
    return quotation_number in [dir for dir in os.listdir(database_dir) if dir[0].isdigit()]

off_set = None
i = 0
exist_task = 0
unprocessed_task_list = []
quotation_number_not_accepted_list = []
task_with_more_than_6_invoice_list = []
start = time.time()
while True:
    if off_set is None:
        ori_tasks = task_api_instance.get_tasks_for_project(project_gid=MP_project_gid, limit=100)
    else:
        ori_tasks = task_api_instance.get_tasks_for_project(project_gid=MP_project_gid, limit=100, offset=off_set)
    tasks_list = ori_tasks.to_dict()["data"]
    for task in tasks_list:
        i+=1
        print(f"Processing Task {i}, the task gid is {task['gid']}")
        if task["name"].startswith("P:\\Archived"):
            print("Archived Tasks, Skipped")
            continue

        if task["name"][9] == "-":
            project_number = task["name"][3:9]
            stories = stories_api_instance.get_stories_for_task(task["gid"]).to_dict()["data"]
            for story in stories:
                text = story["text"]
                if "changed the name to" in text:
                    quotation_number = text.split('"')[1].split("-")[0].split("P:\\")[-1].strip()
                    break
            if quotation_number is None or type(quotation_number) == float:
                quotation_number = project_number
            if quotation_number.isdigit() or i == 366:
                folder_dir = task["name"]
                excel_dir = None
                for root, dirs, files in os.walk(folder_dir):
                    for file in files:
                        if "Fee Proposal" in str(file) and file.endswith(".xlsx"):
                            excel_dir = os.path.join(root, file)
                            break
                if not excel_dir is None:
                    work_book = excel.Workbooks.Open(excel_dir)
                    work_sheets = work_book.Worksheets[0]
                    quotation_number = work_sheets.Cells(3, 2).Value
                    work_book.Close(False)
                    print(f"Found Quotation Number in {excel_dir} with quotation number {quotation_number}")
                    if quotation_number is None or str(quotation_number).isdigit() or type(quotation_number)==float:
                        quotation_number_not_accepted_list.append(
                            {
                                "Asana id": task["gid"],
                                "Asana Name": task["name"],
                                "Excel path": excel_dir,
                                "Project Number": task["name"].split("-")[0].split("\\")[-1],
                                "Quotation Number": quotation_number
                            }
                        )

                        print("Unaccepted Quotation")
                        continue
                else:
                    unprocessed_task_list.append(
                        {
                            "Asana id": task["gid"],
                            "Asana Name": task["name"],
                            "Project Number": task["name"].split("-")[0].split("\\")[-1],
                            "Quotation Number": ""
                        }
                    )

                    print("Only Found Project Number")
                    continue
        else:
            quotation_number = task["name"].split("-")[0][3:]

        if check_if_quotation_exist(quotation_number):
            print("Quotation Exists, continous")
            exist_task+=1
            continue
        data_json = json.load(open(os.path.join(conf["database_dir"], "data_template.json")))
        asana_task = task_api_instance.get_task(task["gid"]).to_dict()["data"]
        sub_tasks = task_api_instance.get_subtasks_for_task(task["gid"]).to_dict()["data"]
        data_json["Project Info"]["Project"]["Project Name"] = asana_task["name"].split("-", 1)[1]
        data_json["Project Info"]["Project"]["Quotation Number"] = quotation_number
        if asana_task["name"][9] == "-":
            data_json["Project Info"]["Project"]["Project Number"] = asana_task["name"][3:9]
        data_json["Asana_id"] = asana_task["gid"]
        data_json["Asana_url"] = asana_task["permalink_url"]
        data_json["Email_Content"] = asana_task["notes"]
        for project in asana_task["projects"]:
            if project["name"] in project_types:
                data_json["Project Info"]["Project"]["Project Type"] = project["name"]
                if project["name"] in ["Group House", "Apartment", "Mixed-use complex"]:
                    data_json["Project Info"]["Project"]["Proposal Type"] = "Major"
                else:
                    data_json["Project Info"]["Project"]["Proposal Type"] = "Minor"
                break
        for custom_field in asana_task["custom_fields"]:
            if custom_field["name"] == "Status":
                if custom_field["display_value"] == "Quote Unsuccessful":
                    data_json["State"]["Set Up"] = True
                    data_json["State"]["Quote Unsuccessful"] = True
                elif custom_field["display_value"] == "Fee Proposal":
                    data_json["State"]["Set Up"] = True
                else:
                    data_json["State"]["Set Up"] = True
                    data_json["State"]["Generate Proposal"] = True
                    data_json["State"]["Email to Client"] = True
                    data_json["State"]["Fee Accepted"] = True
                data_json["State"]["Asana State"] = custom_field["display_value"]
            elif custom_field["name"] == "Services":
                if not custom_field["display_value"] is None:
                    include_services_list = custom_field["display_value"].split(", ")
                    for service in include_services_list:
                        data_json["Project Info"]["Project"]["Service Type"][service]["Include"]=True
            elif custom_field["name"] == "Shop Name":
                if not custom_field["display_value"] is None:
                    data_json["Project Info"]["Project"]["Shop Name"]=custom_field["display_value"]
            elif custom_field["name"] == "Apt/Room/Area":
                if not custom_field["display_value"] is None:
                    data_json["Project Info"]["Building Features"]["Apt"] = custom_field["display_value"]
            elif custom_field["name"] == "Feature/Notes":
                if not custom_field["display_value"] is None:
                    data_json["Project Info"]["Building Features"]["Feature"] = custom_field["display_value"]
            elif custom_field["name"] == "Basement/Car Spots":
                if not custom_field["display_value"] is None:
                    data_json["Project Info"]["Building Features"]["Basement"] = custom_field["display_value"]
            elif custom_field["name"] == "Client":
                if not custom_field["display_value"] is None:
                    data_json["Project Info"]["Client"]["Full Name"] = custom_field["display_value"]
            elif custom_field["name"] == "Main Contact":
                if not custom_field["display_value"] is None:
                    data_json["Project Info"]["Main Contact"]["Full Name"] = custom_field["display_value"]
            elif custom_field["name"] == "Contact Type":
                if not custom_field["display_value"] is None:
                    data_json["Project Info"]["Main Contact"]["Contact Type"] = custom_field["display_value"]
        num = 0
        paid_amount = 0
        for sub_task in sub_tasks:
            if sub_task["name"].startswith("INV"):
                invoice_task = task_api_instance.get_task(sub_task["gid"]).to_dict()["data"]
                if sub_task["name"][4:10].isdigit():
                    number = sub_task["name"][4:10]
                    data_json["Invoices Number"][num]["Number"] = number
                    state = invoice_task["custom_fields"][2]["display_value"]
                    data_json["Invoices Number"][num]["State"] = state
                    inv_json[number] = state
                for custom_field in invoice_task["custom_fields"]:
                    if custom_field["name"] == "Net" and not custom_field["display_value"] is None:
                        data_json["Invoices Number"][num]["Fee"] = custom_field["display_value"]
                    elif custom_field["name"] == "Gross" and not custom_field["display_value"] is None:
                        data_json["Invoices Number"][num]["in.GST"] = custom_field["display_value"]
                num+=1
                if num == 6:
                    task_with_more_than_6_invoice_list.append(
                        {
                            "Asana id": task["gid"],
                            "Asana Name": task["name"],
                            "Project Number": task["name"].split("-")[0].split("\\")[-1]
                        }
                    )
                    print("Total Invoices Number Exceeding 6, Skip the remaining Invoices")
                    break



        os.makedirs(os.path.join(database_dir, quotation_number))
        log.log_sync_from_asana("Admin", quotation_number)
        with open(os.path.join(database_dir, quotation_number, "data.json"), "w") as f:
            json_object = json.dumps(data_json, indent=4)
            f.write(json_object)



        with open(os.path.join(database_dir, "invoices.json"), "w") as f:
            json_object = json.dumps(inv_json, indent=4)
            f.write(json_object)
        with open(os.path.join(database_dir, "unprocessed_tasks.json"), "w") as f:
            json_object = json.dumps(unprocessed_task_list, indent=4)
            f.write(json_object)

        with open(os.path.join(database_dir, "quotation_number_not_accepted.json"), "w") as f:
            json_object = json.dumps(quotation_number_not_accepted_list, indent=4)
            f.write(json_object)

        with open(os.path.join(database_dir, "project_with_more_than_6_invoices.json"), "w") as f:
            json_object = json.dumps(task_with_more_than_6_invoice_list, indent=4)
            f.write(json_object)



    if ori_tasks.to_dict()["next_page"] is None:
        with open(os.path.join(database_dir, "invoices.json"), "w") as f:
            json_object = json.dumps(inv_json, indent=4)
            f.write(json_object)
        with open(os.path.join(database_dir, "unprocessed_tasks.json"), "w") as f:
            json_object = json.dumps(unprocessed_task_list, indent=4)
            f.write(json_object)
        print(f"Exist Tasks: {exist_task} {exist_task/i}")
        print(f"Unprocessed Tasks: {len(unprocessed_task_list)} {len(unprocessed_task_list) / i}")
        print(f"The Sync take {time.time()-start}s")
        print("Download Complete!!!")
        break
    off_set = ori_tasks.to_dict()["next_page"]["offset"]

