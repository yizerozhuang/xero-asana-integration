import asana
from config import CONFIGURATION as conf
from app_log import AppLog
from asana_function import name_id_map, clean_response, flatter_custom_fields

from utility import generate_all_project, isfloat

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
accounting_dir = conf["accounting_dir"]

inv_dir = os.path.join(conf["database_dir"], "invoices.json")
inv_json = json.load(open(inv_dir))

scope_dir = os.path.join(conf["database_dir"], "scope_of_work.json")
scope_json = json.load(open(scope_dir))

scope_of_staging_dir = os.path.join(conf["database_dir"], "general_scope_of_staging.json")
scope_of_staging_json = json.load(open(scope_of_staging_dir))


all_bridge_projects = generate_all_project(conf["database_dir"])

assign_quotation_map = {}


manuel_quotation_dir = os.path.join(conf["database_dir"], "manuel_assign_quotation.json")
manuel_quotation_json = json.load(open(manuel_quotation_dir))

log=AppLog()

MP_project_gid = "1203405141297991"
project_types = ["Restaurant", "Office", "Commercial", "Group House", "Apartment", "Mixed-use Complex", "School", "Others"]
excel = win32client.Dispatch("Excel.Application")



def create_project_by_quotation(quotation_number, asana_task):
    print(f"Loading asana project with quotation: {quotation_number} into Bridge")
    if check_if_quotation_exist(quotation_number):
        print(f"Quotation {quotation_number} Exists")
        raise ValueError
    data_json = json.load(open(os.path.join(conf["database_dir"], "data_template.json")))
    subtasks_opt = ["name", "custom_fields", "projects", "projects.name"]
    sub_tasks = clean_response(task_api_instance.get_subtasks_for_task(asana_task["gid"], opt_fields=subtasks_opt))
    data_json["Project Info"]["Project"]["Project Name"] = asana_task["name"].split("-", 1)[1]


    data_json["Project Info"]["Project"]["Quotation Number"] = quotation_number
    if asana_task["name"].split("-")[0].split("P:\\")[-1].isdigit():
        data_json["Project Info"]["Project"]["Project Number"] = asana_task["name"].split("-")[0].split("P:\\")[-1]
    data_json["Asana_id"] = asana_task["gid"]
    data_json["Asana_url"] = asana_task["permalink_url"]
    data_json["Email_Content"] = asana_task["notes"]
    for project in asana_task["projects"]:
        if project["name"] in project_types:
            data_json["Project Info"]["Project"]["Project Type"] = project["name"]
            if project["name"] in ["Group House", "Apartment", "Mixed-use Complex"]:
                data_json["Project Info"]["Project"]["Proposal Type"] = "Major"
            else:
                data_json["Project Info"]["Project"]["Proposal Type"] = "Minor"
            break
        if asana_task["Status"] == "Quote Unsuccessful":
            data_json["State"]["Set Up"] = True
            data_json["State"]["Quote Unsuccessful"] = True
        elif asana_task["Status"] == "Fee Proposal":
            data_json["State"]["Set Up"] = True
            data_json["State"]["Generate Proposal"] = True
            data_json["State"]["Email to Client"] = True
        else:
            data_json["State"]["Set Up"] = True
            data_json["State"]["Generate Proposal"] = True
            data_json["State"]["Email to Client"] = True
            data_json["State"]["Fee Accepted"] = True
            data_json["State"]["Asana State"] = asana_task["Status"]
        data_json["Email"]["Fee Proposal"] = asana_task["created_at"].strftime("%Y-%m-%d")
        if not asana_task["Services"] is None:
            include_services_list = asana_task["Services"].split(", ")
            for service in include_services_list:
                data_json["Project Info"]["Project"]["Service Type"][service]["Include"] = True
                if service != "Kitchen Ventilation":
                    data_json["Invoices"]["Details"][service]["Expand"] = True
                    data_json["Invoices"]["Details"][service]["Include"] = True
                    data_json["Bills"]["Details"][service]["Include"] = True
                    data_json["Profits"]["Details"][service]["Include"] = True
                else:
                    data_json["Invoices"]["Details"]["Mechanical Service"]["Expand"] = True
                    data_json["Invoices"]["Details"]["Mechanical Service"]["Include"] = True
                    data_json["Bills"]["Details"]["Mechanical Service"]["Include"] = True
                    data_json["Profits"]["Details"]["Mechanical Service"]["Include"] = True
        if not asana_task["Shop name"] is None:
            data_json["Project Info"]["Project"]["Shop Name"] = asana_task["Shop name"]
        if not asana_task["Apt/Room/Area"] is None:
            data_json["Project Info"]["Building Features"]["Apt"] = asana_task["Apt/Room/Area"]
        if not asana_task["Feature/Notes"] is None:
            data_json["Project Info"]["Building Features"]["Feature"] = asana_task["Feature/Notes"]
        if not asana_task["Basement/Car Spots"] is None:
            data_json["Project Info"]["Building Features"]["Basement"] = asana_task["Basement/Car Spots"]
        if not asana_task["Client"] is None:
            if "-" in asana_task["Client"]:
                company_name, full_name = asana_task["Client"].split("-", 1)
                data_json["Project Info"]["Client"]["Full Name"] = full_name.strip()
                data_json["Project Info"]["Client"]["Company"] = company_name.strip()
            else:
                data_json["Project Info"]["Client"]["Full Name"] = asana_task["Client"]
        data_json["Project Info"]["Client"]["Contact Type"] = "None"
        if not asana_task["Main Contact"] is None:
            if "-" in asana_task["Main Contact"]:
                company_name, full_name = asana_task["Main Contact"].split("-", 1)
                data_json["Project Info"]["Main Contact"]["Full Name"] = full_name.strip()
                data_json["Project Info"]["Main Contact"]["Company"] = company_name.strip()
            else:
                data_json["Project Info"]["Main Contact"]["Full Name"] = asana_task["Main Contact"]
        data_json["Address_to"] = "Main Contact" if asana_task["Client"] is None else "Client"
        if not asana_task["Contact Type"] is None:
            data_json["Project Info"]["Main Contact"]["Contact Type"] = asana_task["Contact Type"]
    num = 0
    # paid_amount = 0
    # over_due_amount = 0
    # amount = 0
    for sub_task in sub_tasks:
        if "Invoice status" in [project["name"] for project in sub_task["projects"]]:
            if num == 6:
                task_with_more_than_6_invoice_list.append(
                    {
                        "Asana id": asana_task["gid"],
                        "Asana Name": asana_task["name"],
                        "Project Number": asana_task["name"].split("-")[0].split("\\")[-1]
                    }
                )
                print("Total Invoices Number Exceeding 6, Skip the remaining Invoices")
                break
            invoice_task = flatter_custom_fields(clean_response(task_api_instance.get_task(sub_task["gid"])))
            if sub_task["name"][4:10].isdigit():
                number = sub_task["name"][4:10]
                data_json["Invoices Number"][num]["Number"] = number
                inv_json[number] = invoice_task["Invoice status"]
            data_json["Invoices Number"][num]["Fee"] = invoice_task["Net"]
            data_json["Invoices Number"][num]["in.GST"] = invoice_task["Gross"]
            data_json["Invoices Number"][num]["State"] = invoice_task["Invoice status"]
            data_json["Invoices Number"][num]["Asana_id"] = invoice_task["gid"]
            # for custom_field in invoice_task["custom_fields"]:
            #     if custom_field["name"] == "Net" and not custom_field["display_value"] is None:
            #         amount += float(custom_field["display_value"])
            #         if state == "Paid":
            #             paid_amount += float(custom_field["display_value"])
            #         elif state == "Sent":
            #             over_due_amount += float(custom_field["display_value"])
            #         data_json["Invoices Number"][num]["Fee"] = custom_field["display_value"]
            #     elif custom_field["name"] == "Gross" and not custom_field["display_value"] is None:
            #         data_json["Invoices Number"][num]["in.GST"] = custom_field["display_value"]
            num += 1

    fee = "0" if not isfloat(asana_task["Fee ExGST"]) else str(asana_task["Fee ExGST"])
    ingst = str(float(fee)*1.1)
    paid_fee = "0" if not isfloat(asana_task["Total Paid ExGST"]) else str(asana_task["Total Paid ExGST"])
    over_due_fee = "0" if not isfloat(asana_task["Overdue Amount"]) else str(asana_task["Overdue Amount"])

    data_json["Invoices"]["Fee"] = fee
    data_json["Invoices"]["in.GST"] = ingst
    data_json["Invoices"]["Paid Fee"] = paid_fee
    data_json["Invoices"]["Over Due Fee"] = over_due_fee

    stage_dic = {}
    for stage in scope_of_staging_json.keys():
        stage_dic[stage] = {
            "Service": "",
            "Include": False,
            "Items": []
        }
        for item in scope_of_staging_json[stage]:
            stage_dic[stage]["Items"].append({
                "Include": False,
                "Item": item
            }
            )
    data_json["Fee Proposal"]["Stage"] = stage_dic

    scope_dic = {}
    for project_type in scope_json.keys():
        scope_dic[project_type] = {}
        for service in scope_json[project_type]:
            scope_dic[project_type][service] = {}
            for extra in conf["extra_list"]:
                scope_dic[project_type][service][extra] = []
                for item in scope_json[project_type][service][extra]:
                    scope_dic[project_type][service][extra].append({
                        "Include": False,
                        "Item": item
                    }
                    )
    data_json["Fee Proposal"]["Scope"] = scope_dic
    data_json["Fee Proposal"]["Reference"]["Revision"] = ""
    data_json["Fee Proposal"]["Installation Reference"]["Revision"] = ""
    data_json["Fee Proposal"]["Time"] = {
        "Fee Proposal": "",
        "Pre-design": "",
        "Documentation": ""
    }

    # all_custom_fields = asana_task["custom_fields"]
    # custom_field_id_map = name_id_map(all_custom_fields)

    # asana_update_body = asana.TasksTaskGidBody({
    #     "custom_fields": {
    #         custom_field_id_map["Fee ExGST"]: amount,
    #         custom_field_id_map["Total Paid ExGST"]: paid_amount
    #     }
    # })
    # task_api_instance.update_task(task_gid=asana_task["gid"], body=asana_update_body)

    os.makedirs(os.path.join(database_dir, quotation_number))
    if not os.path.exists(os.path.join(accounting_dir, quotation_number)):
        os.makedirs(os.path.join(accounting_dir, quotation_number))
    log.log_sync_from_asana("Admin", quotation_number)
    with open(os.path.join(database_dir, quotation_number, "data.json"), "w") as f:
        json_object = json.dumps(data_json, indent=4)
        f.write(json_object)


def assign_quotation(project_number, quote_unsuccessful):
    prefix = project_number[0:3]+"000"
    if quote_unsuccessful:
        current_letter = "ZA"
    else:
        current_letter = "AA"
    while True:
        if not os.path.exists(os.path.join(database_dir, prefix+current_letter)):
            return prefix+current_letter
        current_letter = current_letter[0] + chr(ord(current_letter[1]) + 1) if current_letter[1] != "Z" else chr(ord(current_letter[0]) + 1) + "A"

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
        elif task["gid"] in all_bridge_projects.keys():
            print(f"Asana Task {task['gid']} exist to project {all_bridge_projects[task['gid']]}")
            continue
        if task["name"].split("-")[0].split("P:\\")[-1].isdigit():
            # #project number
            # project_number = task["name"][3:9]
            # #Story mathod
            # stories = stories_api_instance.get_stories_for_task(task["gid"]).to_dict()["data"]
            # for story in stories:
            #     text = story["text"]
            #     if "changed the name to" in text:
            #         quotation_number = text.split('"')[1].split("-")[0].split("P:\\")[-1].strip()
            #         break
            # if quotation_number is None or type(quotation_number) == float:
            #     quotation_number = project_number
            # if quotation_number.isdigit() or i == 366:
            #     folder_dir = task["name"]
            #     excel_dir = None
            #     for root, dirs, files in os.walk(folder_dir):
            #         for file in files:
            #             if "Fee Proposal" in str(file) and file.endswith(".xlsx"):
            #                 excel_dir = os.path.join(root, file)
            #                 break
            #     if not excel_dir is None:
            #         work_book = excel.Workbooks.Open(excel_dir)
            #         work_sheets = work_book.Worksheets[0]
            #         quotation_number = work_sheets.Cells(3, 2).Value
            #         work_book.Close(False)
            #         print(f"Found Quotation Number in {excel_dir} with quotation number {quotation_number}")
            #         if quotation_number is None or str(quotation_number).isdigit() or type(quotation_number)==float:
            #             quotation_number_not_accepted_list.append(
            #                 {
            #                     "Asana id": task["gid"],
            #                     "Asana Name": task["name"],
            #                     "Excel path": excel_dir,
            #                     "Project Number": task["name"].split("-")[0].split("\\")[-1],
            #                     "Quotation Number": quotation_number
            #                 }
            #             )
            #
            #             print("Unaccepted Quotation")
            #             generate_quotation = True
            #     else:
            #         unprocessed_task_list.append(
            #             {
            #                 "Asana id": task["gid"],
            #                 "Asana Name": task["name"],
            #                 "Project Number": task["name"].split("-")[0].split("\\")[-1],
            #                 "Quotation Number": ""
            #             }
            #         )
            #
            #         print("Only Found Project Number")
            project_number = task["name"].split("-")[0].split("P:\\")[-1]
            if project_number in manuel_quotation_json.keys():
                quotation_number = manuel_quotation_json[project_number]
                print(f"Using Manuel Quotation Number {quotation_number} for Project Number {project_number}")
                assert not os.path.exists(os.path.join(database_dir, quotation_number))
            else:
                assign_quotation_map[project_number] = {
                    "Asana id": task["gid"]
                }
                print(f"found project number {project_number} for asana project {task['name']}")
                continue
        else:
            # quotation number
            quotation_number = task["name"].split("-")[0].split("P:\\")[-1]
            print(f"found quotation number {quotation_number} for asana project {task['name']}")
        asana_task = flatter_custom_fields(clean_response(task_api_instance.get_task(task["gid"])))
        create_project_by_quotation(quotation_number, asana_task)

    if ori_tasks.to_dict()["next_page"] is None:
        for project_number, value in assign_quotation_map.items():
            asana_task = flatter_custom_fields(clean_response(task_api_instance.get_task(value["Asana id"])))
            for custom_task in asana_task["custom_fields"]:
                if custom_task["name"] == 'Status':
                    state = custom_task["display_value"]
                    break
            if state == "Quote Unsuccessful":
                quotation_number = assign_quotation(project_number, True)
                assign_quotation_map[project_number]["Quotation"] = quotation_number
                print(f"Assign quotation {quotation_number} to project {project_number}, this project is quote unsuccessful")
                create_project_by_quotation(quotation_number, asana_task)
            else:
                quotation_number = assign_quotation(project_number, False)
                assign_quotation_map[project_number]["Quotation"] = quotation_number
                print(f"Assign quotation {quotation_number} to project {project_number}")
                create_project_by_quotation(quotation_number, asana_task)

        inv_json_list = list(inv_json.keys())
        inv_json_list.sort()
        new_inv_list = {}
        for inv_number in inv_json_list:
            new_inv_list[inv_number] = inv_json[inv_number]

        with open(os.path.join(database_dir, "invoices.json"), "w") as f:
            json_object = json.dumps(new_inv_list, indent=4)
            f.write(json_object)
        with open(os.path.join(database_dir, "unprocessed_tasks.json"), "w") as f:
            json_object = json.dumps(unprocessed_task_list, indent=4)
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

        with open(os.path.join(database_dir, "assign_quotation_map.json"), "w") as f:
            json_object = json.dumps(assign_quotation_map, indent=4)
            f.write(json_object)
        print(f"Exist Tasks: {exist_task} {exist_task/i}")
        print(f"Unprocessed Tasks: {len(unprocessed_task_list)} {len(unprocessed_task_list) / i}")
        print(f"The Sync take {time.time()-start}s")
        print("Download Complete!!!")
        break
    off_set = ori_tasks.to_dict()["next_page"]["offset"]
    print(f"Finish Update current page the next off_set_id is {off_set}")

