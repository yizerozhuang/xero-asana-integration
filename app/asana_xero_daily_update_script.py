from xero_python.accounting import AccountingApi
import asana
import dateutil.parser
from config import CONFIGURATION as conf
from utility import generate_all_projects_and_invoices, isfloat, increment_excel_column
from xero_function import xero_token_required, refresh_token, get_xero_tenant_id, api_client
from asana_function import clean_response, name_id_map, flatter_custom_fields
from search_bar_page import return_self_mp

import os
import json
import time
import shutil
import openpyxl
from datetime import date, timedelta
from win32com import client as win32client


# database_dir = conf["database_dir"]
database_dir = "P:\\app\\database"
backup_dir = "C:\\Users\\Admin\\Desktop\\test_back_up"
working_dir = "P:"
accounting_dir = "A:\\00-Bridge Database"

invoice_dir = os.path.join(database_dir, "invoices.json")
invoice_json = json.load(open(invoice_dir))
bill_dir = os.path.join(database_dir, "bills.json")
bill_json = {}

xero_asana_status_map = {
    "DRAFT": "Draft",
    "SUBMITTED": "Awaiting Approval",
    "AUTHORISED": "Awaiting Payment",
    "PAID": "Paid",
    "VOIDED": "Voided"
}


# generate_management_report = False

if_modified_since = dateutil.parser.parse("2020-01-01")

asana_configuration = asana.Configuration()
asana_configuration.access_token = '2/1203283895754383/1206354773081941:c116d68430be7b2832bf5d7ea2a0a415'
asana_api_client = asana.ApiClient(asana_configuration)
custom_fields_api_instance = asana.CustomFieldsApi(asana_api_client)
custom_fields_setting_api_instance = asana.CustomFieldSettingsApi(asana_api_client)
invoice_status_asana_gid = "1203204784052206"
bill_status_asana_gid = "1203219273836510"
mp_asana_gid = "1203405141297991"
# project_api_instance = asana.ProjectsApi(asana_api_client)
task_api_instance = asana.TasksApi(asana_api_client)

all_project, backlog_invoice, all_bills = generate_all_projects_and_invoices(database_dir)
def _process_invoice(invoices, accounting_api, xero_tenant_id):
    res = {
        "INV": {
            "Sent": {},
            "Paid": {},
            "Voided": {}
        },
        "BIL": {}
    }
    invoices = invoices.to_dict()["invoices"]
    for invoice in invoices:
        if invoice["type"] == "ACCREC":
            if not invoice["invoice_number"].isdigit():
                continue
            elif invoice["invoice_number"][0] in ["8", "9"]:
                continue
            if invoice["status"] in ["DRAFT", "SUBMITTED", "AUTHORISED"]:
                res["INV"]["Sent"][invoice["invoice_number"]] = {
                        "invoice_number": invoice["invoice_number"],
                        "Payment Date": None if len(invoice["payments"]) == 0 else invoice["payments"][0]["date"],
                        "Payment InGST": None if float(invoice["amount_paid"]) == 0 else str(invoice["amount_paid"]),
                        "Net": str(invoice["sub_total"]),
                        "Gross": str(invoice["total"]),
                        "reference": invoice["reference"]
                }
            elif invoice["status"] == "PAID":
                res["INV"]["Paid"][invoice["invoice_number"]] = {
                        "invoice_number": invoice["invoice_number"],
                        "Payment Date": invoice["fully_paid_on_date"],
                        "Payment InGST": None if float(invoice["amount_paid"]) == 0 else str(invoice["amount_paid"]),
                        "Net": str(invoice["sub_total"]),
                        "Gross": str(invoice["total"]),
                        "reference": invoice["reference"]
                    }
            elif invoice["status"] == "VOIDED":
                res["INV"]["Voided"][invoice["invoice_number"]] = {
                        "invoice_number": invoice["invoice_number"],
                        "Payment Date": None,
                        "Payment InGST": None,
                        "Net": str(invoice["sub_total"]),
                        "Gross": str(invoice["total"]),
                        "reference": invoice["reference"]
                }
        elif invoice["type"] == "ACCPAY":
            if invoice["status"] in ["DRAFT", "SUBMITTED", "AUTHORISED", "PAID", "VOIDED"]:
                res["BIL"][invoice["invoice_id"]]={
                    "name": invoice["invoice_number"].split("-")[0],
                    "Bill status": xero_asana_status_map[invoice["status"]],
                    "From": invoice["contact"]["name"],
                    "Issue Date": str(invoice["date"]),
                    "Amount Excl GST": str(invoice["sub_total"]),
                    "Amount Incl GST": str(invoice["total"]),
                    "HeadsUp": str(accounting_api.get_invoice(xero_tenant_id, invoice["invoice_id"]).to_dict()["invoices"][0]["line_items"][0]["description"])
                }
            # elif invoice["status"] == "SUBMITTED":
            #     res["BIL"]["Awaiting Approval"].append(invoice["invoice_number"])
            # elif invoice["status"] == "AUTHORISED":
            #     res["BIL"]["Awaiting Payment"].append(invoice["invoice_number"])
            # elif invoice["status"] == "PAID":
            #     res["BIL"]["Paid"].append(invoice["invoice_number"])
            # elif invoice["status"] == "VOIDED":
            #     res["BIL"]["Voided"].append(invoice["invoice_number"])
        else:
            raise TypeError
            # print("Error, Unsupported Invoice Type")
            # return
    return res
@xero_token_required
def update_xero_and_asana_invoice_script():
    start = time.time()
    xero_tenant_id = get_xero_tenant_id()
    accounting_api = AccountingApi(api_client)

    invoices = _process_invoice(accounting_api.get_invoices(xero_tenant_id, if_modified_since=if_modified_since), accounting_api, xero_tenant_id)
    for invoices_states in invoices["INV"].keys():
        for invoice_number in invoices["INV"][invoices_states].keys():
            invoice_json[invoice_number] = invoices_states

    for key, value in invoices["BIL"].items():
        bill_json[key] = value["Bill status"]

    off_set = None
    opt_field = ["name", "custom_fields.name", "custom_fields.display_value"]

    custom_fields_setting = clean_response(custom_fields_setting_api_instance.get_custom_field_settings_for_project(invoice_status_asana_gid))
    all_custom_fields = [custom_field["custom_field"] for custom_field in custom_fields_setting]
    custom_field_id_map = name_id_map(all_custom_fields)

    status_field = clean_response(
        custom_fields_api_instance.get_custom_field(custom_field_id_map["Invoice status"]))
    status_id_map = name_id_map(status_field["enum_options"])


    bill_custom_fields_setting = clean_response(custom_fields_setting_api_instance.get_custom_field_settings_for_project(bill_status_asana_gid))
    bill_all_custom_fields = [custom_field["custom_field"] for custom_field in bill_custom_fields_setting]
    bill_custom_field_id_map = name_id_map(bill_all_custom_fields)

    bill_status_field = clean_response(
        custom_fields_api_instance.get_custom_field(bill_custom_field_id_map["Bill status"]))
    bill_status_id_map = name_id_map(bill_status_field["enum_options"])
    # if generate_management_report:
    #     resource_dir = conf["resource_dir"]
    #     management_report_template = os.path.join(resource_dir, "xlsx", "management_report_template.xlsx")
    #     management_report = os.path.join(conf["database_dir"], f"Management Report {str(date.today().strftime('%Y%m%d'))}.xlsx")
    #     shutil.copy(management_report_template, management_report)
    #     excel = win32client.Dispatch("Excel.Application")
    #     work_book = excel.Workbooks.Open(management_report)
    #     work_sheets = work_book.Worksheets[1]
    #     cur_row = 8

    #update_invoice
    while True:
        if off_set is None:
            ori_tasks = task_api_instance.get_tasks_for_project(project_gid=invoice_status_asana_gid, opt_fields=opt_field, limit=100).to_dict()
        else:
            ori_tasks = task_api_instance.get_tasks_for_project(project_gid=invoice_status_asana_gid, opt_fields=opt_field, limit=100, offset=off_set).to_dict()
        for task in ori_tasks["data"]:
            task = flatter_custom_fields(task)
            print(f"Processing task {task['name'][4:10]}, the gid is {task['gid']}")
            for invoices_states in invoices["INV"].keys():
                if task["name"][4:10] in invoices["INV"][invoices_states].keys():
                    inv = invoices["INV"][invoices_states][task["name"][4:10]]
                    custom_field_body = {}

                    if task["Invoice status"] != invoices_states:
                        custom_field_body[custom_field_id_map["Invoice status"]] = status_id_map[invoices_states]

                    display_value = None if task["Payment Date"] is None else task["Payment Date"][0:10]
                    payment_date = None if inv["Payment Date"] is None else str(inv["Payment Date"])
                    if display_value != payment_date:
                        if inv["Payment Date"] is None:
                            custom_field_body[custom_field_id_map["Payment Date"]] = None
                        else:
                            custom_field_body[custom_field_id_map["Payment Date"]] = {"date": str(inv["Payment Date"])}

                    if task["Payment InGST"] != inv["Payment InGST"]:
                        custom_field_body[custom_field_id_map["Payment InGST"]] = inv["Payment InGST"]

                    if task["Net"] != inv["Net"]:
                        custom_field_body[custom_field_id_map["Net"]] = inv["Net"]

                    # for custom_field in task["custom_fields"]:
                    #     if custom_field["name"] == "Invoice status" and custom_field["display_value"] != invoices_states:
                    #         custom_field_body[custom_field_id_map["Invoice status"]] = status_id_map[invoices_states]
                    #     elif custom_field["name"] == "Payment Date":
                    #         # if custom_field["display_value"] != inv["Payment Date"]:
                    #         display_value = None if custom_field["display_value"] is None else custom_field["display_value"][0:10]
                    #         payment_date = None if inv["Payment Date"] is None else str(inv["Payment Date"])
                    #         if display_value != payment_date:
                    #            custom_field_body[custom_field_id_map["Payment Date"]] = {"date":str(inv["Payment Date"])}
                    #     elif custom_field["name"] == "Payment Amount" and custom_field["display_value"] != inv["Payment Amount"]:
                    #         custom_field_body[custom_field_id_map["Payment Amount"]] = inv["Payment Amount"]
                        # elif custom_field["name"] == "Net" and custom_field["display_value"] != inv["Net"]:
                        #     if task["name"][4:10] in net_gross_different_dict.keys():
                        #         net_gross_different_dict[task["name"][4:10]]["Net"]={
                        #                 "Asana":custom_field["display_value"],
                        #                 "Xero":inv["Net"]
                        #             }
                        #     else:
                        #         net_gross_different_dict[task["name"][4:10]] = {
                        #             "Net":{
                        #                 "Asana":custom_field["display_value"],
                        #                 "Xero":inv["Net"]
                        #             }
                        #         }
                    if len(custom_field_body) !=0:
                        body = asana.TasksTaskGidBody(
                            {
                                "custom_fields":custom_field_body
                            }
                        )

                        task_api_instance.update_task(task_gid=task["gid"], body=body)
                        print(f"Processed task with asana id {task['gid']} and invoice number {task['name']}")
                        #     custom_field_body[custom_field_id_map["Invoice status"]]: status_id_map[invoices_states]
                    # if generate_management_report:
                    #     work_sheets.Cells(cur_row, 1).Value = inv["reference"]
                    #     work_sheets.Cells(cur_row, 2).Value = task["name"][4:10]

                        # work_sheets.Cells(cur_row, 1).Value = inv["reference"]
                        # work_sheets.Cells(cur_row, 1).Value = inv["reference"]
            if task["gid"] in backlog_invoice.keys():
                task_body = {"custom_fields":{}}
                update = False
                if len(backlog_invoice[task["gid"]]["Number"])==0 and task["name"]!="INV xxxxxx":
                    task_body["name"] = "INV xxxxxx"
                    update = True
                elif len(backlog_invoice[task["gid"]]["Number"])!=0 and task["name"]!=f"INV {backlog_invoice[task['gid']]['Number']}":
                    task_body["name"] = f"INV {task['Number']}"
                    update = True

                if float(backlog_invoice[task["gid"]]["Fee"]) != float(task["Net"]):
                    task_body["custom_fields"][custom_field_id_map["Net"]] = float(backlog_invoice[task["gid"]]["Fee"])
                    update = True

                if update:
                    body = asana.TasksTaskGidBody(task_body)
                    task_api_instance.update_task(task_gid=task["gid"], body=body)
                    print(f"Processed backlog task with asana id {task['gid']} and invoice number {task['name']}")

        if ori_tasks["next_page"] is None:
            break
        off_set = ori_tasks["next_page"]["offset"]
        # messagebox.showerror("Error", e)



    # update_bills
    # for bill in all_bills.keys():
    #     asana_id = all_bills[bill]
    #     asana_task = flatter_custom_fields(clean_response(task_api_instance.get_task(asana_id)))
    #     print(bill)
    #     assert bill in invoices["BIL"].keys()
    #     xero_bill = invoices["BIL"][bill]
    #
    #     asana_update_body = {"custom_fields":{}}
    #     update = False
    #     print(f"Processed Bill with asana id {asana_id} with asana task name {asana_task['name']}")
    #     if asana_task["name"] != "BIL "+xero_bill["name"]:
    #         asana_update_body["name"] = "BIL "+xero_bill["name"]
    #         update = True
    #     if asana_task["Bill status"] != xero_bill["Bill status"]:
    #         asana_update_body["custom_fields"][bill_custom_field_id_map["Bill status"]] = bill_status_id_map[xero_bill["Bill status"]]
    #         update = True
    #
    #     if asana_task["From"] != xero_bill["From"]:
    #         asana_update_body["custom_fields"][bill_custom_field_id_map["From"]] = xero_bill["From"]
    #         update = True
    #
    #     display_value = None if asana_task["Bill In Date"] is None else asana_task["Bill In Date"][0:10]
    #     issue_date = None if xero_bill["Issue Date"] is None else xero_bill["Issue Date"]
    #     if display_value != issue_date:
    #         if issue_date is None:
    #             asana_update_body["custom_fields"][bill_custom_field_id_map["Bill In Date"]] = None
    #         else:
    #             asana_update_body["custom_fields"][bill_custom_field_id_map["Bill In Date"]] = {"date": issue_date}
    #         update=True
    #
    #     if float(asana_task["Amount Excl GST"]) != float(xero_bill["Amount Excl GST"]):
    #         asana_update_body["custom_fields"][bill_custom_field_id_map["Amount Excl GST"]] = float(xero_bill["Amount Excl GST"])
    #         update = True
    #
    #     if float(asana_task["Amount Incl GST"]) != float(xero_bill["Amount Incl GST"]):
    #         asana_update_body["custom_fields"][bill_custom_field_id_map["Amount Incl GST"]] = float(xero_bill["Amount Incl GST"])
    #         update = True
    #
    #     if update:
    #         body = asana.TasksTaskGidBody(asana_update_body)
    #         task_api_instance.update_task(task_gid=asana_id, body=body)
    #         print(f"Update Bill with asana id {asana_id} with asana task name {asana_task['name']}")

    inv_json_list = list(invoice_json.keys())
    inv_json_list.sort()
    new_inv_list = {}
    for inv_number in inv_json_list:
        new_inv_list[inv_number] = invoice_json[inv_number]

    with open(os.path.join(database_dir, "invoices.json"), "w") as f:
        json_object = json.dumps(new_inv_list, indent=4)
        f.write(json_object)
    with open(os.path.join(database_dir, "bills.json"), "w") as f:
        json_object = json.dumps(bill_json, indent=4)
        f.write(json_object)
    print(f"The Sync take {time.time() - start}s")
    print("Complete")

def update_asana_project_script():
    start = time.time()
    off_set = None
    # opt_field = ["name", "custom_fields.name", "custom_fields.display_value"]

    mp_dir = os.path.join(database_dir, "mp.json")
    mp_json = json.load(open(mp_dir))
    mp_convert_map = return_self_mp()


    custom_fields_setting = clean_response(
        custom_fields_setting_api_instance.get_custom_field_settings_for_project(mp_asana_gid))
    all_custom_fields = [custom_field["custom_field"] for custom_field in custom_fields_setting]
    custom_field_id_map = name_id_map(all_custom_fields)

    status_field = clean_response(custom_fields_api_instance.get_custom_field(custom_field_id_map["Status"]))
    status_id_map = name_id_map(status_field["enum_options"])

    contact_field = clean_response(custom_fields_api_instance.get_custom_field(custom_field_id_map["Contact Type"]))
    contact_id_map = name_id_map(contact_field["enum_options"])

    service_filed = clean_response(custom_fields_api_instance.get_custom_field(custom_field_id_map["Services"]))
    service_id_map = name_id_map(service_filed["enum_options"])

    task_opt = ["name", "custom_fields"]
    sub_task_opt = ["projects", "custom_fields"]
    while True:
        if off_set is None:
            ori_tasks = task_api_instance.get_tasks_for_project(project_gid=mp_asana_gid, limit=100, opt_fields=task_opt)
        else:
            ori_tasks = task_api_instance.get_tasks_for_project(project_gid=mp_asana_gid, limit=100, opt_fields=task_opt, offset=off_set)
        tasks_list = clean_response(ori_tasks)
        for task in tasks_list:
            if task["gid"] in all_project.keys():
                print(f"Processing task: {task['name']} with Asana Gid: {task['gid']}")
                total_amount = 0
                paid_amount = 0
                over_due_amount = 0
                task = flatter_custom_fields(task)
                sub_task_list = clean_response(task_api_instance.get_subtasks_for_task(task_gid=task["gid"], opt_fields=sub_task_opt))
                for sub_task in sub_task_list:
                    if sub_task["projects"] is None:
                        continue
                    elif invoice_status_asana_gid in [project["gid"] for project in sub_task["projects"]]:
                        sub_task = flatter_custom_fields(sub_task)
                        invoice_status = sub_task["Invoice status"]
                        net = float(sub_task["Net"])
                        gross = float(sub_task["Gross"])
                        payment_amount = 0 if not isfloat(sub_task["Payment InGST"]) else float(sub_task["Payment InGST"])
                        total_amount += net
                        paid_amount += payment_amount
                        if invoice_status == "Sent":
                            over_due_amount += gross - payment_amount
                quotation = all_project[task["gid"]]["Quotation Number"]
                data_json = json.load(open(os.path.join(database_dir, quotation, "data.json")))

                update = False
                update_body = {
                    "custom_fields":{}
                }

                if len(data_json["Project Info"]["Project"]["Project Number"]) != 0:
                    project_name = "P:\\"+ data_json["Project Info"]["Project"]["Project Number"]+"-"+data_json["Project Info"]["Project"]["Project Name"]
                else:
                    project_name = "P:\\" + data_json["Project Info"]["Project"]["Quotation Number"] + "-" + data_json["Project Info"]["Project"]["Project Name"]

                if task["name"] != project_name:
                    update_body["name"] = project_name
                    update = True





                service_list = sorted([service for service in conf["all_service_list"] if data_json["Project Info"]["Project"]["Service Type"][service]["Include"]])
                asana_service_list = [] if task["Services"] is None else sorted(task["Services"].split(", "))
                # asana_service_list = task["Services"].split(", ")
                if service_list != asana_service_list:
                    if len(service_list) == 0:
                        update_body["custom_fields"]["Services"] = None
                    else:
                        update_body["custom_fields"][custom_field_id_map["Services"]] = [service_id_map[service] for service in service_list]
                    update = True

                status = data_json["State"]["Asana State"]

                if task["Status"] != status and not (task["Status"] is None and status == ""):
                    update_body["custom_fields"][custom_field_id_map["Status"]] = status_id_map[status]
                    update = True


                shop_name = data_json["Project Info"]["Project"]["Shop Name"]
                if task["Shop name"] != shop_name and not (task["Shop name"] is None and shop_name == ""):
                    update_body["custom_fields"][custom_field_id_map["Shop name"]] = shop_name
                    update = True

                apt = data_json["Project Info"]["Building Features"]["Apt"]
                if task["Apt/Room/Area"] != apt and not (task["Apt/Room/Area"] is None and apt == ""):
                    update_body["custom_fields"][custom_field_id_map["Apt/Room/Area"]] = apt
                    update = True

                basement = data_json["Project Info"]["Building Features"]["Basement"]
                if task["Basement/Car Spots"] != basement and not (task["Basement/Car Spots"] is None and basement == ""):
                    update_body["custom_fields"][custom_field_id_map["Basement/Car Spots"]] = basement
                    update = True

                feature = data_json["Project Info"]["Building Features"]["Feature"]
                if task["Feature/Notes"] != feature and not (task["Feature/Notes"] is None and feature == ""):
                    update_body["custom_fields"][custom_field_id_map["Feature/Notes"]] = feature
                    update = True

                client_name_list = []
                if len(data_json["Project Info"]["Client"]["Company"]) != 0:
                    client_name_list.append(data_json["Project Info"]["Client"]["Company"])
                if len(data_json["Project Info"]["Client"]["Full Name"]) != 0:
                    client_name_list.append(data_json["Project Info"]["Client"]["Full Name"])
                client_name = "-".join(client_name_list)

                if task["Client"] != client_name and not (task["Client"] is None and client_name == ""):
                    update_body["custom_fields"][custom_field_id_map["Client"]] = client_name
                    update = True

                main_contact_list = []
                if len(data_json["Project Info"]["Main Contact"]["Company"]) != 0:
                    main_contact_list.append(data_json["Project Info"]["Main Contact"]["Company"])
                if len(data_json["Project Info"]["Main Contact"]["Full Name"]) != 0:
                    main_contact_list.append(data_json["Project Info"]["Main Contact"]["Full Name"])
                main_contact_name = "-".join(main_contact_list)

                if task["Main Contact"] != main_contact_name and not (task["Main Contact"] is None and main_contact_name==""):
                    update_body["custom_fields"][custom_field_id_map["Main Contact"]] = main_contact_name
                    update = True

                if task["Contact Type"] is None:
                    data_json["Project Info"]["Main Contact"]["Contact Type"] = "None"
                contact_type = data_json["Project Info"]["Main Contact"]["Contact Type"]
                if task["Contact Type"] != contact_type and not (task["Contact Type"] is None and contact_type=="None"):
                    if contact_type=="None":
                        update_body["custom_fields"][custom_field_id_map["Contact Type"]] = None
                    else:
                        update_body["custom_fields"][custom_field_id_map["Contact Type"]] = contact_id_map[contact_type]
                    update = True

                # if data_json["State"]["Fee Accepted"]:
                #     data_json["Email_Content"] = task["Email_Content"]
                # else:

                if task["Fee ExGST"] is None:
                    task["Fee ExGST"] = 0
                if float(total_amount) != float(task["Fee ExGST"]):
                    update_body["custom_fields"][custom_field_id_map["Fee ExGST"]] = float(total_amount)
                    update = True

                if task["Overdue InGST"] is None:
                    task["Overdue InGST"] = 0
                if float(over_due_amount) != float(task["Overdue InGST"]):
                    update_body["custom_fields"][custom_field_id_map["Overdue InGST"]] = float(over_due_amount)
                    update = True

                if task["Total Paid InGST"] is None:
                    task["Total Paid InGST"] = 0
                if float(paid_amount) != float(task["Total Paid InGST"]):
                    update_body["custom_fields"][custom_field_id_map["Total Paid InGST"]] = float(paid_amount)
                    update = True

                # data_json["State"]["Asana State"] = task["Status"]
                data_json["Invoices"]["Paid Fee"] = str(paid_amount)
                data_json["Invoices"]["Overdue Fee"] = str(over_due_amount)
                mp_json[quotation] = {k: v(data_json) for k, v in mp_convert_map.items()}
                with open(os.path.join(database_dir, quotation, "data.json"), "w") as f:
                    json_object = json.dumps(data_json, indent=4)
                    f.write(json_object)

                if update:
                    asana_update_body = asana.TasksTaskGidBody(update_body)
                    task_api_instance.update_task(task_gid=task["gid"], body=asana_update_body)
                    print(f"Update Asana Task Gid: {task['gid']}")

        if ori_tasks.to_dict()["next_page"] is None:

            mp_json_list = list(mp_json.keys())
            mp_json_list.sort(reverse=True)
            new_mp_json_list = {}
            for quotation in mp_json_list:
                new_mp_json_list[quotation] = mp_json[quotation]
            mp_json = new_mp_json_list

            with open(mp_dir, "w") as f:
                json_object = json.dumps(mp_json, indent=4)
                f.write(json_object)
            print(f"The Daily Update take {time.time() - start}s")
            print("Update Complete!!!")
            break
        off_set = ori_tasks.to_dict()["next_page"]["offset"]
        print(f"Finish Update current page the next off_set_id is {off_set}")
    # status_field = clean_response(
    #     custom_fields_api_instance.get_custom_field(custom_field_id_map["MP"]))
    # status_id_map = name_id_map(status_field["enum_options"])

def convert_to_mp_excel(report_dir):
    mp_convert_map = return_self_mp()
    cur_col = "A"
    report_wb = openpyxl.Workbook()
    report_ws = report_wb.active
    report_ws.title = "report"
    for title in mp_convert_map.keys():
        report_ws[f"{cur_col}1"] = title
        cur_col = increment_excel_column(cur_col)

    mp_dir = os.path.join(database_dir, "mp.json")
    mp_json = json.load(open(mp_dir))
    master_project = list(mp_json.values())
    cur_row = 2
    for project in master_project:
        cur_col = "A"
        for value in project.values():
            report_ws[f"{cur_col}{cur_row}"] = value
            cur_col = increment_excel_column(cur_col)
        cur_row += 1
    report_wb.save(report_dir)
def backup_database():
    start = time.time()
    bridge_dir = os.path.join(working_dir, "Bridge.exe")


    backup_name = os.path.join(backup_dir, date.today().strftime("%Y%m%d"))
    accounting_backup_dir = os.path.join(backup_name, "accounting_backup")
    database_backup_dir = os.path.join(backup_name, "database_backup")
    bridge_backup_dir = os.path.join(backup_name, "Bridge.exe")
    mp_excel_dir = os.path.join(backup_name, "Bridge_MP.xlsx")


    if not os.path.exists(backup_name):
        # os.makedirs(backup_name)
        # os.makedirs(accounting_backup_dir)
        # os.makedirs(database_backup_dir)
        shutil.copytree(accounting_dir, accounting_backup_dir)
        shutil.copytree(database_dir, database_backup_dir)
        shutil.copy(bridge_dir, bridge_backup_dir)
        convert_to_mp_excel(mp_excel_dir)

    past_backup_name = os.path.join(backup_dir, (date.today()-timedelta(30)).strftime("%Y%m%d"))
    if os.path.exists(past_backup_name):
        shutil.rmtree(past_backup_name)
    print(f"The Backup take {time.time() - start}s")
    print("Backup Complete!!!")




if __name__ == '__main__':
    refresh_token()
    update_xero_and_asana_invoice_script()
    update_asana_project_script()
    backup_database()