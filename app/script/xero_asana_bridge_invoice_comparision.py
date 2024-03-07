from xero_python.accounting import AccountingApi
import asana
import dateutil.parser
from config import CONFIGURATION as conf
from utility import generate_all_project, isfloat
from xero_function import xero_token_required, refresh_token, get_xero_tenant_id, api_client
from asana_function import clean_response, name_id_map, flatter_custom_fields

import os
import json
import time
import shutil
from datetime import date
from win32com import client as win32client


# database_dir = conf["database_dir"]
database_dir = "P:\\app\\database"
invoice_dir = os.path.join(database_dir, "invoices.json")
invoice_json = json.load(open(invoice_dir))
bill_dir = os.path.join(database_dir, "bills.json")
bill_json = json.load(open(bill_dir))

# generate_management_report = False

if_modified_since = dateutil.parser.parse("2021-01-01")
start_search_date = date(2021, 1, 1)


all_project = generate_all_project(database_dir)


xero_fields = ["invoice_id", "amount_credited", "amount_due", "amount_paid", "date", "due_date", "fully_paid_on_date", "status", "sub_total", "total", "total_tax"]
asana_fields = []
bridge_fields = []

def create_invoice_with_invoice_number(all_invoices, invoice_number):
    assert not invoice_number in list(all_invoices.keys())
    all_invoices[invoice_number] = {
        "Invoice Number": invoice_number,
        "Xero": {
            key: "" for key in xero_fields
        }
    }

def process_xero_invoices(all_invoices):
    xero_tenant_id = get_xero_tenant_id()
    accounting_api = AccountingApi(api_client)
    xero_invoices = accounting_api.get_invoices(xero_tenant_id, if_modified_since=if_modified_since, include_archived="False").to_dict()["invoices"]
    for invoice in xero_invoices:
        if invoice["date"] < start_search_date:
            continue
        if invoice["type"] == "ACCREC":
            if invoice["status"] == "DELETED":
                continue
            create_invoice_with_invoice_number(all_invoices, invoice["invoice_number"])
            all_invoices[invoice["invoice_number"]]["Xero"] = {
                key:invoice[key] for key in xero_fields
            }

def process_asana_invoices(all_invoices):
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
    off_set = None
    opt_field = ["name", "completed", "custom_fields", "due_at", "due_on", "parent"]
    while True:
        if off_set is None:
            ori_tasks = clean_response(task_api_instance.get_tasks_for_project(project_gid=invoice_status_asana_gid, opt_fields=opt_field, limit=100))
        else:
            ori_tasks = clean_response(task_api_instance.get_tasks_for_project(project_gid=invoice_status_asana_gid, opt_fields=opt_field, limit=100, offset=off_set))
        for task in ori_tasks:
            task = flatter_custom_fields(task)
            #backlog use asana gid as invoice number
            print()
def process_bridge_invoices(all_invoices):
    pass

@xero_token_required
def compare_xero_asana_bridge_invoices():
    start = time.time()
    all_invoices = {}
    process_xero_invoices(all_invoices)
    process_asana_invoices(all_invoices)
    print()
    # for invoices_states in invoices["INV"].keys():
    #     for invoice_number in invoices["INV"][invoices_states].keys():
    #         invoice_json[invoice_number] = invoices_states
    #
    # for key, value in invoices["BIL"].items():
    #     for inv_number in value:
    #         bill_json[inv_number] = key
    # off_set = None
    # opt_field = ["name", "custom_fields.name", "custom_fields.display_value"]
    #
    # custom_fields_setting = clean_response(custom_fields_setting_api_instance.get_custom_field_settings_for_project(invoice_status_asana_gid))
    # all_custom_fields = [custom_field["custom_field"] for custom_field in custom_fields_setting]
    # custom_field_id_map = name_id_map(all_custom_fields)
    #
    # status_field = clean_response(
    #     custom_fields_api_instance.get_custom_field(custom_field_id_map["Invoice status"]))
    # status_id_map = name_id_map(status_field["enum_options"])

if __name__ == '__main__':
    refresh_token()
    compare_xero_asana_bridge_invoices()
    # update_asana_project_script()