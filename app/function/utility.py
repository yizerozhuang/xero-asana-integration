import asana
from win32com import client as win32client
from xero_python import *
import shutil
import os
import requests
import urllib
import webbrowser
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options

#asana part
asana_configuration = asana.Configuration()
asana_configuration.access_token = '1/1205463377927189:4825d7e7924a9dd8dd44a9c826e45455'
asana_api_client = asana.ApiClient(asana_configuration)

project_api_instance = asana.ProjectsApi(asana_api_client)
task_api_instance = asana.TasksApi(asana_api_client)
custom_fields_api_instance = asana.CustomFieldsApi(asana_api_client)
template_api_instance = asana.TaskTemplatesApi(asana_api_client)
workspace_gid = '1205045058713243'
#xero part
from functools import wraps
from io import BytesIO
from logging.config import dictConfig

from xero_python.accounting import AccountingApi, ContactPerson, Contact, Contacts
from xero_python.api_client import ApiClient, serialize
from xero_python.api_client.configuration import Configuration
from xero_python.api_client.oauth2 import OAuth2Token
from flask import Flask
from xero_python.exceptions import AccountingBadRequestException
from xero_python.identity import IdentityApi
from xero_python.utils import getvalue
import pkce
import json
import http.server
import socketserver
import _thread
from typing import Tuple
from http import HTTPStatus

client_id = "7BC59213098143A68B6ED1DD08EE16BA"
redirect_url = "http://localhost:5000/get_response"
code_verifier, code_challenge = pkce.generate_pkce_pair()
xero_api_client = ApiClient(
    Configuration(
        debug=False,
        oauth2_token=OAuth2Token(
            client_id=client_id
        ),
    ),
    pool_threads=1,
)

def remove_none(obj):
    if isinstance(obj, list):
        return [remove_none(x) for x in obj if x is not None]
    elif isinstance(obj, dict):
        return {k: remove_none(v) for k, v in obj.items() if v is not None}
    else:
        return obj


def clearn_response(response):
    response = response.to_dict()["data"]
    return remove_none(response)


def name_id_map(api_list):
    res = dict()
    for item in api_list:
        res[item["name"]] = item["gid"]
    return res


def update_asana(data, *args):
    all_projects = clearn_response(project_api_instance.get_projects_for_workspace(workspace_gid))
    projects_id_map = name_id_map(all_projects)

    current_project_template = clearn_response(template_api_instance.get_task_templates(
        project=projects_id_map[data["Project Information"]["Project Type"].get()]))
    template_id_map = name_id_map(current_project_template)

    api_respond = clearn_response(template_api_instance.instantiate_task(template_id_map["P:\\300000-XXXX"]))
    task_id = api_respond["new_task"]["gid"]

    all_custom_fields = clearn_response(task_api_instance.get_task(task_id))["custom_fields"]
    custom_field_id_map = name_id_map(all_custom_fields)

    status_field = clearn_response(custom_fields_api_instance.get_custom_field(custom_field_id_map["Status"]))
    status_id_map = name_id_map(status_field["enum_options"])

    service_filed = clearn_response(custom_fields_api_instance.get_custom_field(custom_field_id_map["Services"]))
    service_id_map = name_id_map(service_filed["enum_options"])

    contact_field = clearn_response(custom_fields_api_instance.get_custom_field(custom_field_id_map["Contact Type"]))
    contact_id_map = name_id_map(contact_field["enum_options"])

    body = asana.TasksTaskGidBody(
        {
            "name": data["Project Information"]["Quotation Number"].get() + "-" + data["Project Information"][
                "Project Name"].get(),
            "custom_fields": {
                custom_field_id_map["Status"]: status_id_map["Fee Proposal"],
                custom_field_id_map["Services"]: [service_id_map[service._name] for service in
                                                  data["Project Information"]["Service Type"] if service.get()],
                custom_field_id_map["Shop name"]: data["Project Information"]["Project Name"].get(),
                custom_field_id_map["Apt/Room/Area"]: data["Building Features"]["Total Area"].get() + "m2",
                custom_field_id_map["Feature/Notes"]: data["Building Features"]["Feature/Notes"].get(),
                custom_field_id_map["Client"]: data["Client"]["Client Full Name"].get(),
                custom_field_id_map["Main contact"]: data["Main Contact"]["Main Contact Full Name"].get(),
                custom_field_id_map["Contact Type"]: contact_id_map[
                    data["Main Contact"]["Main Contact Contact Type"].get()]
            }
        }
    )
    task_api_instance.update_task(task_gid=task_id, body=body)
    print(api_respond)




def excel_print_pdf(data, *args):
    if _use_page_2(data):
        shutil.copyfile(os.getcwd() + "\\resource\\xlsx\\fee proposal template 2pages.xlsx",
                        os.getcwd() + "\\test\\fee proposal template.xlsx")
    else:
        shutil.copyfile(os.getcwd() + "\\resource\\xlsx\\fee proposal template.xlsx",
                        os.getcwd() + "\\test\\fee proposal template.xlsx")

    excel = win32client.Dispatch("Excel.Application")
    sheets = excel.Workbooks.Open(os.getcwd() + "\\test\\fee proposal template.xlsx")
    work_sheets = sheets.Worksheets[0]
    work_sheets.Cells(1, 2).Value = data["Client"]["Client Full Name"].get()
    work_sheets.Cells(2, 2).Value = data["Client"]["Client Company"].get()
    work_sheets.Cells(1, 8).Value = data["Project Information"]["Quotation Number"].get()
    work_sheets.Cells(2, 8).Value = data["Fee Proposal Page"]["Date"].get()
    work_sheets.Cells(3, 8).Value = data["Fee Proposal Page"]["Revision"].get()
    work_sheets.Cells(6, 1).Value = "Re: " + data["Project Information"]["Project Name"].get()
    work_sheets.Cells(8,
                      1).Value = f"Thank you for giving us the opportunity to submit this fee proposal for our {', '.join([key for key, value in data['Fee Proposal Page']['Scope of Work'].items() if value['on']])}for the above project."
    work_sheets.Cells(16, 7).Value = data["Fee Proposal Page"]["Time"]["Pre-design information collection Duration"][0].get() + "-" + data["Fee Proposal Page"]["Time"]["Fee proposal stage Duration"][1].get()
    work_sheets.Cells(21, 7).Value = data["Fee Proposal Page"]["Time"]["Pre-design information collection Duration"][0].get() + "-" + data["Fee Proposal Page"]["Time"]["Pre-design information collection Duration"][1].get()
    work_sheets.Cells(26, 7).Value = data["Fee Proposal Page"]["Time"]["Documentation and issue Duration"][0].get() + "-" + data["Fee Proposal Page"]["Time"]["Documentation and issue Duration"][1].get()

    cur_row = 52
    cur_index = 1
    to_page = 52
    for key, service in data['Fee Proposal Page']['Scope of Work'].items():
        if not service["on"]:
            continue
        else:
            extra_list = ["Extend", "Exclusion", "Deliverables"]
            for extra in extra_list:
                work_sheets.Cells(cur_row, 1).Value = "2."+str(cur_index)
                work_sheets.Cells(cur_row, 2).Value = key+"-"+extra
                work_sheets.Cells(cur_row, 1).Font.Bold = True
                work_sheets.Cells(cur_row, 2).Font.Bold = True
                for scope in service[extra]:
                    if scope[0].get():
                        cur_row += 1
                        work_sheets.Cells(cur_row, 1).Value = "•"
                        work_sheets.Cells(cur_row, 2).Value = scope[1].get()
                cur_row += 2
                cur_index+=1
            cur_row+=1
    work_sheets.ExportAsFixedFormat(0, os.getcwd() + "\\test\\excel_output.pdf")
    sheets.Close(True)

    # Client_Full_Name=self.data["Client"]["Client Full Name"].get(),
    # Client_Company=self.data["Client"]["Client Company"].get(),
    # Reference=self.data["Project Information"]["Quotation Number"].get(),
    # Date=self.data["Fee Proposal Page"]["Date"].get(),
    # Revision=self.data["Fee Proposal Page"]["Revision"].get(),
    # Project_Name=self.data["Project Information"]["Project Name"].get(),
    # Fee_Stage_Start=self.data["Fee Proposal Page"]["Time"]["Fee proposal stage Duration"][0].get(),
    # Fee_Stage_End=self.data["Fee Proposal Page"]["Time"]["Fee proposal stage Duration"][1].get(),
    # Design_Stage_Start=self.data["Fee Proposal Page"]["Time"]["Pre-design information collection Duration"][
    #     0].get(),
    # Design_Stage_End=self.data["Fee Proposal Page"]["Time"]["Pre-design information collection Duration"][
    #     1].get(),
    # Issue_Stage_Start=self.data["Fee Proposal Page"]["Time"]["Documentation and issue Duration"][0].get(),
    # Issue_Stage_End=self.data["Fee Proposal Page"]["Time"]["Documentation and issue Duration"][1].get(),
    # Scope_of_Work=self._format_scope(),
    # Past_Project=self._format_experience(),
    # Details=self._format_details()

def _use_page_2(data):
    total_scopt_count = 0
    for service in data['Fee Proposal Page']['Scope of Work'].values():
        if not service["on"]:
            return
        else:
            # three extra and six space
            total_scopt_count+=9
            extra_list = ["Extend", "Exclusion", "Deliverables"]
            for extra in extra_list:
                for scope in service[extra]:
                    if scope[0].get():
                        total_scopt_count+=1

    if total_scopt_count >= 30:
        return True
    else:
        return False
def email(data, *args):
    ol = win32client.Dispatch("Outlook.Application")
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = "Fee Proposal - " + data["Project Information"]["Project Name"].get()
    newmail.To = data["Client"]["Contact Email"].get()
    ###check if possible cc multiple email
    newmail.CC = "felix@pcen.com.au"
    newmail.GetInspector()
    index = newmail.HTMLbody.find(">", newmail.HTMLbody.find("<body"))
    message = f"""
    Dear {data["Client"]["Client Full Name"].get()},<br>

    I hope this email finds you well. Please find the attached fee proposal to this email.<br>

    If you have any questions or need more information regarding the proposal, please don't hesitate to reach out. Felix and I are happy to provide you with whatever information you need.<br>

    I look forward to hearing from you soon.<br>

    Cheers,<br>
    """
    newmail.HTMLbody = newmail.HTMLbody[:index + 1] + message + newmail.HTMLbody[index + 1:]
    newmail.Attachments.Add(os.getcwd() + "\\output.pdf")
    newmail.Display()


def update_xero(data, invoice_number):
    url = f"http://login.xero.com/identity/connect/authorize?response_type=code&client_id={client_id}&redirect_uri={redirect_url}&scope=openid profile email accounting.transactions&state=123&code_challenge={code_challenge}&code_challenge_method=S256"
    url = url.replace(" ", "%20")
    # headers = {
    #     'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.80 Safari/537.36',
    #     'Content-Type': 'text/html',
    # }
    # response = requests.get(url, headers=headers)
    # html = response.text
    def thread_task(threadName):
        start_flask()
    _thread.start_new_thread(thread_task, ("flask",))
    webbrowser.open_new_tab(url)

    # chrome_options = Options()
    # chrome_options.add_argument("--headless")
    # with Chrome(options=chrome_options) as browser:
    #     browser.get(url)
    print()
    # with urllib.request.urlopen(url) as responce:
    #     print(responce)
    # print()
    #
    # code = "RuQQjqRLjgdvdVAUCd5Pl_AlhlK55laytrnKUyQnUN4"
    # code_verifier='MRAt_5M3PhjDAAGFU1Cd0PokTEjncJue2KsszWLwURhDsC6C45jcEbL0aztnAW5NNkwcRte_KgrXKVfq9fPSdg9NCb3Pq2kH2pWj2Jg0h429-OhtBGy59xgA2fTbDXqp'
    # exchange_url = "https://identity.xero.com/connect/token"
    # responce = requests.post(exchange_url,
    #                          headers={"Content-Type":"application/x-www-form-urlencoded"},
    #                          data={
    #                              "grant_type":"authorization_code",
    #                              "client_id":"7BC59213098143A68B6ED1DD08EE16BA",
    #                              "code":code,
    #                              "redirect_url": redirect_url,
    #                              "code_verifier":code_verifier
    #                          })
    # xero_tenant_id = get_xero_tenant_id()
    # accounting_api = AccountingApi(xero_api_client)
    #
    # invoices = accounting_api.get_invoices(
    #     xero_tenant_id, statuses=["DRAFT", "SUBMITTED"]
    # )
    # code = serialize_model(invoices)
    # sub_title = "Total invoices found: {}".format(len(invoices.invoices))
    #
    # return render_template(
    #     "code.html", title="Invoices", code=code, sub_title=sub_title
    # )

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, request:bytes, client_address, server):
        super().__init__(request, client_address, server)
    @property
    def api_response(self):
        return json.dumps({"message":"Hello World"}).encode()
    def do_GET(self) -> None:
        if self.path == "/get_response":
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(bytes(self.api_response))

def start_flask():
    PORT = 5000
    my_server = socketserver.TCPServer(("localhost", PORT), Handler)
    print("my_server start")
    my_server.serve_forever()
