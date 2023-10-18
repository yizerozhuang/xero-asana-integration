import asana
from win32com import client as win32client
import shutil
import os

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
import datetime
from functools import wraps

from flask import Flask, url_for, session, redirect, json, request
from flask_oauthlib.contrib.client import OAuth, OAuth2Application
from flask_session import Session
from xero_python.accounting import AccountingApi, Account, AccountType, Contact, LineItem, Invoice, Invoices
from xero_python.api_client import ApiClient
from xero_python.api_client.configuration import Configuration
from xero_python.api_client.oauth2 import OAuth2Token
from xero_python.exceptions import AccountingBadRequestException
from xero_python.identity import IdentityApi
from xero_python.utils import getvalue

import _thread
import webbrowser
from datetime import datetime


# configure main flask application
app = Flask(__name__)
app.config.from_object("default_settings")
app.config.from_pyfile("config.py", silent=True)

if app.config["ENV"] != "production":
    # allow oauth2 loop to run over http (used for local testing only)
    os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"

# configure persistent session cache
Session(app)

# configure flask-oauthlib application
oauth = OAuth(app)
xero = oauth.remote_app(
    name="xero",
    version="2",
    client_id=app.config["CLIENT_ID"],
    client_secret=app.config["CLIENT_SECRET"],
    endpoint_url="https://api.xero.com/",
    authorization_url="https://login.xero.com/identity/connect/authorize",
    access_token_url="https://identity.xero.com/connect/token",
    refresh_token_url="https://identity.xero.com/connect/token",
    scope="offline_access openid profile email accounting.transactions "
    "accounting.transactions.read accounting.reports.read "
    "accounting.journals.read accounting.settings accounting.settings.read "
    "accounting.contacts accounting.contacts.read accounting.attachments "
    "accounting.attachments.read assets projects "
    "files "
    "payroll.employees payroll.payruns payroll.payslip payroll.timesheets payroll.settings",
    # "paymentservices "
    # "finance.bankstatementsplus.read finance.cashvalidation.read finance.statements.read finance.accountingactivity.read",
)  # type: OAuth2Application


# configure xero-python sdk client
api_client = ApiClient(
    Configuration(
        debug=app.config["DEBUG"],
        oauth2_token=OAuth2Token(
            client_id=app.config["CLIENT_ID"], client_secret=app.config["CLIENT_SECRET"]
        ),
    ),
    pool_threads=1,
)

def login_xero():
    def thread_task():
        start_flask()
    _thread.start_new_thread(thread_task,())
    webbrowser.open_new_tab("http://localhost:1234/login")

def start_flask():
    app.run(host='localhost', port=1234)
@app.route("/login")
def login():
    redirect_url = url_for("oauth_callback", _external=True)
    session["state"] = app.config["STATE"]
    response = xero.authorize(callback_uri=redirect_url, state=session["state"])
    return response

@app.route("/callback")
def oauth_callback():
    if request.args.get("state") != session["state"]:
        return "Error, state doesn't match, no token for you."
    try:
        response = xero.authorized_response()
    except Exception as e:
        print(e)
        raise
    if response is None or response.get("access_token") is None:
        return "Access denied: response=%s" % response
    store_xero_oauth2_token(response)
    return "You are Successfully login, you can go back to the app right now"
token_list = {}
@xero.tokengetter
@api_client.oauth2_token_getter
def obtain_xero_oauth2_token():
    return token_list.get("token")

@xero.tokensaver
@api_client.oauth2_token_saver
def store_xero_oauth2_token(token):
    token_list["token"] = token

def xero_token_required(function):
    @wraps(function)
    def decorator(*args, **kwargs):
        xero_token = obtain_xero_oauth2_token()
        if not xero_token:
            return redirect(url_for("login", _external=True))

        return function(*args, **kwargs)

    return decorator
def get_code_snippet(endpoint,action):
    s = open("C:\\Users\\yeezh\\xero-python-oauth2-app\\app.py").read()
    startstr = "["+ endpoint +":"+ action +"]"
    endstr = "#[/"+ endpoint +":"+ action +"]"
    start = s.find(startstr) + len(startstr)
    end = s.find(endstr)
    substring = s[start:end]
    return substring
def get_xero_tenant_id():
    token = obtain_xero_oauth2_token()
    if not token:
        return None

    identity_api = IdentityApi(api_client)
    for connection in identity_api.get_connections():
        if connection.tenant_type == "ORGANISATION":
            return connection.tenant_id
@xero_token_required
def update_xero(data, invoice):
    xero_tenant_id = get_xero_tenant_id()
    accounting_api = AccountingApi(api_client)
    # try:
    # we need a contact
    contacts = accounting_api.get_contacts(
        xero_tenant_id
    )
    contacts_list = remove_none(contacts.to_dict())["contacts"]
    def contact_name_contact_id(api_list):
        res = dict()
        for item in api_list:
            res[item["name"]] = item["contact_id"]
        return res
    # need to handle if it's not in the client
    contacts_name_map = contact_name_contact_id(contacts_list)
    if len(data["Client"]["Client Company"].get()) == 0:
        contacts_key = data["Client"]["Client Full Name"].get()
    elif len(data["Client"]["Client Full Name"].get()) == 0:
        contacts_key = data["Client"]["Client Company"].get()
    else:
        contacts_key = data["Client"]["Client Company"].get()+", "+data["Client"]["Client Full Name"].get()
    contact_id = contacts_name_map[contacts_key] if contacts_key in contacts_name_map.keys() else create_contact(contacts_key)
    contact = Contact(contact_id)

    # we need an account of type BANK
    where = "Type==\"BANK\""
    try:
        accounts = accounting_api.get_accounts(
            xero_tenant_id, where
        )
        account_id = getvalue(accounts, "accounts.0.account_id", "")
        account = Account(account_id=account_id)
    except AccountingBadRequestException as exception:
        pass
    #only first one
    line_item_list = []
    for key, service in data["Fee Proposal Page"]["Details"].items():
        if not service["on"].get():
            continue
        if service["Expanded"].get():
            for i in range(3):
                if service["Context"][i]["Invoice"].get() == "INV1":
                    try:
                        line_item_list.append(
                            LineItem(
                                description=service["Context"][i]["Service"].get(),
                                quantity=1,
                                unit_amount=int(service["Context"][i]["Fee"].get()),
                                account_code="200"
                            )
                        )
                    except ValueError:
                        continue
        else:
            if service["Invoice"].get() == "INV1":
                try:
                    line_item_list.append(
                        LineItem(
                            description=key,
                            quantity=1,
                            unit_amount=int(service["Fee"].get()),
                            account_code="200"
                        )
                    )
                except ValueError:
                    continue
    # we need multiple invoices

    invoice_1 = Invoice(
        type="ACCREC",
        contact=contact,
        date=datetime.today(),
        due_date=datetime.today(),
        line_items=line_item_list,
        invoice_number=invoice[0]["Invoice Number"].get(),
        reference=data["Project Information"]["Project Name"].get(),
        status="DRAFT"
    )
    invoices = Invoices(invoices=[invoice_1])

    try:
        created_invoices = accounting_api.create_invoices(
            xero_tenant_id, invoices
        )
        invoice_1_id = getvalue(created_invoices, "invoices.0.invoice_id", "")
    except AccountingBadRequestException as exception:
        pass

    #[BATCHPAYMENTS:CREATE]
    # xero_tenant_id = get_xero_tenant_id()
    # accounting_api = AccountingApi(api_client)
    #
    # invoice_1 = Invoice(invoice_id=invoice_1_id)
    # invoice_2 = Invoice(invoice_id=invoice_2_id)
    #
    # payment_1 = Payment(
    #     reference="something 1",
    #     invoice=invoice_1,
    #     amount=3.50
    # )
    # payment_2 = Payment(
    #     reference="something 2",
    #     invoice=invoice_2,
    #     amount=7.25
    # )
    #
    # batch_payment = BatchPayment(
    #     date=dateutil.parser.parse("2020-12-24"),
    #     reference="Something",
    #     account=account,
    #     payments=[payment_1, payment_2]
    # )
    #
    # batch_payments = BatchPayments(batch_payments=[batch_payment])
    #
    # try:
    #     created_batch_payments = accounting_api.create_batch_payment(
    #         xero_tenant_id, batch_payments
    #     )
    # except AccountingBadRequestException as exception:
    #     output = "Error: " + exception.reason
    #     json = jsonify(exception.error_data)
    # else:
    #     output = "Batch payment created with id {} .".format(
    #         getvalue(created_batch_payments, "batch_payments.0.batch_payment_id", "")
    #     )
    #     json = serialize_model(created_batch_payments)

@xero_token_required
def create_contact(name):
    xero_tenant_id = get_xero_tenant_id()
    accounting_api = AccountingApi(api_client)

    account = Account(
        name=name,
        type=AccountType.EXPENSE
    )
    api_response = accounting_api.create_account(xero_tenant_id, account)
    print(api_response)
    return api_response.to_dict()["accounts"][0]["account_id"]

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

def read_response(response):
    response_content = response.content.decode("utf8")
    response_json = json.loads(response_content)
    return response_json


# def update_xero(data, invoice_number):
#     if not token_make:
#         url = f"http://login.xero.com/identity/connect/authorize?response_type=code&client_id={client_id}&redirect_uri={redirect_url}&scope=openid profile email accounting.transactions&state=123&code_challenge={code_challenge}&code_challenge_method=S256"
#         url = url.replace(" ", "%20")
#         query_string = [None]
#         def thread_task(result):
#             start_flask(result)
#         _thread.start_new_thread(thread_task, (query_string,))
#         webbrowser.open_new_tab(url)
#         while True:
#             time.sleep(1)
#             if not query_string[0] is None:
#                 break
#         code = str(query_string[0]).split("&")[0].split("=")[1]
#         exchange_url =  "https://identity.xero.com/connect/token"
#         response = requests.post(exchange_url,
#                                  headers={
#                                      "Content-Type":"application/x-www-form-urlencoded"
#                                  },
#                                  data={
#                                      "grant_type":"authorization_code",
#                                      "client_id":client_id,
#                                      "code":code,
#                                      "redirect_uri":redirect_url,
#                                      "code_verifier":code_verifier
#                                  })
#         access_token = read_response(response)["access_token"]
#         xero_api_client.oauth2_token_saver(access_token)
#         tenant_response = requests.get("https://api.xero.com/connections",
#                                 headers = {
#                                     "Authorization":"Bearer "+ access_token,
#                                     "Content-Type":"application/json"
#                                 })
#         #need to change to choose the oorganisation
#         tenantId = read_response(tenant_response)[0]["tenantId"]
#         xero_api_client.set_oauth2_token(access_token)
#         api_instance = AccountingApi(xero_api_client)
#         xero_tenant_id = tenantId
#         date_value = date_parser.parse('2020-10-10T00:00:00Z')
#         due_date_value = date_parser.parse('2020-10-28T00:00:00Z')
#
#         contact = Contact(
#             contact_id="00000000-0000-0000-0000-000000000000")
#
#         line_item_tracking = LineItemTracking(
#             tracking_category_id="00000000-0000-0000-0000-000000000000",
#             tracking_option_id="00000000-0000-0000-0000-000000000000")
#
#         line_item_trackings = []
#         line_item_trackings.append(line_item_tracking)
#
#         line_item = LineItem(
#             description="Foobar",
#             quantity=1.0,
#             unit_amount=20.0,
#             account_code="000",
#             tracking=line_item_trackings)
#
#         line_items = []
#         line_items.append(line_item)
#
#         invoice = Invoice(
#             type="ACCREC",
#             contact=contact,
#             date=date_value,
#             due_date=due_date_value,
#             line_items=line_items,
#             reference="Website Design",
#             status="DRAFT")
#
#         invoices = Invoices(
#             invoices=[invoice])
#
#         try:
#             api_response = api_instance.create_invoices(xero_tenant_id, invoices)
#             print(api_response)
#         except AccountingBadRequestException as e:
#             print("Exception when calling AccountingApi->createInvoices: %s\n" % e)


    #

    # chrome_options = Options()
    # chrome_options.add_argument("--headless")
    # with Chrome(options=chrome_options) as browser:
    #     browser.get(url)
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

# http request version
# class Handler(http.server.SimpleHTTPRequestHandler):
#     def __init__(self, request:bytes, client_address, server):
#         super().__init__(request, client_address, server)
#     @property
#     def api_response(self):
#         return json.dumps({"message":"Hello World"}).encode()
#     def do_GET(self) -> None:
#         if self.path == "/get_response":
#             self.send_response(HTTPStatus.OK)
#             self.send_header("Content-Type", "application/json")
#             self.end_headers()
#             self.wfile.write(bytes(self.api_response))
#
# def start_flask():
#     PORT = 6789
#     my_server = socketserver.TCPServer(("localhost", PORT), Handler)
#     print("my_server start")
#     my_server.serve_forever()

#flask version
# def start_flask(result):
#     PORT = 6789
#     app = Flask(__name__)
#     @app.route("/get_response", methods=["GET"])
#     def get_response():
#         result[0] = request.query_string
#         return "The Authorization is successful, you can close this tab and return to the app"
#     app.run(port=PORT)


