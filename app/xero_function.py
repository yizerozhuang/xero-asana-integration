from xero_python.accounting import AccountingApi, Contact, Contacts, LineItem, Invoice, Invoices
from xero_python.api_client import ApiClient
from xero_python.api_client.configuration import Configuration
from xero_python.api_client.oauth2 import OAuth2Token
from xero_python.identity import IdentityApi
from flask import Flask, url_for, redirect
from flask_oauthlib.contrib.client import OAuth, OAuth2Application
from flask_session import Session


from utility import remove_none, update_app_invoices, get_invoice_item
from asana_function import update_asana_invoices
from config import CONFIGURATION as conf

import _thread
from datetime import datetime
import webbrowser
from functools import wraps
import os
import base64
import requests


# configure main flask application
flask_app = Flask(__name__)
flask_app.config["SECRET_KEY"] = os.urandom(16)

# configure file based session
flask_app.config["SESSION_TYPE"] = "filesystem"
# SESSION_FILE_DIR = join(dirname(__file__), "../../xero-python-oauth2-starter/cache")

# configure flask app for local development
flask_app.config["ENV"] = "development"

flask_app.config["CLIENT_ID"] = conf["xero_client_id"]
flask_app.config["CLIENT_SECRET"] = conf["xero_client_secret"]

# if flask_app.config["ENV"] != "production":
#     # allow oauth2 loop to run over http (used for local testing only)
os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"

Session(flask_app)

scope = [
    "offline_access",
    "openid",
    "profile",
    "email",
    "accounting.transactions",
    "accounting.transactions.read",
    "accounting.reports.read",
    "accounting.journals.read",
    "accounting.settings",
    "accounting.settings.read",
    "accounting.contacts",
    "accounting.contacts.read",
    "accounting.attachments",
    "accounting.attachments.read",
    "assets",
    "projects",
    "files",
    "payroll.employees",
    "payroll.payruns",
    "payroll.payslip",
    "payroll.timesheets",
    "payroll.settings"
]
# configure flask-oauthlib application
oauth = OAuth(flask_app)
xero = oauth.remote_app(
    name="xero",
    version="2",
    client_id=flask_app.config["CLIENT_ID"],
    client_secret=flask_app.config["CLIENT_SECRET"],
    endpoint_url="https://api.xero.com/",
    authorization_url="https://login.xero.com/identity/connect/authorize",
    access_token_url="https://identity.xero.com/connect/token",
    refresh_token_url="https://identity.xero.com/conneh_ct/token",
    scope=" ".join(scope)

)  # type: OAuth2Application


# configure xero-python sdk client
api_client = ApiClient(
    Configuration(
        debug=flask_app.config["DEBUG"],
        oauth2_token=OAuth2Token(
            client_id=flask_app.config["CLIENT_ID"], client_secret=flask_app.config["CLIENT_SECRET"]
        ),
    ),
    pool_threads=1,
)

def login_xero():
    def thread_task():
        start_flask()
    _thread.start_new_thread(thread_task, ())
    webbrowser.open_new_tab("http://localhost:1234/login")


def start_flask():
    flask_app.run(host='localhost', port=1234)

@flask_app.route("/login")
def login():
    redirect_url = url_for("oauth_callback", _external=True)
    # session["state"] = flask_app.config["STATE"]
    response = xero.authorize(callback_uri=redirect_url)
    return response

@flask_app.route("/callback")
def oauth_callback():
    # if request.args.get("state") != session["state"]:
    #     return "Error, state doesn't match, no token for you."
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
def get_xero_tenant_id():
    token = obtain_xero_oauth2_token()
    if not token:
        return None

    identity_api = IdentityApi(api_client)
    for connection in identity_api.get_connections():
        if connection.tenant_type == "ORGANISATION":
            return connection.tenant_id

def contact_name_contact_id(api_list):
    res = dict()
    for item in api_list:
        res[item["name"]] = item["contact_id"]
    return res
def invoice_number_invoice_id(api_list):
    res = dict()
    for item in api_list:
        res[item["invoice_number"]] = item["invoice_id"]
    return res
def _process_invoices(inv_list):
    res = {
        "Invoices": {
            "DRAFT": {},
            "SUBMITTED": {},
            "DELETED": {},
            "AUTHORISED": {},
            "PAID": {},
            "VOIDED": {},
            "NONE": {}
        },
        "Bills": {
            "DRAFT": {},
            "SUBMITTED": {},
            "DELETED": {},
            "AUTHORISED": {},
            "PAID": {},
            "VOIDED": {},
            "NONE": {}
        }
    }
    for inv in inv_list:
        if inv["type"] == "ACCREC":
            res["Invoices"][inv["status"]][inv["invoice_number"]] ={
                "sub_total": inv["sub_total"],
                "line_amount_types": inv["line_amount_types"].value
            }
        elif inv["type"] == "ACCPAY":
            res["Bills"][inv["status"]][inv["invoice_number"]] = {
                "sub_total": inv["sub_total"],
                "line_amount_types": inv["line_amount_types"].value
            }
    return res
def refresh_token():
    refresh_url = "https://identity.xero.com/connect/token"

    old_refresh_token = open("P:\\app\\xero_refresh_token.txt", 'r').read()

    tokenb4 = f"{conf['xero_client_id']}:{conf['xero_client_secret']}"
    basic_token = base64.urlsafe_b64encode(tokenb4.encode()).decode()

    headers = {
      'Authorization': f"Basic {basic_token}",
      'Content-Type': 'application/x-www-form-urlencoded',
    }
    data = {
        'grant_type': 'refresh_token',
        'refresh_token': old_refresh_token
    }

    try:
        response = requests.post(refresh_url, headers=headers, data=data)

        results = response.json()
        open("P:\\app\\xero_refresh_token.txt", 'w').write(results["refresh_token"])
        open("P:\\app\\xero_access_token.txt", 'w').write(results["access_token"])
        # store_xero_oauth2_token(response)
        store_xero_oauth2_token(
            {
                "access_token": results["access_token"],
                "refresh_token": results["refresh_token"],
                "expires_in": 1800,
                "token_type": "Bearer",
                "scope": scope
            }
        )
    except Exception as e:
        print("ERROR ! Refreshing token error?")
        print(response.text)

@xero_token_required
def update_xero(app, contact_name):
    data = app.data
    xero_tenant_id = get_xero_tenant_id()
    accounting_api = AccountingApi(api_client)
    contacts = accounting_api.get_contacts(
        xero_tenant_id
    )
    contacts_list = remove_none(contacts.to_dict())["contacts"]
    # need to handle if it's not in the client
    contacts_name_map = contact_name_contact_id(contacts_list)
    if contact_name in contacts_name_map.keys():
        contact_id = contacts_name_map[contact_name]
        contact = Contact(contact_id)
    else:
        contact = Contact(create_contact(app, contact_name))

    # we need an account of type BANK
    # where = "Type==\"BANK\""
    # try:
    #     accounts = accounting_api.get_accounts(
    #         xero_tenant_id, where
    #     )
    #     account_id = getvalue(accounts, "accounts.0.account_id", "")
    #     account = Account(account_id=account_id)
    # except AccountingBadRequestException as exception:
    #     pass

    # all_invoices = remove_none(accounting_api.get_invoices(xero_tenant_id).to_dict())["invoices"]
    # invoice_number_map = invoice_number_invoice_id(all_invoices)


    # submitted_where = "Status==SUBMITTED"
    #
    # submitted_invoice = remove_none(accounting_api.get_invoices(xero_tenant_id, where=submitted_where).to_dict())["invoices"]
    # submitted_invoice_id = invoice_number_invoice_id(submitted_invoice)
    #
    # approve_where = 'Status==AUTHORISED'
    #
    # approve_invoice = remove_none(accounting_api.get_invoices(xero_tenant_id, where=approve_where).to_dict())["invoices"]
    # approve_invoice_id = invoice_number_invoice_id(approve_invoice)
    #
    paid_where = 'Status=="PAID"'

    paid_invoice = remove_none(accounting_api.get_invoices(xero_tenant_id, where=paid_where).to_dict())["invoices"]

    paid_invoice_id = invoice_number_invoice_id(paid_invoice)

    invoice_item = get_invoice_item(app)

    invoices_list = []


    for i in range(6):
        if len(data["Invoices Number"][i]["Number"].get()) == 0 or data["Invoices Number"][i]["Number"].get() in paid_invoice_id.keys() or len(invoice_item[i]) == 0:
            continue
        line_item_list = []
        for item in invoice_item[i]:
            line_item_list.append(
                LineItem(
                    description=item["Item"],
                    quantity=1,
                    unit_amount=int(float(item["Fee"])),
                    account_code="000"
                )
            )
        invoices_list.append(
            Invoice(
                type="ACCREC",
                contact=contact,
                date=datetime.today(),
                due_date=datetime.today(),
                line_items=line_item_list,
                invoice_number=data["Invoices Number"][i]["Number"].get(),
                reference=data["Project Info"]["Project"]["Project Name"].get(),
                status="DRAFT"
            )
        )

    # for bill in app.data["Bills"]["Details"].values():
    #     for sub_bill in bill["Content"]:
    #         bill_number = app.data["Project Info"]["Project"]["Quotation Number"].get() + sub_bill["Number"].get()
    #         if len(sub_bill["Number"].get()) == 0:
    #             continue
    #         invoices_list.append(
    #             Invoice(
    #                 type="ACCPAY",
    #                 contact=contact,
    #                 date=datetime.today(),
    #                 due_date=datetime.today(),
    #                 line_items=[
    #                     LineItem(
    #                         description=sub_bill["Service"].get(),
    #                         quantity=1,
    #                         unit_amount=int(sub_bill["Fee"].get()),
    #                         account_code="200"
    #                     )
    #                 ],
    #                 invoice_number=bill_number,
    #                 reference=data["Project Info"]["Project"]["Quotation Number"].get()+"-"+data["Project Info"]["Project"]["Project Name"].get(),
    #                 status="AUTHORISED"
    #             )
    #         )
    invoices = Invoices(invoices=invoices_list)
    try:
        accounting_api.update_or_create_invoices(xero_tenant_id, invoices)
    except Exception as e:
        print(e)
        print("No Data Processed")

    all_invoices = _process_invoices(accounting_api.get_invoices(xero_tenant_id).to_dict()["invoices"])
    #
    update_app_invoices(app, all_invoices)
    #
    if len(app.data["Asana_id"].get()) != 0:
        update_asana_invoices(app)
    return True
    #
    # messagebox.showinfo("Update", "the invoices and bill is updated to xero")




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
def create_contact(app, name):
    xero_tenant_id = get_xero_tenant_id()
    accounting_api = AccountingApi(api_client)

    new_contact = Contact(
        name=name
    )
    contacts = Contacts(contacts=[new_contact])


    api_response = accounting_api.create_contacts(xero_tenant_id, contacts)
    return api_response.to_dict()["contacts"][0]["contact_id"]

