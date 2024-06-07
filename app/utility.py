import textwrap
from tkinter import messagebox
import tkinter as tk

from app_dialog import FileSelectDialog, RadioButtonSelectDialog

from win32com import client as win32client
import shutil
import os
import json
from datetime import date, datetime
import psutil
import subprocess
from config import CONFIGURATION as conf
import _thread

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

import pythoncom
# pythoncom.CoInitialize()


def convert_float(string):
    if string == "":
        return 0
    else:
        return float(string)
def isfloat(string):
    try:
        float(string)
        return True
    except ValueError:
        return False
    except TypeError:
        return False

def increment_excel_column(cur_col):
    #only accept at most two digits of column
    assert len(cur_col) <= 2
    if len(cur_col) == 1:
        return chr(ord(cur_col[-1])+1) if cur_col!="Z" else "AA"
    else:
        if cur_col[-1]=="Z":
            return chr(ord(cur_col[0])+1)+"A"
        else:
            return cur_col[0] + chr(ord(cur_col[-1])+1)

def reset(app):
    data = app.data
    # if len(app.current_quotation.get())

    if len(app.current_quotation.get()) != 0:
        data["Project Info"]["Project"]["Quotation Number"].set(app.current_quotation.get())
        app.data["Login_user"].set("")
        save(app)
        app.current_quotation.set("")
        app.data["Login_user"].set(app.user)

    database_dir = os.path.join(app.conf["database_dir"], "data_template.json")
    template_json = json.load(open(database_dir))
    template_json["Fee Proposal"]["Reference"]["Date"] = datetime.today().strftime("%d-%b-%Y")

    test_data = data.copy()
    try:
        convert_to_data(template_json, test_data)
    except Exception as e:
        print(e)
        app.messagebox.show_error("The template Json file is corrupt, please contact Admin to fix it")
        return


    convert_to_data(template_json, app.data)

    app.fee_proposal_page._reset_scope()

    for service in app.data["Project Info"]["Project"]["Service Type"].values():
        if service["Include"].get():
            service["Include"].set(False)
            service["Include"].set(True)

    app.log_text.set("")

    # app.email_text.delete(1.0, tk.END)

# def _scopt_reset(app):
#     data = app.data
#     for service in conf["service_list"]:
#         if data["Invoices"]["Details"][service]["Include"].get():

def save_to_mp(app, data_json):
    database_dir = app.conf["database_dir"]
    quotation = data_json["Project Info"]["Project"]["Quotation Number"]
    mp_dir = os.path.join(database_dir, "mp.json")
    mp_json = json.load(open(mp_dir))
    mp_convert_map = app.search_bar_page.mp_convert_map

    mp_json[quotation] = {k:v(data_json) for k, v in mp_convert_map.items()}

    mp_json_list = list(mp_json.keys())
    mp_json_list.sort(reverse=True)
    new_mp_json_list = {}
    for quotation in mp_json_list:
        new_mp_json_list[quotation] = mp_json[quotation]
    mp_json = new_mp_json_list

    with open(mp_dir, "w") as f:
        json_object = json.dumps(mp_json, indent=4)
        f.write(json_object)


def save(app):
    data = app.data
    data_json = convert_to_json(data)
    database_dir = os.path.join(app.conf["database_dir"], data_json["Project Info"]["Project"]["Quotation Number"])
    # accounting_dir = os.path.join(app.conf["accounting_dir"], data_json["Project Info"]["Project"]["Quotation Number"])
    # if len(data["Project Info"]["Project"]["Quotation Number"].get()) != 6 or len(data["Project Info"]["Project"]["Quotation Number"].get()) != 8:
    #     return
    if not os.path.exists(database_dir):
        os.mkdir(database_dir)
    # if not os.path.exists(accounting_dir):
    #     os.makedirs(accounting_dir)
    with open(os.path.join(database_dir, "data.json"), "w") as f:
        json_object = json.dumps(data_json, indent=4)
        f.write(json_object)
    save_to_mp(app, data_json)
    # save_invoice_state(app)
    # current_folder_name = [folder for folder in os.listdir(app.conf["working_dir"]) if folder.startswith(data["Project Info"]["Project"]["Quotation Number"])][0]
    # print(current_folder_name)

def load_data(app, quotation_number):
    data = app.data
    database_dir = conf["database_dir"]
    if not os.path.exists(os.path.join(database_dir, quotation_number)):
        messagebox.showerror("Error", f"Can not find the folder {os.path.join(database_dir, quotation_number)}")
        return
    project_json = json.load(open(os.path.join(database_dir, quotation_number, "data.json")))
    if len(project_json["Login_user"]) !=0 and project_json["Login_user"] != app.user:
        messagebox.showerror("Error", f"{project_json['Login_user']} is Using this project right now")
        return

    if len(app.current_quotation.get()) != 0:
        old_quotation = data["Project Info"]["Project"]["Quotation Number"].get()
        data["Project Info"]["Project"]["Quotation Number"].set(app.current_quotation.get())
        app.data["Login_user"].set("")
        save(app)
        data["Project Info"]["Project"]["Quotation Number"].set(old_quotation)
        app.data["Login_user"].set(app.user)
    # reset(app)
    load(app, quotation_number)
    # app.email_text.delete(1.0, tk.END)
    # app.email_text.insert(1.0, app.data["Email_Content"].get())

def load(app, quotation_number):
    data = app.data
    # quotation_number = data["Project Info"]["Project"]["Quotation Number"].get().upper()
    database_dir = os.path.join(app.conf["database_dir"], quotation_number)
    data_json = json.load(open(os.path.join(database_dir, "data.json")))

    test_data = data.copy()
    try:
        convert_to_data(data_json, test_data)
    except Exception as e:
        print(e)
        app.messagebox.show_error("The Json file is corrupt, please contact Admin to fix it")
        return

    convert_to_data(data_json, data)

    app.timer = app.conf["timer"]
    app.data["Login_user"].set(app.user)
    save(app)

    for service in app.data["Project Info"]["Project"]["Service Type"].values():
        if service["Include"].get():
            service["Include"].set(False)
            service["Include"].set(True)

    app.current_quotation.set(quotation_number)
    app.fee_proposal_page.auto_lock()
    # app.fee_proposal_page.reset_stage()
    unlock_invoice(app)
    load_invoice_state(app)
    config_state(app)
    config_log(app)

def unlock_invoice(app):
    for service, value in app.financial_panel_page.invoice_dic.items():
        # for inv in value["Invoice"]:
        #     inv.config(state=tk.NORMAL)
        for j, item in enumerate(value["Content"]):
            for inv in item["Invoice"]:
                inv.config(state=tk.NORMAL)

def save_invoice_state(app):
    data = app.data
    inv_dir = os.path.join(app.conf["database_dir"], "invoices.json")
    inv_json = json.load(open(inv_dir))
    for inv in data["Invoices Number"]:
        if len(inv["Number"].get())!=0:
            inv_json[inv["Number"].get()] = inv["State"].get()
    invoice_number_list = list(inv_json.keys())
    invoice_number_list.sort()
    inv_json = {i: inv_json[i] for i in invoice_number_list}

    with open(inv_dir, "w") as f:
        json.dump(inv_json, f, indent=4)

    bill_dir = os.path.join(app.conf["database_dir"], "bills.json")
    bill_json = json.load(open(bill_dir))
    for bill in data["Bills"]["Details"].values():
        for item in bill["Content"]:
            if len(item["Number"].get()) != 0:
                bill_json[data["Project Info"]["Project"]["Project Number"].get()+item["Number"].get()] = item["State"].get()
    with open(bill_dir, "w") as f:
        json.dump(bill_json, f, indent=4)

def load_invoice_state(app):
    data = app.data
    inv_dir = os.path.join(app.conf["database_dir"], "invoices.json")
    inv_json = json.load(open(inv_dir))
    for inv in data["Invoices Number"]:
        if len(inv["Number"].get()) != 0:
            inv["State"].set(inv_json[inv["Number"].get()])

    bill_dir = os.path.join(app.conf["database_dir"], "bills.json")
    bill_json = json.load(open(bill_dir))
    for bill in data["Bills"]["Details"].values():
        for item in bill["Content"]:
            if len(item["Xero_id"].get()) !=0:
                item["State"].set(bill_json[item["Xero_id"].get()])

def convert_to_json(obj):
    if isinstance(obj, list):
        return [convert_to_json(x) for x in obj]
    elif isinstance(obj, dict):
        return {k: convert_to_json(v) for k, v in obj.items()}
    else:
        return obj.get()

def convert_to_data(json, data):
    if isinstance(json, list):
        if len(json) == len(data):
            [convert_to_data(json[i], data[i]) for i in range(len(json))]
        elif len(json) > len(data):
            for j in range(len(data), len(json)):
                data.append(
                    {
                        "Include": tk.BooleanVar(value=True),
                        "Item": tk.StringVar()
                    }
                )
            for i in range(len(json)):
                convert_to_data(json[i], data[i])
        elif len(json) < len(data):
            for j in range(len(json), len(data)):
                data.pop(-1)
            for i in range(len(data)):
                convert_to_data(json[i], data[i])

    elif isinstance(json, dict):
        for key in json.keys():
            convert_to_data(json[key], data[key])
    else:
        data.get()
        if data.get() != json:
            try:
                data.set(json)
            except Exception as e:
                print(json)
                print(e)


def generate_all_projects_and_invoices(database_dir):
    projects = {}
    invoices = {}
    bills = {}
    for dir in os.listdir(database_dir):
        if os.path.isdir(os.path.join(database_dir, dir)):
            data_json = json.load(open(os.path.join(database_dir, dir, "data.json")))
            if data_json["Asana_id"] != "":
                projects[data_json["Asana_id"]]={
                    "Asana_url": data_json["Asana_url"],
                    "Quotation Number": data_json["Project Info"]["Project"]["Quotation Number"],
                    "Project Number": data_json["Project Info"]["Project"]["Project Number"],
                    "Invoices": []
                }
                for inv in data_json["Invoices Number"]:
                    if len(inv["Asana_id"])!=0 and inv["State"]=="Backlog":
                        invoices[inv["Asana_id"]] = {
                            "Number": inv["Number"],
                            "Fee": inv["Fee"],
                            # "State": inv["State"]
                        }
                        # projects[inv["Asana_id"]]["Invoices"].append(inv["Asana_id"])
                for service in data_json["Bills"]["Details"].values():
                    for content in service["Content"]:
                        if len(content["Xero_id"])!=0:
                            assert len(content["Asana_id"])!=0
                            bills[content["Xero_id"]] = content["Asana_id"]
    return projects, invoices, bills

def move_new_folder_to_external(app, working_dir, new_folder_dir, folder_name, quotation_number):
    os.rename(os.path.join(working_dir, new_folder_dir), os.path.join(working_dir, "External"))
    create_new_folder(folder_name, app.conf, quotation_number, False)
    os.rename(os.path.join(working_dir, "External"), os.path.join(working_dir, folder_name,"External"))
def finish_setup(app):
    data = app.data

    quotation_number = data["Project Info"]["Project"]["Quotation Number"].get()
    current_folder_address = data["Current_folder_address"].get()
    folder_name = quotation_number + "-" + data["Project Info"]["Project"]["Project Name"].get()
    working_dir = conf["working_dir"]

    if data["State"]["Set Up"].get():
        app.messagebox.show_error("You already setup the project")
        return

    if current_folder_address == "":
        new_folder_list = []
        for directory in os.listdir(working_dir):
            if directory.startswith("New folder"):
                new_folder_list.append(directory)

        if len(new_folder_list) > 1:
            file_selection = tk.StringVar()
            FileSelectDialog(app, new_folder_list, file_selection, "Multiple New Folder Found, please select the folder you want to rename")
            if len(file_selection.get())!=0:
                new_folder = file_selection.get()
                move_new_folder_to_external(app, working_dir, new_folder, folder_name, quotation_number)
            else:
                return
        elif len(new_folder_list)==1:
            create_folder = app.messagebox.ask_yes_no(f"Bridge found new folder {new_folder_list[0]}, do you want to create the folder")
            if create_folder:
                new_folder = new_folder_list[0]
                move_new_folder_to_external(app, working_dir, new_folder, folder_name, quotation_number)
            else:
                return
        else:
            create_folder = app.messagebox.ask_yes_no(f"Can not find the folder with quotation number {quotation_number} or New Folder, do you want to create the folder")
            if create_folder:
                create_new_folder(folder_name, app.conf, quotation_number)
            else:
                return
    else:
        if current_folder_address != folder_name:
            try:
                os.rename(os.path.join(working_dir, current_folder_address), os.path.join(working_dir, folder_name))
            except Exception as e:
                print(e)
                app.messagebox.show_error("Please Close the file relate to folder before rename")
                return
        #
        # if os.
    # working_dir = app.conf["working_dir"]
    # quotation_number = data["Project Info"]["Project"]["Quotation Number"].get()
    # folder_name = quotation_number + "-" + data["Project Info"]["Project"]["Project Name"].get()
    # folder_name = data["Project Info"]["Project"]["Quotation Number"].get() + "-" + data["Project Info"]["Project"]["Project Name"].get().strip()
    # working_dir = os.path.join(app.conf["working_dir"], folder_name)

    # found = False
    # current_folder_name = None
    # for dir in os.listdir(working_dir):
    #     if os.path.isdir(os.path.join(working_dir, dir)) and dir.split("-")[0]==quotation_number:
    #         found = True
    #         current_folder_name = dir
    # if not found:
    #     create_folder = messagebox.askyesno("Folder not found",
    #                                         f"Can not find the folder with quotation number {quotation_number}, do you want to create the folder")
    #     if create_folder:
    #         create_new_folder(folder_name, app.conf)
    #     else:
    #         return
    # else:
    #     if current_folder_name != folder_name:
    #         try:
    #             os.rename(os.path.join(working_dir, current_folder_name), os.path.join(working_dir, folder_name))
    #         except Exception as e:
    #             print(e)
    #             app.messagebox.show_error("Error", "Please Close the file relate to folder before rename")
    #             return
    data["State"]["Set Up"].set(True)
    data["Current_folder_address"].set(folder_name)
    app.current_quotation.set(quotation_number)
    save(app)
    app.log.log_finish_set_up(app)
    config_state(app)
    config_log(app)
    app.messagebox.show_info(f"Project {data['Project Info']['Project']['Quotation Number'].get()} set up successful")

def config_state(app):
    return
    database_dir = app.conf["database_dir"]
    res = {
        "Set Up": [],
        "Generate Proposal": [],
        "Email to Client": [],
        "Fee Accepted": []
    }
    current_project = [folder for folder in os.listdir(database_dir) if not folder.endswith(".json")]
    for quotation in current_project:
        data = json.load(open(os.path.join(database_dir, quotation, "data.json")))
        _classify_state(data, res)
    app.state_dict["Set Up"].config(value=res["Set Up"])
    app.state_dict["Generate Proposal"].config(value=res["Generate Proposal"])
    app.state_dict["Email to Client"].config(value=res["Email to Client"])
    res["Fee Accepted"].sort(key=lambda x: int(x.split("-")[1]))
    app.state_dict["Fee Accepted"].config(value=res["Fee Accepted"])

def _classify_state(data, res):
    if data["State"]["Fee Accepted"] or data["State"]["Quote Unsuccessful"]:
        return
    elif data["State"]["Email to Client"]:
        _classify_fee(res, data)
    elif data["State"]["Generate Proposal"]:
        res["Email to Client"].append(
            data["Project Info"]["Project"]["Quotation Number"]+"-"+data["Project Info"]["Project"]["Project Name"])
    elif data["State"]["Set Up"]:
        res["Generate Proposal"].append(
            data["Project Info"]["Project"]["Quotation Number"]+"-"+data["Project Info"]["Project"]["Project Name"])
    else:
        res["Set Up"].append(
            data["Project Info"]["Project"]["Quotation Number"]+"-"+data["Project Info"]["Project"]["Project Name"])

def _classify_fee(res, data):
    # if not "-" in data["Email"]["Fee Proposal"]:
    #     data["Email"]["Fee Proposal"] = datetime.today().strftime("%Y-%m-%d")
    append_list = [str((datetime.today() - datetime.strptime(data["Email"]["Fee Proposal"], "%Y-%m-%d")).days)]
    if len(data["Email"]["First Chase"]) != 0:
        append_list.append(
            str((datetime.today() - datetime.strptime(data["Email"]["First Chase"], "%Y-%m-%d")).days)
        )
    if len(data["Email"]["Second Chase"]) != 0:
        append_list.append(
            str((datetime.today() - datetime.strptime(data["Email"]["Second Chase"], "%Y-%m-%d")).days)
        )
    if len(data["Email"]["Third Chase"]) != 0:
        append_list.append(
            str((datetime.today() - datetime.strptime(data["Email"]["Third Chase"], "%Y-%m-%d")).days)
        )
    append_list.append(data["Project Info"]["Project"]["Quotation Number"])
    append_list.reverse()
    res["Fee Accepted"].append("-".join(append_list))
    # elif len(data["Email"]["Second Chase"]) == 0:
    #     data_json["Email"]["Second Chase"] = datetime.strptime(Date, "%a, %d %b %Y %H:%M:%S %z").strftime("%Y-%m-%d")
    # elif len(data["Email"]["Third Chase"]) == 0:
    #     data_json["Email"]["Third Chase"] = datetime.strptime(Date, "%a, %d %b %Y %H:%M:%S %z").strftime("%Y-%m-%d")

def delete_project(app):
    data = app.data
    if len(data["Project Info"]["Project"]["Project Number"].get()) != 0:
        folder_name = data["Project Info"]["Project"]["Project Number"].get() + "-" + data["Project Info"]["Project"][
            "Project Name"].get()
    else:
        folder_name = data["Project Info"]["Project"]["Quotation Number"].get() + "-" + data["Project Info"]["Project"][
            "Project Name"].get()
    working_dir = os.path.join(app.conf["working_dir"], folder_name)
    database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    recycle_folder = data["Project Info"]["Project"]["Quotation Number"].get() + "-" + datetime.now().strftime(
        "%y%m%d%H%M%S")
    recycle_bin_dir = os.path.join(app.conf["recycle_bin_dir"], recycle_folder)

    os.mkdir(recycle_bin_dir)

    app.log.log_delete(app)
    shutil.move(database_dir, recycle_bin_dir)
    if os.path.exists(working_dir):
        shutil.move(working_dir, recycle_bin_dir)
    app.current_quotation.set("")
    data["Project Info"]["Project"]["Quotation Number"].set("")
    reset(app)

def config_log(app):
    database_dir = app.conf["database_dir"]
    log_file = os.path.join(database_dir, app.data["Project Info"]["Project"]["Quotation Number"].get(), "data.log")
    try:
        log = open(log_file).readlines()
        log.reverse()
    except FileNotFoundError:
        log = ""
    app.log_text.set("".join(log))

def get_quotation_number():
    working_dir = conf["database_dir"]
    current_quotation_list = [dir for dir in os.listdir(working_dir) if
                              dir.startswith(date.today().strftime("%y%m000")[1:])]
    if len(current_quotation_list) == 0:
        current_quotation = date.today().strftime("%y%m000")[1:] + "AA"
    else:
        current_quotation = current_quotation_list[-1]
        quotation_letter = current_quotation[6:8][0] + chr(ord(current_quotation[6:8][1]) + 1) if current_quotation[6:8][1] != "Z" else chr(ord(current_quotation[6:8][0]) + 1) + "A"
        current_quotation = current_quotation[:6] + quotation_letter
    return current_quotation

# def remove_none(obj):
#     if isinstance(obj, list):
#         return [remove_none(x) for x in obj if x is not None]
#     elif isinstance(obj, dict):
#         return {k: remove_none(v) for k, v in obj.items() if v is not None}
#     else:
#         return obj

def rename_project(app):
    #determine whether to change the quotation  number or the project name
    working_dir = app.conf["working_dir"]
    dir_list = os.listdir(working_dir)
    data = app.data

    if len(data["Project Info"]["Project"]["Project Number"].get()) !=0:
        current_quotation = data["Project Info"]["Project"]["Project Number"].get()
    else:
        current_quotation = data["Project Info"]["Project"]["Quotation Number"].get()

    for folder in dir_list:
        if folder[:5].isdigit():
            quotation_number = folder.split("-", 1)[0]
            if quotation_number == current_quotation:
                old_dir = os.path.join(working_dir, folder)
                new_dir = os.path.join(working_dir, current_quotation+"-"+data["Project Info"]["Project"]["Project Name"].get())
                os.rename(old_dir, new_dir)
                return old_dir, new_dir



def change_quotation_number(app, new_quotation_number):
    data = app.data
    working_dir = app.conf["working_dir"]
    current_folder_address = os.path.join(working_dir, data["Current_folder_address"].get())
    # database_dir = app.conf["database_dir"]

    # old_folder_name = app.data["Project Info"]["Project"]["Quotation Number"].get() + "-" + app.data["Project Info"]["Project"]["Project Name"].get()
    # old_folder = os.path.join(working_dir, old_folder_name)

    if not os.path.exists(current_folder_address):
        messagebox.showerror("Error", f"Can not found the folder {current_folder_address}")
        return "", "", False

    new_folder = os.path.join(working_dir, new_quotation_number + "-" + app.data["Project Info"]["Project"]["Project Name"].get())

    try:
        # old_database = os.path.join(database_dir, app.data["Project Info"]["Project"]["Quotation Number"].get())
        # new_database = os.path.join(database_dir, new_quotation_number)
        # os.rename(old_database, new_database)
        os.rename(current_folder_address, new_folder)
        return current_folder_address, new_folder, True
    except PermissionError:
        return current_folder_address, new_folder, False

# def rename_new_folder(app):
#     data = app.data
#     folder_name = data["Project Info"]["Project"]["Quotation Number"].get() + "-" + \
#                   data["Project Info"]["Project"]["Project Name"].get()
#     folder_path = os.path.join(app.conf["working_dir"], folder_name)
#     dir_list = os.listdir(app.conf["working_dir"])
#     rename_list = [dir for dir in dir_list if dir.startswith("New folder")]
#
#     if len(data["Project Info"]["Project"]["Project Name"].get()) == 0:
#         messagebox.showerror(title="Error", message="Please Input a project name")
#         return
#     elif len(data["Project Info"]["Project"]["Quotation Number"].get()) == 0:
#         messagebox.showerror(title="Error", message="Please Input a quotation number")
#         return
#     elif len(rename_list) == 0:
#         messagebox.showerror(title="Error", message="Please create a new folder first")
#         return
#
#     try:
#         if len(rename_list) == 1:
#             mode = 0o666
#             os.mkdir(folder_path, mode)
#             shutil.move(os.path.join(app.conf["working_dir"], rename_list[0]), os.path.join(folder_path, "External"))
#             os.mkdir(os.path.join(folder_path, "Photos"), mode)
#             os.mkdir(os.path.join(folder_path, "Plot"), mode)
#             os.mkdir(os.path.join(folder_path, "SS"), mode)
#             shutil.copyfile(os.path.join(app.conf["resource_dir"], "xlsx", app.conf["calculation_sheet"]),
#                             os.path.join(folder_path, app.conf["calculation_sheet"]))
#             messagebox.showinfo(title="Folder renamed", message=f"Rename Folder {rename_list[0]} to {folder_name}")
#         else:
#             FileSelectDialog(app, rename_list, "Multiple new folders found, please select one")
#     except FileExistsError:
#         messagebox.showerror("Error", f"Cannot create a file when that file already exists:{folder_name}")
#         return
#     save(app)
#     config_state(app)
#     app.log.log_rename_folder(app)
#     config_log(app)

def create_new_folder(folder_name, conf, quotation, creating_external=True):
    database_dir = os.path.join(conf["database_dir"], quotation)
    accounting_dir = os.path.join(conf["accounting_dir"], quotation)
    folder_path = os.path.join(conf["working_dir"], folder_name)
    calculation_sheet = conf["calculation_sheet"]

    os.makedirs(database_dir)
    os.makedirs(accounting_dir)
    os.mkdir(folder_path)
    if creating_external:
        os.mkdir(os.path.join(folder_path, "External"))
    os.mkdir(os.path.join(folder_path, "Photos"))
    os.mkdir(os.path.join(folder_path, "Plot"))
    os.mkdir(os.path.join(folder_path, "SS"))
    shutil.copyfile(os.path.join(conf["resource_dir"], "xlsx", calculation_sheet),
                    os.path.join(folder_path, calculation_sheet))

    shortcut_dir = os.path.join(folder_path, "Database Shortcut.lnk")
    shortcut_working_dir = conf["accounting_dir"]
    shell = win32client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(shortcut_dir)
    shortcut.Targetpath = accounting_dir
    shortcut.WorkingDirectory = shortcut_working_dir
    # shortcut.IconLocation = os.path.join(conf["resource_dir"], "jpg", "logo.jpg")
    shortcut.save()


def _check_fee(app):
    data = app.data
    for service in conf["service_list"]:
        fee = data["Invoices"]["Details"][service]
        if len(fee["Fee"].get()) == 0 and fee["Include"].get():
            return False
    return True
    # for service_fee in data["Invoices"]["Details"].values():
    #     if len(service_fee["Fee"].get()) == 0 and service_fee["Service"].get() != "Variation" and service_fee["Include"].get():
    #         return False
    # return True

def check_accounting_folder(app):
    accounting_dir = os.path.join(conf["accounting_dir"], app.data["Project Info"]["Project"]["Quotation Number"].get())
    if not os.path.exists(accounting_dir):
        os.makedirs(accounting_dir)


def open_design_certificate(app):
    data = app.data
    quotation_number = data["Project Info"]["Project"]["Quotation Number"].get().upper()
    adobe_address = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"

    # if len(data["Project Info"]["Project"]["Project Number"].get()) == 0:
    #     folder_name = data["Project Info"]["Project"]["Quotation Number"].get() + "-" + \
    #                   data["Project Info"]["Project"]["Project Name"].get()
    # else:
    #     folder_name = data["Project Info"]["Project"]["Project Number"].get() + "-" + \
    #                   data["Project Info"]["Project"]["Project Name"].get()
    folder_name = data["Current_folder_address"].get()
    folder_path = conf["working_dir"]+"\\"+folder_name

    if len(quotation_number) == 0:
        app.messagebox.show_error("Please enter a Quotation Number before you load")
        return
    elif not os.path.exists(folder_path):
        app.messagebox.show_error(f"Python cannot find the folder {folder_path}")
        return

    # proposal_type = data["Project Info"]["Project"]["Proposal Type"].get()
    # if proposal_type == "Minor":
    calculation_sheet = "Preliminary Calculation"
    # else:
    #     calculation_sheet = "Mech System Calculation"
    excel_path = None
    for file in os.listdir(folder_path):
        if file.endswith(".xlsx") and file.startswith(calculation_sheet):
            excel_path =file
            break
    if excel_path is None:
        app.messagebox.show_error(f"Can not find the {calculation_sheet} file under {folder_path}")
        return
    # excel_full_path = os.path.join(folder_path, excel_path)

    excel_full_path = os.path.join(folder_path, excel_path)
    # wb = load_workbook(excel_full_path)
    # calculation_sheet = load_workbook(excel_full_path)
    # sheet_names = calculation_sheet.sheetnames
    # if not ("Mechanical Design Certificate" in sheet_names and "Mech Design Compliance Cert" in sheet_names):
    #     app.messagebox.show_error("Mechanical Design Certificate and Compliance Cert are not in the excel sheet")
    #     return
    # design_certificate_worksheet = calculation_sheet["Mechanical Design Certificate"]
    # design_compliance_cert_work_sheet = calculation_sheet["Mech Design Compliance Cert"]
    # design_certificate_worksheet["A2"] = data["Project Info"]["Project"]["Project Name"].get()
    # design_compliance_cert_work_sheet["A2"] = data["Project Info"]["Project"]["Project Name"].get()
    # print(calculation_sheet)

    excel = win32client.Dispatch("Excel.Application")
    try:
        excel.Visible = True
    except Exception as e:
        print(e)
    # excel.ScreenUpdating = False
    # excel.DisplayAlerts = False
    # excel.EnableEvents = False
    work_book = excel.Workbooks.Open(excel_full_path)
    work_book.Worksheets[0].Activate()

    design_certificate_worksheet = work_book.Worksheets["Mechanical Design Certificate"]
    design_compliance_cert_work_sheet = work_book.Worksheets["Mech Design Compliance Cert"]
    design_certificate_worksheet.Cells(2, 1).Value = data["Project Info"]["Project"]["Project Name"].get()
    design_compliance_cert_work_sheet.Cells(2, 1).Value = data["Project Info"]["Project"]["Project Name"].get()

    # design_certificate_first_cell = 48 #if proposal_type=="Minor" else 55
    # design_compliance_first_cell = 26

    i = 1
    design_certificate_first_cell = None
    while i < 100:
        if design_certificate_worksheet.Cells(i, 2).Value == "Drawing Number":
            design_certificate_first_cell = i+1
            break
        i+=1
    if design_certificate_first_cell is None:
        app.messagebox.show_error("Can not found the drawing row of the design certificate")
        return
    i = 1
    design_compliance_first_cell = None
    while i < 100:
        if design_compliance_cert_work_sheet.Cells(i, 2).Value == "Drawing Number":
            design_compliance_first_cell = i + 1
            break
        i += 1
    if design_compliance_cert_work_sheet is None:
        app.messagebox.show_error("Can not found the drawing row of the design compliance certificate")
        return


    for i, drawing in enumerate(data["Project Info"]["Drawing"]):
        design_certificate_worksheet.Cells(design_certificate_first_cell+i, 2).Value = drawing["Drawing Number"].get()
        design_certificate_worksheet.Cells(design_certificate_first_cell+i, 4).Value = drawing["Drawing Name"].get()
        design_certificate_worksheet.Cells(design_certificate_first_cell+i, 8).Value = drawing["Revision"].get()

        design_compliance_cert_work_sheet.Cells(design_compliance_first_cell+i, 2).Value = drawing["Drawing Number"].get()
        design_compliance_cert_work_sheet.Cells(design_compliance_first_cell+i, 4).Value = drawing["Drawing Name"].get()
        design_compliance_cert_work_sheet.Cells(design_compliance_first_cell+i, 8).Value = drawing["Revision"].get()



    print_pdf = tk.StringVar(value="None")
    RadioButtonSelectDialog(app,
                            ["Design Certificate", "Design Compliance Certificate"],
                            print_pdf,
                            "Do you want to generate",
                            data["Project Info"]["Project"]["Project Name"].get())
    # print_pdf = app.messagebox.ask_yes_no("Do you want to Save the design certificate and design compliance in the Plot Folder?")
    print_pdf = print_pdf.get()
    if print_pdf == "Design Certificate":
        pdf_path = os.path.join(folder_path, "Plot", "Mechanical Design Certificate.pdf")
        design_certificate_worksheet.ExportAsFixedFormat(0, pdf_path)
        def open_pdf():
            subprocess.call([adobe_address, pdf_path])

        _thread.start_new_thread(open_pdf, ())
    elif print_pdf == "Design Compliance Certificate":
        pdf_path = os.path.join(folder_path, "Plot", "Mechanical Design Compliance Certificate.pdf")
        design_compliance_cert_work_sheet.ExportAsFixedFormat(0, pdf_path)
        def open_pdf():
            subprocess.call([adobe_address, pdf_path])

        _thread.start_new_thread(open_pdf, ())


# def change_type_of_calculator(app):
#     data = app.data
#     proposal_type = data["Project Info"]["Project"]["Proposal Type"].get()
#     quotation_number = data["Project Info"]["Project"]["Quotation Number"].get().upper()
#
#
#     if len(data["Project Info"]["Project"]["Project Number"].get()) == 0:
#         folder_name = data["Project Info"]["Project"]["Quotation Number"].get() + "-" + \
#                       data["Project Info"]["Project"]["Project Name"].get()
#     else:
#         folder_name = data["Project Info"]["Project"]["Project Number"].get() + "-" + \
#                       data["Project Info"]["Project"]["Project Name"].get()
#     folder_path = os.path.join(app.conf["working_dir"], folder_name)
#
#     if len(quotation_number) == 0:
#         app.messagebox.show_error("Please enter a Quotation Number before you update asana")
#         return
#     elif not os.path.exists(folder_path):
#         app.messagebox.show_error(f"Python cannot find the folder {folder_path}")
#         return
#     excel = None
#     for dir in os.listdir(folder_path):
#         if (dir.startswith("Preliminary Calculation") or dir.startswith("Mech System Calculation")) and dir.endswith(".xlsx"):
#             excel = dir
#             break
#     if proposal_type == "Minor" and excel.startswith("Mech System Calculation"):
#         try:
#             os.remove(os.path.join(folder_path, excel))
#             shutil.copyfile(os.path.join(app.conf["resource_dir"], "xlsx", "Preliminary Calculation v2.6.xlsx"),
#                             os.path.join(folder_path, "Preliminary Calculation v2.6.xlsx"))
#         except Exception as e:
#             print(e)
#             return False
#     elif proposal_type == "Major" and excel.startswith("Preliminary Calculation"):
#         try:
#             os.remove(os.path.join(folder_path, excel))
#             shutil.copyfile(os.path.join(app.conf["resource_dir"], "xlsx", "Mech System Calculation v0.4.xlsx"),
#                             os.path.join(folder_path, "Mech System Calculation v0.4.xlsx"))
#         except Exception as e:
#             print(e)
#             return False
#     elif excel is None:
#         file = "Preliminary Calculation v2.6.xlsx" if proposal_type == "Minor" else "Mech System Calculation v0.4.xlsx"
#         shutil.copyfile(os.path.join(app.conf["resource_dir"], "xlsx", file),
#                         os.path.join(folder_path, file))
#     return True

def preview_fee_proposal(app, *args):
    check_accounting_folder(app)
    if app.data["Project Info"]["Project"]["Proposal Type"].get() == "Minor":
        minor_project_fee_proposal(app)
    elif app.data["Project Info"]["Project"]["Proposal Type"].get() == "Major":
        major_project_fee_proposal(app)
    else:
        print("Unsupported Project Type")

def get_first_name(name):
    return name.split(" ")[0]

def separate_line(line):
    len_per_line = conf["len_per_line"]
    return textwrap.wrap(line, len_per_line, break_long_words=False)

def minor_project_fee_proposal(app, *args):
    data = app.data
    pdf_name = _get_proposal_name(app) + ".pdf"
    excel_name = f'{data["Project Info"]["Project"]["Project Name"].get()} Back Up.xlsx'
    # database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    accounting_dir = os.path.join(app.conf["accounting_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    resource_dir = app.conf["resource_dir"]
    adobe_address = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
    past_projects_dir = os.path.join(app.conf["database_dir"], "past_projects.json")
    services = []
    for service in app.conf["proposal_list"]:
        if data["Invoices"]["Details"][service]["Include"].get():
            services.append(service)
    page = len(services)
    row_per_page = 46

    if not data["State"]["Set Up"].get():
        messagebox.showerror("Error", "Please finish Set Up first")
        return
    elif len(services) == 0:
        messagebox.showerror("Error", "Please at least select 1 service")
        return
    elif not _check_fee(app):
        messagebox.showerror("Error", "Please go to fee proposal page to complete the fee first")
        return
    elif data["Invoices"]["Fee"].get() == "Error":
        messagebox.showerror(title="Error", message="There are error in the fee proposal section, please fix the fee section before generate the fee proposal")
        return
    elif not page in [1, 2, 3, 4, 5]:
        messagebox.showerror("Error", "Excess the maximum value of service, please contact administrator")
    pdf_list = [file for file in os.listdir(os.path.join(accounting_dir)) if
                "Fee Proposal" in str(file) and str(file).endswith(".pdf") and not "Mechanical Installation Fee Proposal" in str(file)]

    if len(pdf_list) != 0:
        current_pdf = sorted(pdf_list, key=lambda pdf: str(pdf).split(" ")[-1].split(".")[0])[-1]
        current_revision = current_pdf.split(" ")[-1].split(".")[0]
        if data["Fee Proposal"]["Reference"]["Revision"].get() == current_revision or data["Fee Proposal"]["Reference"]["Revision"].get() == str(int(current_revision) + 1):
            old_pdf_path = os.path.join(accounting_dir, current_pdf)
            # avDoc.open(old_pdf_path, old_pdf_path)
            def open_pdf():
                subprocess.call([adobe_address, old_pdf_path])
            _thread.start_new_thread(open_pdf, ())

            overwrite = messagebox.askyesno(f"Warming", f"Revision {current_revision} found, do you want to generate Revision {data['Fee Proposal']['Reference']['Revision'].get()}")
            if not overwrite:
                return
            else:
                # avDoc.close(0)
                for proc in psutil.process_iter():
                    if proc.name() == "Acrobat.exe":
                        proc.kill()
        else:
            messagebox.showerror("Error",
                                 f'Current revision is {current_revision}, you can not use revision {data["Fee Proposal"]["Reference"]["Revision"].get()}')
            return
    else:
        if not data["Fee Proposal"]["Reference"]["Revision"].get() == "1":
            messagebox.showerror("Error", "There is no other existing fee proposal found, can only have revision 1")
            return

    total_fee = 0
    total_ist = 0
    for service in conf["proposal_list"]:
        value = data["Invoices"]["Details"][service]
        if value["Include"].get():
            total_fee += float(value["Fee"].get())
            total_ist += float(value["in.GST"].get())

    shutil.copy(os.path.join(resource_dir, "xlsx", f"fee_proposal_template_{page}.xlsx"),
                os.path.join(accounting_dir, excel_name))
    excel = win32client.Dispatch("Excel.Application")
    try:
        excel.Visible = False
        # excel.ScreenUpdating = False
        # excel.DisplayAlerts = False
        # excel.EnableEvents = False
    except Exception as e:
        print(e)
        pass
    work_book = excel.Workbooks.Open(os.path.join(accounting_dir, excel_name))
        # messagebox.showerror("Error", e)
    try:
        work_sheets = work_book.Worksheets[0]
        address_to = data["Address_to"].get()
        work_sheets.Cells(1, 2).Value = data["Project Info"][address_to]["Full Name"].get()
        work_sheets.Cells(2, 2).Value = data["Project Info"][address_to]["Company"].get()
        work_sheets.Cells(3, 2).Value = data["Project Info"][address_to]["Address"].get()
        first_name = get_first_name(data["Project Info"][address_to]["Full Name"].get())
        work_sheets.Cells(5, 1).Value = f'Dear {first_name},'
        work_sheets.Cells(1, 8).Value = data["Project Info"]["Project"]["Quotation Number"].get()
        work_sheets.Cells(2, 8).Value = data["Fee Proposal"]["Reference"]["Date"].get()
        work_sheets.Cells(3, 8).Value = data["Fee Proposal"]["Reference"]["Revision"].get()
        work_sheets.Cells(6, 1).Value = "Re: " + data["Project Info"]["Project"]["Project Name"].get()
        work_sheets.Cells(8, 1).Value = f"Thank you for giving us the opportunity to submit this fee proposal for our {', '.join([key for key, value in data['Project Info']['Project']['Service Type'].items() if value['Include'].get()])} for the above project."
        work_sheets.Cells(16, 7).Value = data["Fee Proposal"]["Time"]["Fee Proposal"].get()
        work_sheets.Cells(21, 7).Value = data["Fee Proposal"]["Time"]["Pre-design"].get()
        work_sheets.Cells(26, 7).Value = data["Fee Proposal"]["Time"]["Documentation"].get()
        cur_index = 1
        i = 0
        for key in app.conf["proposal_list"]:
            if not data["Invoices"]["Details"][key]["Include"].get():
                continue

            service = data["Fee Proposal"]["Scope"]["Minor"][key]

            # cur_row = cur_row if i % 2 == 0 else 84 + (i - 1) // 2 * row_per_page
            if i == 0:
                cur_row = 52
            else:
                cur_row = 84 + (i-1) * row_per_page
            extra_list = conf["extra_list"]
            for extra in extra_list:
                if len([scope for scope in service[extra] if scope["Include"].get()]) == 0:
                    continue
                work_sheets.Cells(cur_row, 1).Value = "2." + str(cur_index)
                work_sheets.Cells(cur_row, 2).Value = key + "-" + extra
                work_sheets.Cells(cur_row, 1).Font.Bold = True
                work_sheets.Cells(cur_row, 2).Font.Bold = True
                cur_row += 1
                for scope in service[extra]:
                    if scope["Include"].get():
                        work_sheets.Cells(cur_row, 1).Value = "•"
                        for line in separate_line(scope["Item"].get()):
                            work_sheets.Cells(cur_row, 2).Value = line
                            cur_row += 1
                cur_row += 1
                cur_index += 1
            i+=1

        cur_row = 102 + (page - 1) * row_per_page
        project_type = data["Project Info"]["Project"]["Project Type"].get()
        past_projects = json.load(open(past_projects_dir, encoding="utf-8"))[project_type]
        for i, project in enumerate(past_projects):
            work_sheets.Cells(cur_row + i, 1).Value = "•"
            work_sheets.Cells(cur_row + i, 2).Value = project

        cur_row += 34
        for i, service in enumerate([ value for value in data["Invoices"]["Details"].values() if value["Service"].get() != "Variation" and value["Include"].get()]):
            work_sheets.Cells(cur_row + i, 2).Value = service["Service"].get() + " design and documentation"
            work_sheets.Cells(cur_row + i, 6).Value = service["Fee"].get()
            work_sheets.Cells(cur_row + i, 7).Value = service["in.GST"].get()
        if page !=1:
            work_sheets.Cells(cur_row + page, 6).Value = total_fee
            work_sheets.Cells(cur_row + page, 7).Value = total_ist
        work_sheets.Cells(cur_row+13+page, 2).Value = "Re: " + data["Project Info"]["Project"]["Project Name"].get()
        work_sheets.ExportAsFixedFormat(0, os.path.join(accounting_dir, pdf_name))
        work_book.Close(True)
    except PermissionError:
        print(e)
        messagebox.showerror("Error", "Please close the preview or file before you use it")
    except FileNotFoundError:
        print(e)
        messagebox.showerror("Error", "Please Contact Administer, the app cant find the file")
    except Exception as e:
        print(e)
        messagebox.showerror("Error", "Please Close the pdf before you generate a new one")
    else:
        app.data["State"]["Generate Proposal"].set(True)
        # avDoc.open(os.path.join(database_dir, pdf_name), os.path.join(database_dir, pdf_name))
        def open_pdf():
            subprocess.call([adobe_address, os.path.join(accounting_dir, pdf_name)])
        _thread.start_new_thread(open_pdf, ())
        # webbrowser.open(os.path.join(database_dir, pdf_name))
        save(app)
        config_state(app)
        app.log.log_fee_proposal(app)
        config_log(app)
    try:
        excel.Visible = True
        # excel.ScreenUpdating = True
        # excel.DisplayAlerts = True
        # excel.EnableEvents = True
        work_book.Close(True)
        # adobe.close(0)
    except Exception as e:
        print(e)

def major_project_fee_proposal(app, *args):
    data = app.data
    pdf_name = _get_proposal_name(app) + ".pdf"
    excel_name = f'{data["Project Info"]["Project"]["Project Name"].get()} Back Up.xlsx'
    # database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    accounting_dir = os.path.join(app.conf["accounting_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    resource_dir = app.conf["resource_dir"]
    adobe_address = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"

    past_projects_dir = os.path.join(app.conf["database_dir"], "past_projects.json")
    services = []
    for service in app.conf["proposal_list"]:
        if data["Invoices"]["Details"][service]["Include"].get():
            services.append(service)
    page = len(services)

    row_per_page = 46

    if not data["State"]["Set Up"].get():
        messagebox.showerror("Error", "Please finish Set Up first")
        return
    elif len(services) == 0:
        messagebox.showerror("Error", "Please at least select 1 service")
        return
    elif not _check_fee(app):
        messagebox.showerror("Error", "Please go to fee proposal page to complete the fee first")
        return
    elif data["Invoices"]["Fee"].get() == "Error":
        messagebox.showerror(title="Error",
                             message="There are error in the fee proposal section, please fix the fee section before generate the fee proposal")
        return
    elif not page in [1, 2, 3, 4, 5]:
        messagebox.showerror("Error", "Excess the maximum value of service, please contact administrator")
    pdf_list = [file for file in os.listdir(os.path.join(accounting_dir)) if
                "Fee Proposal" in str(file) and str(file).endswith(".pdf") and not "Mechanical Installation Fee Proposal" in str(file)]

    # case = _check_case(app)
    # if case == "Error":
    #     app.messagebox.show_error("Unsupported Version")
    #     return

    if len(pdf_list) != 0:
        current_pdf = sorted(pdf_list, key=lambda pdf: str(pdf).split(" ")[-1].split(".")[0])[-1]
        current_revision = current_pdf.split(" ")[-1].split(".")[0]
        if data["Fee Proposal"]["Reference"]["Revision"].get() == current_revision or data["Fee Proposal"]["Reference"][
            "Revision"].get() == str(int(current_revision) + 1):
            old_pdf_path = os.path.join(accounting_dir, current_pdf)
            def open_pdf():
                subprocess.call([adobe_address, old_pdf_path])

            _thread.start_new_thread(open_pdf, ())

            overwrite = messagebox.askyesno(f"Warming", f"Revision {current_revision} found, do you want to generate Revision {data['Fee Proposal']['Reference']['Revision'].get()}")
            if not overwrite:
                return
            else:
                for proc in psutil.process_iter():
                    if proc.name() == "Acrobat.exe":
                        proc.kill()
        else:
            messagebox.showerror("Error",
                                 f'Current revision is {current_revision}, you can not use revision {data["Fee Proposal"]["Reference"]["Revision"].get()}')
            return
    else:
        if not data["Fee Proposal"]["Reference"]["Revision"].get() == "1":
            messagebox.showerror("Error", "There is no other existing fee proposal found, can only have revision 1")
            return

    # total_fee = 0
    # total_ist = 0
    # for service in [value for value in data["Invoices"]["Details"].values() if
    #                 value["Service"].get() != "Variation" and value["Include"].get()]:
    #     total_fee += float(service["Fee"].get())
    #     total_ist += float(service["in.GST"].get())
    stage = len([s for s in data["Fee Proposal"]["Stage"].values() if s["Include"].get()])
    if stage == 0:
        messagebox.showerror("Error", "You need to select at least one stage")
        return
    shutil.copy(os.path.join(resource_dir, "xlsx", f"major_proposal_template_{page}_stage{stage}.xlsx"),
                os.path.join(accounting_dir, excel_name))
    excel = win32client.Dispatch("Excel.Application")
    try:
        excel.Visible = False
        # excel.ScreenUpdating = False
        # excel.DisplayAlerts = False
        # excel.EnableEvents = False
    except Exception as e:
        print(e)
        pass
    work_book = excel.Workbooks.Open(os.path.join(accounting_dir, excel_name))
    # messagebox.showerror("Error", e)
    try:
        work_sheets = work_book.Worksheets[0]
        address_to = data["Address_to"].get()
        work_sheets.Cells(27, 2).Value = data["Project Info"]["Project"]["Project Name"].get()
        work_sheets.Cells(31, 3).Value = _get_service_name(app) + " Fee Proposal"
        work_sheets.Cells(38, 1).Value = data["Project Info"][address_to]["Full Name"].get()
        work_sheets.Cells(39, 1).Value = data["Project Info"][address_to]["Company"].get()
        work_sheets.Cells(40, 1).Value = data["Project Info"][address_to]["Address"].get()
        work_sheets.Cells(43, 1).Value = data["Fee Proposal"]["Reference"]["Date"].get()

        work_sheets.Cells(45, 2).Value = data["Project Info"][address_to]["Full Name"].get()
        work_sheets.Cells(46, 2).Value = data["Project Info"][address_to]["Company"].get()
        work_sheets.Cells(47, 2).Value = data["Project Info"][address_to]["Address"].get()

        work_sheets.Cells(45, 8).Value = data["Project Info"]["Project"]["Quotation Number"].get()
        work_sheets.Cells(46, 8).Value = data["Fee Proposal"]["Reference"]["Date"].get()
        work_sheets.Cells(47, 8).Value = data["Fee Proposal"]["Reference"]["Revision"].get()

        first_name = get_first_name(data["Project Info"][address_to]["Full Name"].get())

        work_sheets.Cells(50, 1).Value = f"Dear {first_name},"

        work_sheets.Cells(51, 1).Value = "Re: " + data["Project Info"]["Project"]["Project Name"].get()
        work_sheets.Cells(53, 1).Value = f"Thank you for giving us the opportunity to submit this fee proposal for our {', '.join([key for key, value in data['Project Info']['Project']['Service Type'].items() if value['Include'].get()])} for the above project."

        work_sheets.Cells(86, 2).Value = "We have prepared this submission in response to an invitation of " + data["Project Info"][address_to]["Full Name"].get() + " for the provision of our consulting services for " + data["Project Info"]["Project"]["Project Name"].get()
        work_sheets.Cells(97, 2).Value = "Email sent from " + data["Project Info"][address_to]["Full Name"].get() + " on " + data["Fee Proposal"]["Reference"]["Date"].get() + " with Architectural Plans"

        cur_row = 97
        # for building in data["Project Info"]["Building Features"]["Details"]:
        #     if len(building["Space"].get())!=0:
        #         work_sheets.Cells(cur_row, 2).Value = building["Space"].get()
        #         cur_row+=1
        if not data["Project Info"]["Building Features"]["Notes"].get() is None:
            for line in data["Project Info"]["Building Features"]["Notes"].get().split("\n"):
                work_sheets.Cells(cur_row, 2).Value = line
                cur_row+=1

        cur_row+=1
        work_sheets.Cells(cur_row, 1).Value = "2.3"
        work_sheets.Cells(cur_row, 1).Font.Bold = True
        work_sheets.Cells(cur_row, 2).Value = "Deliverables at each stage"
        work_sheets.Cells(cur_row, 2).Font.Bold = True
        cur_row+=1

        n_stage = 0
        for value in data["Fee Proposal"]["Stage"].values():
            if not value["Include"].get():
                continue
            work_sheets.Cells(cur_row, 2).value = value["Service"].get()
            work_sheets.Cells(cur_row, 2).Font.Bold = True
            cur_row+=1
            n_stage+=1
            for item in value["Items"]:
                if item["Include"].get():
                    work_sheets.Cells(cur_row, 1).value = "•"
                    work_sheets.Cells(cur_row, 2).value = item["Item"].get()
                    cur_row+=1
            cur_row += 1
        cur_index = 3
        i = 0
        for key in app.conf["proposal_list"]:
            if not data["Invoices"]["Details"][key]["Include"].get():
                continue
            service = data["Fee Proposal"]["Scope"]["Major"][key]

            cur_row = 131 + i * row_per_page
            extra_list = conf["extra_list"]
            for extra in extra_list:
                if len([scope for scope in service[extra] if scope["Include"].get()]) == 0:
                    continue
                work_sheets.Cells(cur_row, 1).Value = "2." + str(cur_index)
                work_sheets.Cells(cur_row, 2).Value = key + "-" + extra
                work_sheets.Cells(cur_row, 1).Font.Bold = True
                work_sheets.Cells(cur_row, 2).Font.Bold = True
                cur_row+=1
                for scope in service[extra]:
                    if scope["Include"].get():
                        work_sheets.Cells(cur_row, 1).Value = "•"
                        for line in separate_line(scope["Item"].get()):
                            work_sheets.Cells(cur_row, 2).Value = line
                            cur_row+=1
                cur_row += 1
                cur_index += 1
            i += 1


        cur_row = 148 + page * row_per_page
        project_type = data["Project Info"]["Project"]["Project Type"].get()
        past_projects = json.load(open(past_projects_dir, encoding="utf-8"))[project_type]
        for i, project in enumerate(past_projects):
            work_sheets.Cells(cur_row + i, 1).Value = "•"
            work_sheets.Cells(cur_row + i, 2).Value = project

        _major_project_fee_section(data, work_sheets, n_stage, page, row_per_page)

        cur_row = 200 + page * (row_per_page+1)
        work_sheets.Cells(cur_row, 2).Value = "Re: "+data["Project Info"]["Project"]["Project Name"].get()
        work_sheets.ExportAsFixedFormat(0, os.path.join(accounting_dir, pdf_name))
        work_book.Close(True)
    except PermissionError:
        messagebox.showerror("Error", "Please close the preview or file before you use it")
    except FileNotFoundError:
        messagebox.showerror("Error", "Please Contact Administer, the app cant find the file")
    # except Exception as e:
    #     messagebox.showerror("Error", "Please Close the pdf before you generate a new one")
    #     print(e)
    else:
        app.data["State"]["Generate Proposal"].set(True)

        # avDoc.open(os.path.join(database_dir, pdf_name), os.path.join(database_dir, pdf_name))
        def open_pdf():
            subprocess.call([adobe_address, os.path.join(accounting_dir, pdf_name)])

        _thread.start_new_thread(open_pdf, ())
        # webbrowser.open(os.path.join(database_dir, pdf_name))
        save(app)
        config_state(app)
        app.log.log_fee_proposal(app)
        config_log(app)
    try:
        excel.Visible = True
        # excel.ScreenUpdating = True
        # excel.DisplayAlerts = True
        # excel.EnableEvents = True
        work_book.Close(True)
        # adobe.close(0)
    except Exception as e:
        print(e)

def _major_project_fee_section(data, work_sheets, n_stage, page, row_per_page):

    cur_row = 180 + page * row_per_page
    stage_total = [0]*4
    total = 0
    gst = 0
    total_gst = 0

    stage_list = [ key for key, value in data["Fee Proposal"]["Stage"].items() if value["Include"].get()]

    if n_stage == 1:
        work_sheets.Cells(cur_row, 4).Value = data["Fee Proposal"]["Stage"][stage_list[0]]["Service"].get()
    elif n_stage == 2:
        work_sheets.Cells(cur_row, 4).Value = data["Fee Proposal"]["Stage"][stage_list[0]]["Service"].get()
        work_sheets.Cells(cur_row, 6).Value = data["Fee Proposal"]["Stage"][stage_list[1]]["Service"].get()
    elif n_stage == 3:
        work_sheets.Cells(cur_row, 4).Value = data["Fee Proposal"]["Stage"][stage_list[0]]["Service"].get()
        work_sheets.Cells(cur_row, 5).Value = data["Fee Proposal"]["Stage"][stage_list[1]]["Service"].get()
        work_sheets.Cells(cur_row, 6).Value = data["Fee Proposal"]["Stage"][stage_list[2]]["Service"].get()
    elif n_stage == 4:
        work_sheets.Cells(cur_row, 4).Value = data["Fee Proposal"]["Stage"][stage_list[0]]["Service"].get()
        work_sheets.Cells(cur_row, 5).Value = data["Fee Proposal"]["Stage"][stage_list[1]]["Service"].get()
        work_sheets.Cells(cur_row, 6).Value = data["Fee Proposal"]["Stage"][stage_list[2]]["Service"].get()
        work_sheets.Cells(cur_row, 7).Value = data["Fee Proposal"]["Stage"][stage_list[3]]["Service"].get()

    cur_row+=2

    for service in conf["proposal_list"]:
        if not data["Invoices"]["Details"][service]["Include"].get():
            continue
        total += float(data["Invoices"]["Details"][service]["Fee"].get())
        gst += float(data["Invoices"]["Details"][service]["in.GST"].get()) - float(data["Invoices"]["Details"][service]["Fee"].get())
        total_gst += float(data["Invoices"]["Details"][service]["in.GST"].get())
        work_sheets.Cells(cur_row, 2).Value = service

        if n_stage == 1:
            # if data["Invoices"]["Details"][service]["Expand"].get():
            for j in range(n_stage):
                item = data["Invoices"]["Details"][service]["Content"][j]
                item_fee = float(item["Fee"].get()) if len(item["Fee"].get()) !=0 else 0
                work_sheets.Cells(cur_row, 4 + j).Value = item_fee
                stage_total[j] += item_fee
            work_sheets.Cells(cur_row, 8).Value = float(data["Invoices"]["Details"][service]["Fee"].get())
            cur_row +=1
        elif n_stage == 2:
            # if data["Invoices"]["Details"][service]["Expand"].get():
            for j in range(n_stage):
                item = data["Invoices"]["Details"][service]["Content"][j]
                item_fee = float(item["Fee"].get()) if len(item["Fee"].get()) !=0 else 0
                work_sheets.Cells(cur_row, 4 + 2*j).Value = item_fee
                stage_total[j] += item_fee
            work_sheets.Cells(cur_row, 8).Value = float(data["Invoices"]["Details"][service]["Fee"].get())
            cur_row +=1
        elif n_stage == 3:
            # if data["Invoices"]["Details"][service]["Expand"].get():
            for j in range(n_stage):
                item = data["Invoices"]["Details"][service]["Content"][j]
                item_fee = float(item["Fee"].get()) if len(item["Fee"].get()) !=0 else 0
                work_sheets.Cells(cur_row, 4 + j).Value = item_fee
                stage_total[j] += item_fee
            work_sheets.Cells(cur_row, 7).Value = float(data["Invoices"]["Details"][service]["Fee"].get())
            cur_row +=1
        elif n_stage == 4:
            # if data["Invoices"]["Details"][service]["Expand"].get():
            for j in range(n_stage):
                item = data["Invoices"]["Details"][service]["Content"][j]
                item_fee = float(item["Fee"].get()) if len(item["Fee"].get()) !=0 else 0
                work_sheets.Cells(cur_row, 4 + j).Value = item_fee
                stage_total[j] += item_fee
            work_sheets.Cells(cur_row, 8).Value = float(data["Invoices"]["Details"][service]["Fee"].get())
            cur_row +=1

    if n_stage == 1:
        cur_row = 182 + (row_per_page+1)*page
        work_sheets.Cells(cur_row - 1, 8).Value = total
        work_sheets.Cells(cur_row, 8).Value = gst
        work_sheets.Cells(cur_row + 1, 8).Value = total_gst
    elif n_stage == 2:
        cur_row = 182 + (row_per_page+1)*page
        work_sheets.Cells(cur_row, 4).Value = stage_total[0]
        work_sheets.Cells(cur_row, 6).Value = stage_total[1]
        work_sheets.Cells(cur_row, 8).Value = total
        work_sheets.Cells(cur_row + 1, 8).Value = gst
        work_sheets.Cells(cur_row + 2, 8).Value = total_gst
    elif n_stage == 3:
        cur_row = 182 + (row_per_page+1)*page
        work_sheets.Cells(cur_row, 4).Value = stage_total[0]
        work_sheets.Cells(cur_row, 5).Value = stage_total[1]
        work_sheets.Cells(cur_row, 6).Value = stage_total[2]
        work_sheets.Cells(cur_row, 7).Value = total
        work_sheets.Cells(cur_row + 1, 7).Value = gst
        work_sheets.Cells(cur_row + 2, 7).Value = total_gst
    elif n_stage == 4:
        cur_row = 182 + (row_per_page+1)*page
        work_sheets.Cells(cur_row, 4).Value = stage_total[0]
        work_sheets.Cells(cur_row, 5).Value = stage_total[1]
        work_sheets.Cells(cur_row, 6).Value = stage_total[2]
        work_sheets.Cells(cur_row, 7).Value = stage_total[3]
        work_sheets.Cells(cur_row, 8).Value = total
        work_sheets.Cells(cur_row + 1, 8).Value = gst
        work_sheets.Cells(cur_row + 2, 8).Value = total_gst
    # work_sheets.Cells(cur_row, 8).Value = total
    # work_sheets.Cells(cur_row+1, 8).Value = gst
    # work_sheets.Cells(cur_row+2, 8).Value = total_gst

def _get_proposal_name(app):
    data = app.data
    service_name = _get_service_name(app)
    return f'{service_name} Fee Proposal for {data["Project Info"]["Project"]["Project Name"].get()} Rev {data["Fee Proposal"]["Reference"]["Revision"].get()}'

def _get_service_name(app):
    data = app.data
    service_list = []
    for service in app.conf["proposal_list"]:
        if data["Invoices"]["Details"][service]["Include"].get():
            service_list.append(service)

    if len(service_list) == 1:
        service_name = service_list[0].split(" ")[0]
    elif len(service_list) == 2:
        service_name = f'{service_list[0].split(" ")[0]} and {service_list[1].split(" ")[0]}'
    elif len(service_list) == 3:
        service_name = f'{service_list[0].split(" ")[0]}, {service_list[1].split(" ")[0]} and {service_list[2].split(" ")[0]}'
    else:
        service_name = "Multi Service"
    return service_name
def excel_print_invoice(app, inv):
    data = app.data
    check_accounting_folder(app)
    excel_name = f'PCE INV {data["Invoices Number"][inv]["Number"].get()}.xlsx'
    invoice_name = f'PCE INV {data["Invoices Number"][inv]["Number"].get()}.pdf'
    # database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    accounting_dir = os.path.join(app.conf["accounting_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    resource_dir = app.conf["resource_dir"]
    adobe_address = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
    # win32com.client.makepy.GenerateFromTypeLibSpec('Acrobat')
    # adobe = win32com.client.DispatchEx('AcroExch.App')
    # avDoc = win32client.dynamic.Dispatch("AcroExch.AVDoc")

    if len(data["Invoices Number"][inv]["Number"].get()) == 0:
        messagebox.showerror("Error", "You need to generate a invoice number before you generate the Invoice")
        return
    elif data["Invoices Number"][inv]["Fee"].get() == "0" or len(data["Invoices Number"][inv]["Fee"].get())==0:
        process = app.messagebox.ask_yes_no("The Invoice amount is 0, Do you want to process?")
        if not process:
            return
    rewrite = True
    if os.path.exists(os.path.join(accounting_dir, invoice_name)):
        old_pdf_path=os.path.join(accounting_dir, invoice_name)
        # avDoc.open(old_pdf_path, old_pdf_path)
        def open_pdf():
            subprocess.call([adobe_address, old_pdf_path])
        _thread.start_new_thread(open_pdf, ())
        rewrite = messagebox.askyesno("Warming", f"Existing file PCE INV {invoice_name} do you want to rewrite?")
        if not rewrite:
            return
        else:
            for proc in psutil.process_iter():
                if proc.name() == "Acrobat.exe":
                    proc.kill()
    if rewrite:
        shutil.copy(os.path.join(resource_dir, "xlsx", f"invoice_template.xlsx"),
                    os.path.join(accounting_dir, excel_name))
        excel = win32client.Dispatch("Excel.Application")
        excel.Visible = False
        # excel.ScreenUpdating = False
        # excel.DisplayAlerts = False
        # excel.EnableEvents = False
        work_book = excel.Workbooks.Open(os.path.join(accounting_dir, excel_name))
        try:
            work_sheets = work_book.Worksheets[0]
            address_to = data["Address_to"].get()

            cur_row = 5
            for key in ["Full Name", "Company", "Address", "ABN"]:
                if len(data["Project Info"][address_to][key].get()) != 0:
                    if key == "ABN":
                        work_sheets.Cells(cur_row, 1).Value = "ABN\\ACN: " + data["Project Info"][address_to][key].get()
                    else:
                        work_sheets.Cells(cur_row, 1).Value = data["Project Info"][address_to][key].get()
                    cur_row += 1

            # work_sheets.Cells(5, 1).Value = data["Project Info"][address_to]["Full Name"].get()
            # work_sheets.Cells(6, 1).Value = data["Project Info"][address_to]["Company"].get()
            # work_sheets.Cells(7, 1).Value = data["Project Info"][address_to]["Address"].get()
            work_sheets.Cells(4, 10).Value = datetime.today().strftime("%d-%b-%Y")
            work_sheets.Cells(5, 10).Value = data["Invoices Number"][inv]["Number"].get()
            work_sheets.Cells(12, 1).Value = data["Project Info"]["Project"]["Project Name"].get()
            cur_row = 14
            total_fee = 0
            total_inGST = 0
            for service in data["Invoices"]["Details"].values():
                if not service["Include"].get() or service["Service"].get()=="Variation":
                    continue
                # if service["Expand"].get():
                work_sheets.Cells(cur_row, 1).Value = service["Service"].get() + " design and documentation"
                # work_sheets.Cells(cur_row, 7).Value = "Amount"
                # work_sheets.Cells(cur_row, 8).Value = "This pay"
                cur_row += 1
                for item in service["Content"]:
                    if len(item["Fee"].get()) != 0:
                        work_sheets.Cells(cur_row, 1).Value = item["Service"].get()
                        work_sheets.Cells(cur_row, 7).Value = item["Fee"].get()
                        if item["Number"].get() == f"INV{str(inv+1)}":
                            work_sheets.Cells(cur_row, 8).Value = item["Fee"].get()
                            work_sheets.Cells(cur_row, 9).Value = item["Fee"].get()
                            work_sheets.Cells(cur_row, 10).Value = item["in.GST"].get()
                            total_fee += float(item["Fee"].get())
                            total_inGST += float(item["in.GST"].get()) if len(item["in.GST"].get()) != 0 else 0
                        cur_row += 1
                work_sheets.Cells(cur_row, 6).Value = "Total: "
                work_sheets.Cells(cur_row, 6).Font.Bold = True
                work_sheets.Cells(cur_row, 7).Value = service["Fee"].get()
                work_sheets.Cells(cur_row, 7).Font.Bold = True
                cur_row += 2
                # else:
                #     work_sheets.Cells(cur_row, 1).Value = service["Service"].get() + " design and documentation"
                #     # work_sheets.Cells(cur_row+1, 2).Value = "Payment Instalments"
                #     # work_sheets.Cells(cur_row+1, 7).Value = "Amount"
                #     # work_sheets.Cells(cur_row+1, 8).Value = "This pay"
                #     # work_sheets.Cells(cur_row+2, 1).Value = "The Full amount for design and documentation"
                #     work_sheets.Cells(cur_row, 7).Value = service["Fee"].get()
                #     if service["Number"].get() == f"INV{str(inv+1)}":
                #         work_sheets.Cells(cur_row, 8).Value = service["Fee"].get()
                #         work_sheets.Cells(cur_row, 9).Value = service["Fee"].get()
                #         work_sheets.Cells(cur_row, 10).Value = service["in.GST"].get()
                #         total_fee += float(service["Fee"].get()) if len(service["Fee"].get()) !=0 else 0
                #         total_inGST += float(service["in.GST"].get()) if len(service["Fee"].get()) !=0 else 0
                #     cur_row+=2
                # cur_row+=1
            cur_row+=1
            if len(data["Invoices"]["Details"]["Variation"]["Fee"].get()) != 0 and data["Invoices"]["Details"]["Variation"]["Fee"].get() != "0":
                work_sheets.Cells(cur_row, 1).Value = "Variation"
                cur_row+=1
                for service in data["Invoices"]["Details"]["Variation"]["Content"]:
                    if service["Fee"].get()==0:
                        continue
                    work_sheets.Cells(cur_row, 1).Value = service["Service"].get()
                    work_sheets.Cells(cur_row, 7).Value = service["Fee"].get()
                    if service["Number"].get() == f"INV{str(inv+1)}":
                        work_sheets.Cells(cur_row, 8).Value = service["Fee"].get()
                        work_sheets.Cells(cur_row, 9).Value = service["Fee"].get()
                        work_sheets.Cells(cur_row, 10).Value = service["in.GST"].get()
                        total_fee += float(service["Fee"].get()) if len(service["Fee"].get()) !=0 else 0
                        total_inGST += float(service["in.GST"].get()) if len(service["in.GST"].get()) !=0 else 0
                    cur_row += 1
            work_sheets.Cells(44, 9).Value = data["Invoices Number"][0]["Number"].get()

            # for service in data["Variation"]:
            #     if len(service["Service"].get())==0 and len(service["Fee"].get())==0:
            #         continue
            #     work_sheets.Cells(cur_row, 1).Value = service["Service"].get()
            #     work_sheets.Cells(cur_row, 7).Value = service["Fee"].get()
            #     if service["Number"].get() == inv:
            #         work_sheets.Cells(cur_row, 8).Value = service["Fee"].get()
            #         work_sheets.Cells(cur_row, 9).Value = service["Fee"].get()
            #         work_sheets.Cells(cur_row, 10).Value = service["in.GST"].get()
            #         total_fee += float(service["Fee"].get())
            #         total_inGST += float(service["in.GST"].get())
            #     cur_row += 2

            work_sheets.Cells(39, 9).Value = str(total_fee)
            work_sheets.Cells(39, 10).Value = str(total_inGST)
            work_sheets.ExportAsFixedFormat(0, os.path.join(accounting_dir, invoice_name))
            work_book.Close(True)
        except PermissionError:
            messagebox.showerror("Error", "Please close the preview or file before you use it")
        except Exception as e:
            messagebox.showerror("Error", "Please Close the pdf before you generate a new one")
            print(e)
        else:
            invoice_path = os.path.join(accounting_dir, invoice_name)

            def open_pdf():
                subprocess.call([adobe_address, invoice_path])

            _thread.start_new_thread(open_pdf, ())
            # webbrowser.open(invoice_path)
            # avDoc.open(invoice_path, invoice_path)
            data["Remittances"][inv]["Preview_Upload"].set(True)
            save(app)
            app.log.log_generate_invoices(app, inv)
            config_state(app)
            config_log(app)
        try:
            excel.Visible = True
            # excel.ScreenUpdating = True
            # excel.DisplayAlerts = True
            # excel.EnableEvents = True
            work_book.Close(True)
        except Exception as e:
            print(e)

def preview_installation_fee_proposal(app):
    data = app.data
    check_accounting_folder(app)
    pdf_name = f"Mechanical Installation Fee Proposal for {data['Project Info']['Project']['Project Name'].get()} Rev {data['Fee Proposal']['Installation Reference']['Revision'].get()}.pdf"
    excel_name = f'Mechanical Installation {data["Project Info"]["Project"]["Project Name"].get()} Back Up.xlsx'
    # database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    accounting_dir = os.path.join(app.conf["accounting_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    resource_dir = app.conf["resource_dir"]
    adobe_address = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
    proposal_type = data["Project Info"]["Project"]["Proposal Type"].get()
    past_projects_dir = os.path.join(app.conf["database_dir"], "past_projects.json")

    # if not _check_fee(app):
    #     messagebox.showerror("Error", "Please go to fee proposal page to complete the fee first")
    #     return
    # if data["Invoices"]["Fee"].get() == "Error":
    #     messagebox.showerror(title="Error", message="There are error in the fee proposal section, please fix the fee section before generate the fee proposal")
    #     return

    pdf_list = [file for file in os.listdir(os.path.join(accounting_dir)) if
                str(file).endswith(".pdf") and "Mechanical Installation Fee Proposal" in str(file)]

    if len(pdf_list) != 0:
        current_pdf = sorted(pdf_list, key=lambda pdf: str(pdf).split(" ")[-1].split(".")[0])[-1]
        current_revision = current_pdf.split(" ")[-1].split(".")[0]
        if data["Fee Proposal"]["Installation Reference"]["Revision"].get() == current_revision or data["Fee Proposal"]["Installation Reference"]["Revision"].get() == str(int(current_revision) + 1):
            old_pdf_path = os.path.join(accounting_dir, current_pdf)
            # avDoc.open(old_pdf_path, old_pdf_path)
            def open_pdf():
                subprocess.call([adobe_address, old_pdf_path])
            _thread.start_new_thread(open_pdf, ())

            overwrite = messagebox.askyesno(f"Warming", f"Revision {current_revision} found, do you want to generate Revision {data['Fee Proposal']['Installation Reference']['Revision'].get()}")
            if not overwrite:
                return
            else:
                # avDoc.close(0)
                for proc in psutil.process_iter():
                    if proc.name() == "Acrobat.exe":
                        proc.kill()
        else:
            messagebox.showerror("Error",
                                 f'Current revision is {current_revision}, you can not use revision {data["Fee Proposal"]["Installation Reference"]["Revision"].get()}')
            return
    else:
        if not data["Fee Proposal"]["Installation Reference"]["Revision"].get() == "1":
            messagebox.showerror("Error", "There is no other existing fee proposal found, can only have revision 1")
            return


    shutil.copy(os.path.join(resource_dir, "xlsx", f"installation_fee_proposal_template.xlsx"),
                os.path.join(accounting_dir, excel_name))
    excel = win32client.Dispatch("Excel.Application")
    try:
        excel.Visible = False
        # excel.ScreenUpdating = False
        # excel.DisplayAlerts = False
        # excel.EnableEvents = False
    except Exception as e:
        print(e)
    work_book = excel.Workbooks.Open(os.path.join(accounting_dir, excel_name))
        # messagebox.showerror("Error", e)
    try:
        work_sheets = work_book.Worksheets[0]
        address_to = data["Address_to"].get()
        work_sheets.Cells(1, 2).Value = data["Project Info"][address_to]["Full Name"].get()
        work_sheets.Cells(2, 2).Value = data["Project Info"][address_to]["Company"].get()
        work_sheets.Cells(3, 2).Value = data["Project Info"][address_to]["Address"].get()
        first_name = get_first_name(data["Project Info"][address_to]["Full Name"].get())
        work_sheets.Cells(5, 1).Value = f'Dear {first_name},'
        work_sheets.Cells(1, 8).Value = data["Project Info"]["Project"]["Quotation Number"].get()
        work_sheets.Cells(2, 8).Value = data["Fee Proposal"]["Installation Reference"]["Date"].get()
        work_sheets.Cells(3, 8).Value = data["Fee Proposal"]["Installation Reference"]["Revision"].get()
        work_sheets.Cells(6, 1).Value = "Re: " + data["Project Info"]["Project"]["Project Name"].get()



        cur_row = 51
        for i, drawing in enumerate(data["Project Info"]["Drawing"]):
            work_sheets.Cells(cur_row+i, 2).Value = drawing["Drawing Number"].get()
            work_sheets.Cells(cur_row+i, 4).Value = drawing["Drawing Name"].get()
            work_sheets.Cells(cur_row+i, 8).Value = drawing["Revision"].get()

        cur_row = 96
        for i, extra in enumerate(data["Fee Proposal"]["Scope"][proposal_type]["Installation"].keys()):

            work_sheets.Cells(cur_row, 1).Value = f"2.{i+1}"
            if extra == "Extent":
                work_sheets.Cells(cur_row, 2).Value = "MECHANICAL SERVICES INSTALLATION– EXTENT"
            elif extra == "Clarifications":
                work_sheets.Cells(cur_row, 2).Value = "MECHANICAL SERVICES – EXCLUSION"
            else:
                work_sheets.Cells(cur_row, 2).Value = "MECHANICAL EQUIPMENTS LIST"
            work_sheets.Cells(cur_row, 1).Font.Bold = True
            work_sheets.Cells(cur_row, 2).Font.Bold = True
            cur_row+=1
            for item in data["Fee Proposal"]["Scope"][proposal_type]["Installation"][extra]:
                if item["Include"].get():
                    work_sheets.Cells(cur_row, 1).Value = "•"
                    for line in separate_line(item["Item"].get()):
                        work_sheets.Cells(cur_row, 2).Value = line
                        cur_row += 1
            cur_row+=1

        cur_row = 168
        project_type = data["Project Info"]["Project"]["Project Type"].get()
        past_projects = json.load(open(past_projects_dir, encoding="utf-8"))[project_type]
        for i, project in enumerate(past_projects):
            work_sheets.Cells(cur_row + i, 1).Value = "•"
            work_sheets.Cells(cur_row + i, 2).Value = project

        cur_row = 181
        if not data["Fee Proposal"]["Installation Reference"]["Program"].get() is None:
            for line in data["Fee Proposal"]["Installation Reference"]["Program"].get().split("\n"):
                work_sheets.Cells(cur_row, 1).Value = "•"
                work_sheets.Cells(cur_row, 2).Value = line
                cur_row+=1

        cur_row = 195
        for i, content in enumerate(data["Invoices"]["Details"]["Installation"]["Content"]):
            work_sheets.Cells(cur_row+i, 2).Value = content["Service"].get()
            work_sheets.Cells(cur_row+i, 6).Value = content["Fee"].get()
            work_sheets.Cells(cur_row+i, 7).Value = content["in.GST"].get()

        cur_row = 199
        work_sheets.Cells(cur_row, 6).Value = data["Invoices"]["Details"]["Installation"]["Fee"].get()
        work_sheets.Cells(cur_row, 7).Value = data["Invoices"]["Details"]["Installation"]["in.GST"].get()

        cur_row = 214
        work_sheets.Cells(cur_row, 2).Value = "Re: "+data["Project Info"]["Project"]["Project Name"].get()
        work_sheets.ExportAsFixedFormat(0, os.path.join(accounting_dir, pdf_name))
        work_book.Close(True)
    except PermissionError:
        print(e)
        messagebox.showerror("Error", "Please close the preview or file before you use it")
    except FileNotFoundError:
        print(e)
        messagebox.showerror("Error", "Please Contact Administer, the app cant find the file")
    except Exception as e:
        print(e)
        messagebox.showerror("Error", "Please Close the pdf before you generate a new one")
    else:
        app.data["State"]["Generate Proposal"].set(True)
        # avDoc.open(os.path.join(database_dir, pdf_name), os.path.join(database_dir, pdf_name))
        def open_pdf():
            subprocess.call([adobe_address, os.path.join(accounting_dir, pdf_name)])
        _thread.start_new_thread(open_pdf, ())
        # webbrowser.open(os.path.join(database_dir, pdf_name))
        save(app)
        config_state(app)
        app.log.log_fee_proposal(app)
        config_log(app)
    try:
        excel.Visible = True
        # excel.ScreenUpdating = True
        # excel.DisplayAlerts = True
        # excel.EnableEvents = True
        work_book.Close(True)
        # adobe.close(0)
    except Exception as e:
        print(e)


def email_fee_proposal(app, *args):
    data = app.data
    check_accounting_folder(app)
    # database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    accounting_dir = os.path.join(app.conf["accounting_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    pdf_list = [file for file in os.listdir(os.path.join(accounting_dir)) if
                "Fee Proposal" in str(file) and str(file).endswith(".pdf") and not "Mechanical Installation Fee Proposal" in str(file)]

    pdf_name = sorted(pdf_list, key=lambda pdf: str(pdf).split(" ")[-1].split(".")[0])[-1]

    if not data["State"]["Generate Proposal"].get():
        messagebox.showerror("Error", "Please Generate a pdf first")
        return
    elif not os.path.exists(os.path.join(accounting_dir, pdf_name)):
        messagebox.showerror("Error",
                             f'Python cant found fee proposal for {data["Project Info"]["Project"]["Quotation Number"].get()} revision {data["Fee Proposal"]["Reference"]["Revision"].get()}')
        return

    if not "OUTLOOK.EXE" in (p.name() for p in psutil.process_iter()):
        os.startfile("outlook")

    # user_email_dic = json.load(open(os.path.join(app.conf["database_dir"], "user_email.json")))
    # pythoncom.pythoncom.CoInitialize()

    ol = win32client.Dispatch("Outlook.Application", pythoncom.CoInitialize())
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    # newmail.From = user_email_dic[app.user]
    newmail.Subject = f'{data["Project Info"]["Project"]["Quotation Number"].get()}-{_get_proposal_name(app)}'
    newmail.To = f'{data["Project Info"]["Client"]["Contact Email"].get()}; {data["Project Info"]["Main Contact"]["Contact Email"].get()}'
    newmail.BCC = "bridge@pcen.com.au"
    newmail.GetInspector()
    index = newmail.HTMLbody.find(">", newmail.HTMLbody.find("<body"))

    address_to = data["Address_to"].get()
    full_name = data["Project Info"][address_to]["Full Name"].get()
    first_name = get_first_name(full_name)

    message = f"""
    Dear {first_name},<br><br>

    I hope this email finds you well. Please find the attached fee proposal to this email.<br>

    If you have any questions or need more information regarding the proposal, please don't hesitate to reach out.<br>

    I look forward to hearing from you soon.<br>
    """
    newmail.HTMLbody = newmail.HTMLbody[:index + 1] + message + newmail.HTMLbody[index + 1:]
    newmail.Attachments.Add(os.path.join(accounting_dir, pdf_name))
    newmail.Display()
    save(app)
    config_state(app)
    return True
def convert_time_format(time_string):
    year, month, day = time_string.split("-")
    if day == "01":
        return datetime(int(year), int(month), int(day)).strftime("%d<sup>st</sup> %B %Y")
    elif day == "02":
        return datetime(int(year), int(month), int(day)).strftime("%d<sup>nd</sup> %B %Y")
    else:
        return datetime(int(year), int(month), int(day)).strftime("%d<sup>th</sup> %B %Y")
def chase(app, *args):
    data = app.data
    # database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    accounting_dir = os.path.join(app.conf["accounting_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    pdf_list = [file for file in os.listdir(os.path.join(accounting_dir)) if
                "Fee Proposal" in str(file) and str(file).endswith(".pdf") and not "Mechanical Installation Fee Proposal" in str(file)]

    pdf_name = sorted(pdf_list, key=lambda pdf: str(pdf).split(" ")[-1].split(".")[0])[-1]
    if not data["State"]["Email to Client"].get():
        messagebox.showerror("Error", "Please Sent the fee proposal to client first")
        return

    if not "OUTLOOK.EXE" in (p.name() for p in psutil.process_iter()):
        os.startfile("outlook")
    ol = win32client.Dispatch("Outlook.Application", pythoncom.CoInitialize())
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = f'Re: {data["Project Info"]["Project"]["Quotation Number"].get()}-{data["Project Info"]["Project"]["Project Name"].get()}'
    newmail.To = f'{data["Project Info"]["Client"]["Contact Email"].get()}; {data["Project Info"]["Main Contact"]["Contact Email"].get()}'
    newmail.CC = "felix@pcen.com.au"
    newmail.BCC = "bridge@pcen.com.au"
    newmail.GetInspector()
    index = newmail.HTMLbody.find(">", newmail.HTMLbody.find("<body"))

    address_to = data["Address_to"].get()
    full_name = data["Project Info"][address_to]["Full Name"].get()
    first_name = get_first_name(full_name)

    message = f"""
    Dear {first_name},<br><br>
    
    
    I hope this email finds you well.<br><br>
    
    I am writing to follow up on the fee proposal we sent on <b>{convert_time_format(data["Email"]["Fee Proposal"].get())}, attached</b>. I wanted to check if there is any update. <br><br>

    If you need any further clarification, please do not hesitate to contact us.<br><br>
    
    Looking forward to hearing from you and contribute to the project.<br><br>
    """
    newmail.HTMLbody = newmail.HTMLbody[:index + 1] + message + newmail.HTMLbody[index + 1:]
    newmail.Attachments.Add(os.path.join(accounting_dir, pdf_name))
    newmail.Display()
    save(app)
    config_state(app)
def email_invoice(app, inv):
    data = app.data
    check_accounting_folder(app)
    # database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    accounting_dir = os.path.join(app.conf["accounting_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    pdf_name = f'PCE INV {data["Invoices Number"][inv]["Number"].get()}.pdf'

    if not os.path.exists(os.path.join(accounting_dir, pdf_name)):
        messagebox.showerror("Error", f"Please generate the invoice before you send to client")
        return
    # if not data["State"]["Fee Accepted"].get():
    #     messagebox.showerror("Error", "Please Sent the fee proposal to client first")
    #     return


    if not "OUTLOOK.EXE" in (p.name() for p in psutil.process_iter()):
        os.startfile("outlook")
    ol = win32client.Dispatch("Outlook.Application", pythoncom.CoInitialize())
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = f'PCE INV {data["Invoices Number"][inv]["Number"].get()}-{data["Project Info"]["Project"]["Project Name"].get()}'
    newmail.To = f'{data["Project Info"]["Client"]["Contact Email"].get()}; {data["Project Info"]["Main Contact"]["Contact Email"].get()}'
    newmail.BCC = "bridge@pcen.com.au"
    newmail.GetInspector()
    index = newmail.HTMLbody.find(">", newmail.HTMLbody.find("<body"))

    address_to = data["Address_to"].get()
    full_name = data["Project Info"][address_to]["Full Name"].get()
    first_name = get_first_name(full_name)

    message = f"""
    Hi {first_name},<br><br>
    
    I hope this email finds you well and energized for the week ahead.<br>

    Please find the attached invoice for client payment. We appreciate your prompt attention to this matter.<br>

    If you have any questions or concerns regarding the invoice, please do not hesitate to contact Felix and me. We are happy to discuss further.<br>

    Thank you for your time and consideration.<br>
    """
    newmail.HTMLbody = newmail.HTMLbody[:index + 1] + message + newmail.HTMLbody[index + 1:]
    newmail.Display()
    newmail.Attachments.Add(os.path.join(accounting_dir, pdf_name))
def email_installation_proposal(app):
    data = app.data
    check_accounting_folder(app)
    # database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    accounting_dir = os.path.join(app.conf["accounting_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    pdf_list = [file for file in os.listdir(os.path.join(accounting_dir)) if
                "Mechanical Installation Fee Proposal" in str(file) and str(file).endswith(".pdf")]

    pdf_name = sorted(pdf_list, key=lambda pdf: str(pdf).split(" ")[-1].split(".")[0])[-1]

    if not os.path.exists(os.path.join(accounting_dir, pdf_name)):
        messagebox.showerror("Error",
                             f'Python cant found fee proposal for {data["Project Info"]["Project"]["Quotation Number"].get()} revision {data["Fee Proposal"]["Reference"]["Revision"].get()}')
        return

    if not "OUTLOOK.EXE" in (p.name() for p in psutil.process_iter()):
        os.startfile("outlook")

    # user_email_dic = json.load(open(os.path.join(app.conf["database_dir"], "user_email.json")))

    ol = win32client.Dispatch("Outlook.Application", pythoncom.CoInitialize())
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    # newmail.From = user_email_dic[app.user]
    newmail.Subject = f'{data["Project Info"]["Project"]["Quotation Number"].get()}-Mechanical Installation Fee Proposal for {data["Project Info"]["Project"]["Project Name"].get()}'
    newmail.To = f'{data["Project Info"]["Client"]["Contact Email"].get()}; {data["Project Info"]["Main Contact"]["Contact Email"].get()}'
    newmail.BCC = "bridge@pcen.com.au"
    newmail.GetInspector()
    index = newmail.HTMLbody.find(">", newmail.HTMLbody.find("<body"))

    address_to = data["Address_to"].get()
    full_name = data["Project Info"][address_to]["Full Name"].get()
    first_name = get_first_name(full_name)

    message = f"""
        Dear {first_name},<br><br>

        I hope this email finds you well. Please find the attached fee proposal to this email.<br>

        If you have any questions or need more information regarding the proposal, please don't hesitate to reach out. I am happy to provide you with whatever information you need.<br>

        I look forward to hearing from you soon.<br>
        """
    newmail.HTMLbody = newmail.HTMLbody[:index + 1] + message + newmail.HTMLbody[index + 1:]
    newmail.Attachments.Add(os.path.join(accounting_dir, pdf_name))
    newmail.Display()
    save(app)
    config_state(app)
    return True


def get_invoice_item(app):
    data = app.data
    res = [[] for _ in range(6)]
    for i in range(6):
        for key, service in data["Invoices"]["Details"].items():
            if not service["Include"].get():
                continue
            # if service["Expand"].get():
            for j in range(app.conf["n_items"]):
                if len(service["Content"][j]["Service"].get()) == 0 or len(service["Content"][j]["Fee"].get()) == 0:
                    continue
                if service["Content"][j]["Number"].get() == f"INV{i+1}":
                    res[i].append(
                        {
                            "Item": service["Content"][j]["Service"].get(),
                            "Fee": service["Content"][j]["Fee"].get(),
                            "in.GST": service["Content"][j]["in.GST"].get()
                        }
                    )
            # else:
            #     if service["Number"].get() == f"INV{i+1}":
            #         if len(service["Fee"].get()) != 0:
            #             res[i].append(
            #                 {
            #                     "Item": service["Service"].get(),
            #                     "Fee": service["Fee"].get(),
            #                     "in.GST": service["in.GST"].get()
            #                 }
            #             )
    return res

def update_app_invoices(app, inv_list):
    data = app.data
    invoices_dir = os.path.join(app.conf["database_dir"], "invoices.json")
    invoices_json = json.load(open(invoices_dir))

    bills_dir = os.path.join(app.conf["database_dir"], "bills.json")
    bills_json = json.load(open(bills_dir))

    new_bill = {}
    for state in inv_list["Bills"].keys():
        new_bill[state] = dict()
        for bill in inv_list["Bills"][state].keys():
            new_bill[state][bill.split("-")[0]] = inv_list["Bills"][state][bill]
    inv_list["Bills"] = new_bill

    for state, inv_dict in inv_list["Invoices"].items():
        for inv_number, value in inv_dict.items():
            if state == "PAID":
                invoices_json[inv_number] = "Paid"
            elif state == "VOIDED":
                invoices_json[inv_number] = "Voided"

    for state, bill_dict in inv_list["Bills"].items():
        # bill_number = bill_number.split("-")[0]
        for bill_number, value in bill_dict.items():
            if state == "DRAFT":
                bills_json[bill_number] = "Draft"
            elif state == "SUBMITTED":
                bills_json[bill_number] = "Awaiting Approval"
            elif state == "AUTHORISED":
                bills_json[bill_number] = "Awaiting Payment"
            elif state == "PAID":
                bills_json[bill_number] = "Paid"
            elif state == "VOIDED":
                bills_json[bill_number] = "Voided"

    for value in data["Invoices Number"]:
        if value["Number"].get() in inv_list["Invoices"]["PAID"].keys():
            value["State"].set("Paid")
        elif value["Number"].get() in inv_list["Invoices"]["VOIDED"].keys():
            value["State"].set("Voided")

    for value in data["Bills"]["Details"].values():
        for item in value["Content"]:
            if len(item["Number"].get()) == 0:
                continue
            bill_number = data["Project Info"]["Project"]["Project Number"].get() + item["Number"].get()
            if bill_number in inv_list["Bills"]["DRAFT"].keys():
                item["State"].set("Draft")
                item["no.GST"].set(False)
                item["Fee"].set("")
            elif bill_number in inv_list["Bills"]["SUBMITTED"].keys():
                item["State"].set("Awaiting Approval")
                # value = inv_list["Bills"]["SUBMITTED"][bill_number]
                # if value["line_amount_types"] == 'NoTax':
                #     item["no.GST"].set(True)
                #     item["Fee"].set(str(value["sub_total"]))
                # elif value["line_amount_types"] == 'Exclusive':
                #     item["no.GST"].set(False)
                #     item["Fee"].set(str(value["sub_total"]))
                # elif value["line_amount_types"] == 'Inclusive':
                #     item["no.GST"].set(False)
                #     item["Fee"].set(str(value["sub_total"]))
            elif bill_number in inv_list["Bills"]["AUTHORISED"].keys():
                item["State"].set("Awaiting Payment")
            elif bill_number in inv_list["Bills"]["PAID"].keys():
                item["State"].set("Paid")
            elif bill_number in inv_list["Bills"]["VOIDED"].keys():
                item["State"].set("Voided")

    with open(invoices_dir, "w") as f:
        json.dump(invoices_json, f, indent=4)

    with open(bills_dir, "w") as f:
        json.dump(bills_json, f, indent=4)



def send_email_with_attachment(filename):
    msg = MIMEMultipart()
    msg['From'] = conf["bridge_email"]
    msg['To'] = conf["xero_bill_email"]
    msg['Subject'] = f"Bill Upload By Bridge"

    part = MIMEBase("application", "octet-stream")
    with open(filename, "rb") as file:
        part.set_payload(file.read())
    encoders.encode_base64(part)
    attach_msg =f"attachment; filename={os.path.basename(filename)}"
    part.add_header("Content-Disposition", attach_msg)
    msg.attach(part)
    smtp = smtplib.SMTP(conf["smap_server"], conf["smap_port"])
    smtp.starttls()
    smtp.login(conf["email_username"], conf["email_password"])
    smtp.sendmail(conf["bridge_email"], conf["xero_bill_email"], msg.as_string())
    smtp.quit()
    print("Email sent to xero")