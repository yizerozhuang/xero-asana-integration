import textwrap
from tkinter import messagebox
import tkinter as tk

from app_dialog import FileSelectDialog

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

def isfloat(string):
    try:
        float(string)
        return True
    except ValueError:
        return False
    except TypeError:
        return False

def reset(app):
    data = app.data
    # if len(app.current_quotation.get())

    if len(app.current_quotation.get()) != 0:
        data["Project Info"]["Project"]["Quotation Number"].set(app.current_quotation.get())
        save(app)
        app.current_quotation.set("")

    database_dir = os.path.join(app.conf["database_dir"], "data_template.json")
    template_json = json.load(open(database_dir))
    template_json["Fee Proposal"]["Reference"]["Date"] = datetime.today().strftime("%d-%b-%Y")

    convert_to_data(template_json, app.data)

    app.fee_proposal_page._reset_scope()

    for service in app.data["Project Info"]["Project"]["Service Type"].values():
        if service["Include"].get():
            service["Include"].set(False)
            service["Include"].set(True)



    app.log_text.set("")

    # app.email_text.delete(1.0, tk.END)

def save(app):
    data = app.data
    data_json = convert_to_json(data)
    database_dir = os.path.join(app.conf["database_dir"], data_json["Project Info"]["Project"]["Quotation Number"])
    # if len(data["Project Info"]["Project"]["Quotation Number"].get()) != 6 or len(data["Project Info"]["Project"]["Quotation Number"].get()) != 8:
    #     return
    if not os.path.exists(database_dir):
        os.mkdir(database_dir)
    with open(os.path.join(database_dir, "data.json"), "w") as f:
        json_object = json.dumps(data_json, indent=4)
        f.write(json_object)
    save_invoice_state(app)
    # current_folder_name = [folder for folder in os.listdir(app.conf["working_dir"]) if folder.startswith(data["Project Info"]["Project"]["Quotation Number"])][0]
    # print(current_folder_name)

def load_data(app):
    data = app.data
    if len(app.current_quotation.get()) != 0:
        old_quotation = data["Project Info"]["Project"]["Quotation Number"].get()
        data["Project Info"]["Project"]["Quotation Number"].set(app.current_quotation.get())
        save(app)
        data["Project Info"]["Project"]["Quotation Number"].set(old_quotation)
    load(app)
    # app.email_text.delete(1.0, tk.END)
    # app.email_text.insert(1.0, app.data["Email_Content"].get())

def load(app):
    data = app.data
    quotation_number = data["Project Info"]["Project"]["Quotation Number"].get().upper()
    database_dir = os.path.join(app.conf["database_dir"], quotation_number)
    data_json = json.load(open(os.path.join(database_dir, "data.json")))
    convert_to_data(data_json, data)


    for service in app.data["Project Info"]["Project"]["Service Type"].values():
        if service["Include"].get():
            service["Include"].set(False)
            service["Include"].set(True)

    app.current_quotation.set(quotation_number)
    unlock_invoice(app)
    load_invoice_state(app)
    config_state(app)
    config_log(app)

def unlock_invoice(app):
    for service, value in app.financial_panel_page.invoice_dic.items():
        for inv in value["Invoice"]:
            inv.config(state=tk.NORMAL)
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
            if len(item["Number"].get()) !=0:
                item["State"].set(bill_json[data["Project Info"]["Project"]["Project Number"].get()+item["Number"].get()])

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
            for i in range(len(json)):
                try:
                    convert_to_data(json[i], data[i])
                except IndexError:
                    data.append(
                        {
                            "Include": tk.BooleanVar(value=True),
                            "Item": tk.StringVar()
                        }
                    )
                    convert_to_data(json[i], data[i])
        elif len(json) < len(data):
            data = data[0: len(json)]
            for i in range(len(data)):
                convert_to_data(json[i], data[i])

    elif isinstance(json, dict):
        for key in json.keys():
            convert_to_data(json[key], data[key])
    else:
        if data.get() != json:
            try:
                data.set(json)
            except Exception as e:
                print(e)


def finish_setup(app):
    data = app.data
    folder_name = data["Project Info"]["Project"]["Quotation Number"].get() + "-" + data["Project Info"]["Project"]["Project Name"].get().strip()
    working_dir = os.path.join(app.conf["working_dir"], folder_name)

    if not os.path.exists(working_dir):
        create_folder = messagebox.askyesno("Folder not found",
                                            f"Can not find the folder {folder_name}, do you want to create the folder")
        if create_folder:
            create_new_folder(folder_name, app.conf)
        else:
            return
    data["State"]["Set Up"].set(True)
    save(app)
    app.log.log_finish_set_up(app)
    config_state(app)
    config_log(app)
    messagebox.showinfo("Set Up",
                        f"Project {data['Project Info']['Project']['Quotation Number'].get()} set up successful")

def config_state(app):
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
    folder_name = data["Project Info"]["Project"]["Quotation Number"].get() + "-" + data["Project Info"]["Project"][
        "Project Name"].get()
    working_dir = os.path.join(app.conf["working_dir"], folder_name)
    database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    recycle_folder = data["Project Info"]["Project"]["Quotation Number"].get() + "-" + datetime.now().strftime(
        "%y%m%d%H%M%S")
    recycle_bin_dir = os.path.join(app.conf["recycle_bin_dir"], recycle_folder)

    os.mkdir(recycle_bin_dir)
    if os.path.exists(database_dir):
        app.log.log_delete(app)
        shutil.move(database_dir, recycle_bin_dir)
    if os.path.exists(working_dir):
        shutil.move(working_dir, recycle_bin_dir)
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
        quotation_letter = current_quotation[6:8][0] + chr(ord(current_quotation[6:8][1]) + 1) if \
            current_quotation[6:8][1] != "Z" else chr(ord(current_quotation[6:8][0]) + 1) + "A"
        current_quotation = current_quotation[:6] + quotation_letter
    return current_quotation

def remove_none(obj):
    if isinstance(obj, list):
        return [remove_none(x) for x in obj if x is not None]
    elif isinstance(obj, dict):
        return {k: remove_none(v) for k, v in obj.items() if v is not None}
    else:
        return obj

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
            quotation_number, project_name = folder.split("-", 1)
            if quotation_number == current_quotation:
                old_dir = os.path.join(working_dir, folder)
                new_dir = os.path.join(working_dir, current_quotation+"-"+data["Project Info"]["Project"]["Project Name"].get())
                os.rename(old_dir, new_dir)
                return old_dir, new_dir



def change_quotation_number(app, new_quotation_number):
    working_dir = app.conf["working_dir"]
    # database_dir = app.conf["database_dir"]

    old_folder_name = app.data["Project Info"]["Project"]["Quotation Number"].get() + "-" + app.data["Project Info"]["Project"]["Project Name"].get()
    old_folder = os.path.join(working_dir, old_folder_name)

    if not os.path.exists(old_folder):
        messagebox.showerror("Error", f"Can not found the folder {old_folder_name}")
        return "", "", False

    new_folder = os.path.join(working_dir, new_quotation_number + "-" + app.data["Project Info"]["Project"]["Project Name"].get())

    try:
        # old_database = os.path.join(database_dir, app.data["Project Info"]["Project"]["Quotation Number"].get())
        # new_database = os.path.join(database_dir, new_quotation_number)
        # os.rename(old_database, new_database)
        os.rename(old_folder, new_folder)
        return old_folder, new_folder, True
    except PermissionError:
        return old_folder, new_folder, False

def rename_new_folder(app):
    data = app.data
    folder_name = data["Project Info"]["Project"]["Quotation Number"].get() + "-" + \
                  data["Project Info"]["Project"]["Project Name"].get()
    folder_path = os.path.join(app.conf["working_dir"], folder_name)
    dir_list = os.listdir(app.conf["working_dir"])
    rename_list = [dir for dir in dir_list if dir.startswith("New folder")]

    if len(data["Project Info"]["Project"]["Project Name"].get()) == 0:
        messagebox.showerror(title="Error", message="Please Input a project name")
        return
    elif len(data["Project Info"]["Project"]["Quotation Number"].get()) == 0:
        messagebox.showerror(title="Error", message="Please Input a quotation number")
        return
    elif len(rename_list) == 0:
        messagebox.showerror(title="Error", message="Please create a new folder first")
        return

    try:
        if len(rename_list) == 1:
            mode = 0o666
            os.mkdir(folder_path, mode)
            shutil.move(os.path.join(app.conf["working_dir"], rename_list[0]), os.path.join(folder_path, "External"))
            os.mkdir(os.path.join(folder_path, "Photos"), mode)
            os.mkdir(os.path.join(folder_path, "Plot"), mode)
            os.mkdir(os.path.join(folder_path, "SS"), mode)
            shutil.copyfile(os.path.join(app.conf["resource_dir"], "xlsx", "Preliminary Calculation v2.5.xlsx"),
                            os.path.join(folder_path, "Preliminary Calculation v2.5.xlsx"))
            messagebox.showinfo(title="Folder renamed", message=f"Rename Folder {rename_list[0]} to {folder_name}")
        else:
            FileSelectDialog(app, rename_list, "Multiple new folders found, please select one")
    except FileExistsError:
        messagebox.showerror("Error", f"Cannot create a file when that file already exists:{folder_name}")
        return
    save(app)
    config_state(app)
    app.log.log_rename_folder(app)
    config_log(app)

def create_new_folder(folder_name, conf):
    folder_path = os.path.join(conf["working_dir"], folder_name)
    os.mkdir(folder_path)
    os.mkdir(os.path.join(folder_path, "External"))
    os.mkdir(os.path.join(folder_path, "Photos"))
    os.mkdir(os.path.join(folder_path, "Plot"))
    os.mkdir(os.path.join(folder_path, "SS"))
    shutil.copyfile(os.path.join(conf["resource_dir"], "xlsx", "Preliminary Calculation v2.5.xlsx"),
                    os.path.join(folder_path, "Preliminary Calculation v2.5.xlsx"))

def _check_fee(app):
    data = app.data
    for service_fee in data["Invoices"]["Details"].values():
        if len(service_fee["Fee"].get()) == 0 and service_fee["Service"].get() != "Variation" and service_fee["Include"].get():
            return False
    return True

def preview_fee_proposal(app, *args):
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
    pdf_name = _get_proposal_name(app)
    excel_name = f'Mechanical_Fee_Proposal for {data["Project Info"]["Project"]["Project Name"].get()} Back Up.xlsx'
    database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    resource_dir = app.conf["resource_dir"]
    adobe_address = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"

    past_projects_dir = os.path.join(app.conf["database_dir"], "past_projects.json")
    services = []
    for service in app.conf["proposal_list"]:
        if data["Project Info"]["Project"]["Service Type"][service]["Include"].get():
            services.append(service)
    # services = [key for key, value in data["Project Info"]["Project"]["Service Type"].items() if value['Include'].get()]
    # win32com.client.makepy.GenerateFromTypeLibSpec('Acrobat')
    # adobe = win32com.client.DispatchEx('AcroExch.App')
    # avDoc = win32client.dynamic.Dispatch("AcroExch.AVDoc")
    page = len(services) // 2 + 1
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
    elif not page in [1, 2, 3]:
        messagebox.showerror("Error", "Excess the maximum value of service, please contact administrator")
    pdf_list = [file for file in os.listdir(os.path.join(database_dir)) if
                str(file).startswith("Mechanical Fee Proposal") and str(file).endswith(".pdf")]

    if len(pdf_list) != 0:
        current_revision = str(max([str(pdf).split(" ")[-1].split(".")[0] for pdf in pdf_list]))
        if data["Fee Proposal"]["Reference"]["Revision"].get() == current_revision or data["Fee Proposal"]["Reference"]["Revision"].get() == str(int(current_revision) + 1):
            old_pdf_path = os.path.join(database_dir, f'Mechanical Fee Proposal for {data["Project Info"]["Project"]["Project Name"].get()} Rev {current_revision}.pdf')
            # avDoc.open(old_pdf_path, old_pdf_path)
            def open_pdf():
                subprocess.call([adobe_address, old_pdf_path])
            _thread.start_new_thread(open_pdf, ())

            overwrite = messagebox.askyesno(f"Warming", f"Revision {current_revision} found, do you want to overwrite")
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
                os.path.join(database_dir, excel_name))
    excel = win32client.Dispatch("Excel.Application")
    try:
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
    except Exception as e:
        print(e)
        pass
    work_book = excel.Workbooks.Open(os.path.join(database_dir, excel_name))
        # messagebox.showerror("Error", e)
    try:
        work_sheets = work_book.Worksheets[0]
        address_to = data["Address_to"].get()
        work_sheets.Cells(1, 2).Value = data["Project Info"][address_to]["Full Name"].get()
        first_name = get_first_name(data["Project Info"][address_to]["Full Name"].get())
        work_sheets.Cells(5, 1).Value = f'Dear {first_name},'
        work_sheets.Cells(2, 2).Value = data["Project Info"][address_to]["Company"].get()
        work_sheets.Cells(1, 8).Value = data["Project Info"]["Project"]["Quotation Number"].get()
        work_sheets.Cells(2, 8).Value = data["Fee Proposal"]["Reference"]["Date"].get()
        work_sheets.Cells(3, 8).Value = data["Fee Proposal"]["Reference"]["Revision"].get()
        work_sheets.Cells(6, 1).Value = "Re: " + data["Project Info"]["Project"]["Project Name"].get()
        work_sheets.Cells(8, 1).Value = f"Thank you for giving us the opportunity to submit this fee proposal for our {', '.join([key for key, value in data['Project Info']['Project']['Service Type'].items() if value['Include'].get()])} for the above project."
        work_sheets.Cells(16, 7).Value = data["Fee Proposal"]["Time"]["Fee Proposal"]["Start"].get() + "-" + \
                                         data["Fee Proposal"]["Time"]["Fee Proposal"]["End"].get()
        work_sheets.Cells(21, 7).Value = data["Fee Proposal"]["Time"]["Pre-design"]["Start"].get() + "-" + \
                                         data["Fee Proposal"]["Time"]["Pre-design"]["End"].get()
        work_sheets.Cells(26, 7).Value = data["Fee Proposal"]["Time"]["Documentation"]["Start"].get() + "-" + \
                                         data["Fee Proposal"]["Time"]["Documentation"]["End"].get()
        cur_row = 52
        cur_index = 1
        i = 0
        for key in app.conf["proposal_list"]:
            if not data["Project Info"]["Project"]["Service Type"][key]["Include"].get():
                continue

            service = data["Fee Proposal"]["Scope"]["Minor"][key]

            cur_row = cur_row if i % 2 == 0 else 84 + (i - 1) // 2 * row_per_page
            extra_list = conf["extra_list"]
            for extra in extra_list:
                work_sheets.Cells(cur_row, 1).Value = "2." + str(cur_index)
                work_sheets.Cells(cur_row, 2).Value = key + "-" + extra
                work_sheets.Cells(cur_row, 1).Font.Bold = True
                work_sheets.Cells(cur_row, 2).Font.Bold = True
                for scope in service[extra]:
                    if scope["Include"].get():
                        cur_row += 1
                        work_sheets.Cells(cur_row, 1).Value = "•"
                        work_sheets.Cells(cur_row, 2).Value = scope["Item"].get()
                cur_row += 2
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

        if page == 3:
            work_sheets.Cells(cur_row + 4, 6).Value = str(total_fee)
            work_sheets.Cells(cur_row + 4, 7).Value = str(total_ist)
        else:
            work_sheets.Cells(cur_row + 3, 6).Value = str(total_fee)
            work_sheets.Cells(cur_row + 3, 7).Value = str(total_ist)
        cur_row += 17
        work_sheets.Cells(cur_row, 2).Value = "Re: " + data["Project Info"]["Project"]["Project Name"].get()
        work_sheets.ExportAsFixedFormat(0, os.path.join(database_dir, pdf_name))
        work_book.Close(True)
    except PermissionError:
        messagebox.showerror("Error", "Please close the preview or file before you use it")
    except FileNotFoundError:
        messagebox.showerror("Error", "Please Contact Administer, the app cant find the file")
    except Exception as e:
        messagebox.showerror("Error", "Please Close the pdf before you generate a new one")
        print(e)
    else:
        app.data["State"]["Generate Proposal"].set(True)
        # avDoc.open(os.path.join(database_dir, pdf_name), os.path.join(database_dir, pdf_name))
        def open_pdf():
            subprocess.call([adobe_address, os.path.join(database_dir, pdf_name)])
        _thread.start_new_thread(open_pdf, ())
        # webbrowser.open(os.path.join(database_dir, pdf_name))
        save(app)
        config_state(app)
        app.log.log_fee_proposal(app)
        config_log(app)
    try:
        excel.ScreenUpdating = True
        excel.DisplayAlerts = True
        excel.EnableEvents = True
        work_book.Close(True)
        # adobe.close(0)
    except Exception as e:
        print(e)
        pass

def major_project_fee_proposal(app, *args):
    data = app.data
    pdf_name = _get_proposal_name(app)
    excel_name = f'Mechanical_Fee_Proposal for {data["Project Info"]["Project"]["Project Name"].get()} Back Up.xlsx'
    database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    resource_dir = app.conf["resource_dir"]
    adobe_address = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"

    past_projects_dir = os.path.join(app.conf["database_dir"], "past_projects.json")
    services = []
    for service in app.conf["proposal_list"]:
        if data["Project Info"]["Project"]["Service Type"][service]["Include"].get():
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
    pdf_list = [file for file in os.listdir(os.path.join(database_dir)) if
                str(file).startswith("Mechanical Fee Proposal") and str(file).endswith(".pdf")]

    case = _check_case(app)
    if case == "Error":
        app.messagebox.show_error("Unsupported Version")
        return

    if len(pdf_list) != 0:
        current_revision = str(max([str(pdf).split(" ")[-1].split(".")[0] for pdf in pdf_list]))
        if data["Fee Proposal"]["Reference"]["Revision"].get() == current_revision or data["Fee Proposal"]["Reference"][
            "Revision"].get() == str(int(current_revision) + 1):
            old_pdf_path = os.path.join(database_dir,
                                        f'Mechanical Fee Proposal for {data["Project Info"]["Project"]["Project Name"].get()} Rev {current_revision}.pdf')
            def open_pdf():
                subprocess.call([adobe_address, old_pdf_path])

            _thread.start_new_thread(open_pdf, ())

            overwrite = messagebox.askyesno(f"Warming", f"Revision {current_revision} found, do you want to overwrite")
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

    shutil.copy(os.path.join(resource_dir, "xlsx", f"major_proposal_template_{page}_{case}.xlsx"),
                os.path.join(database_dir, excel_name))
    excel = win32client.Dispatch("Excel.Application")
    try:
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
    except Exception as e:
        print(e)
        pass
    work_book = excel.Workbooks.Open(os.path.join(database_dir, excel_name))
    # messagebox.showerror("Error", e)
    try:
        work_sheets = work_book.Worksheets[0]
        address_to = data["Address_to"].get()
        work_sheets.Cells(27, 2).Value = data["Project Info"]["Project"]["Project Name"].get()
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
        cur_row = 99
        # for building in data["Project Info"]["Building Features"]["Details"]:
        #     if len(building["Space"].get())!=0:
        #         work_sheets.Cells(cur_row, 2).Value = building["Space"].get()
        #         cur_row+=1
        if not data["Project Info"]["Building Features"]["Notes"].get() is None:
            for line in data["Project Info"]["Building Features"]["Notes"].get().split("\n"):
                work_sheets.Cells(cur_row, 2).Value = line
                cur_row+=1

        cur_index = 3
        i = 0
        for key in app.conf["proposal_list"]:
            if not data["Project Info"]["Project"]["Service Type"][key]["Include"].get():
                continue
            service = data["Fee Proposal"]["Scope"]["Major"][key]

            cur_row = 131 + i * row_per_page
            extra_list = conf["extra_list"]
            for extra in extra_list:
                if len(service[extra]) == 0:
                    continue
                work_sheets.Cells(cur_row, 1).Value = "2." + str(cur_index)
                work_sheets.Cells(cur_row, 2).Value = key + "-" + extra
                work_sheets.Cells(cur_row, 1).Font.Bold = True
                work_sheets.Cells(cur_row, 2).Font.Bold = True
                cur_row+=1
                for scope in service[extra]:
                    if scope["Include"].get():
                        if scope["Item"].get() in conf["sub_title"]:
                            work_sheets.Cells(cur_row, 2).Value = scope["Item"].get()
                            work_sheets.Cells(cur_row, 2).Font.Bold = True
                            cur_row+=1
                        else:
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

        cur_row = 182 + page * row_per_page
        stage_total = [0]*4
        total = 0
        gst = 0
        total_gst = 0

        for service in conf["proposal_list"]:
            if not data["Project Info"]["Project"]["Service Type"][service]["Include"].get():
                continue
            total += float(data["Invoices"]["Details"][service]["Fee"].get())
            gst += float(data["Invoices"]["Details"][service]["in.GST"].get()) - float(data["Invoices"]["Details"][service]["Fee"].get())
            total_gst += float(data["Invoices"]["Details"][service]["in.GST"].get())

            work_sheets.Cells(cur_row, 2).Value = service
            if data["Invoices"]["Details"][service]["Expand"].get():
                j = 0
                for item in data["Invoices"]["Details"][service]["Content"]:
                    item_fee = float(item["Fee"].get()) if len(item["Fee"].get()) !=0 else 0
                    work_sheets.Cells(cur_row, 4+j).Value = item_fee
                    stage_total[j] += item_fee
                    j+=1
            work_sheets.Cells(cur_row, 8).Value = float(data["Invoices"]["Details"][service]["Fee"].get())
            cur_row +=1
        cur_row = 182 + (row_per_page+1)*page
        work_sheets.Cells(cur_row, 4).Value = stage_total[0]
        work_sheets.Cells(cur_row, 5).Value = stage_total[1]
        work_sheets.Cells(cur_row, 6).Value = stage_total[2]
        work_sheets.Cells(cur_row, 7).Value = stage_total[3]



        work_sheets.Cells(cur_row, 8).Value = total
        work_sheets.Cells(cur_row+1, 8).Value = gst
        work_sheets.Cells(cur_row+2, 8).Value = total_gst

        cur_row = 200 + page * (row_per_page+1)
        work_sheets.Cells(cur_row, 2).Value = "Re: "+data["Project Info"]["Project"]["Project Name"].get()
        work_sheets.ExportAsFixedFormat(0, os.path.join(database_dir, pdf_name))
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
            subprocess.call([adobe_address, os.path.join(database_dir, pdf_name)])

        _thread.start_new_thread(open_pdf, ())
        # webbrowser.open(os.path.join(database_dir, pdf_name))
        save(app)
        config_state(app)
        app.log.log_fee_proposal(app)
        config_log(app)
    try:
        excel.ScreenUpdating = True
        excel.DisplayAlerts = True
        excel.EnableEvents = True
        work_book.Close(True)
        # adobe.close(0)
    except Exception as e:
        print(e)
        pass
def _get_proposal_name(app):
    data = app.data
    return f'Mechanical Fee Proposal for {data["Project Info"]["Project"]["Project Name"].get()} Rev {data["Fee Proposal"]["Reference"]["Revision"].get()}.pdf'

def _check_case(app):
    data = app.data
    stages = app.conf["major_stage"]
    stage_1 = data["Fee Proposal"]["Stage"][stages[0]].get()
    stage_2 = data["Fee Proposal"]["Stage"][stages[1]].get()
    stage_3 = data["Fee Proposal"]["Stage"][stages[2]].get()
    stage_4 = data["Fee Proposal"]["Stage"][stages[3]].get()

    if stage_1 and stage_2 and stage_3 and stage_4:
        return "case1"
    elif stage_1 and stage_2 and stage_3 and not stage_4:
        return "case2"
    elif stage_1 and not stage_2 and not stage_3 and not stage_4:
        return "case3"
    elif not stage_1 and stage_2 and stage_3 and stage_4:
        return "case4"
    elif not stage_1 and stage_2 and stage_3 and not stage_4:
        return "case5"
    else:
        return "Error"
def excel_print_invoice(app, inv):
    data = app.data
    excel_name = f'PCE INV {data["Invoices Number"][inv]["Number"].get()}.xlsx'
    invoice_name = f'PCE INV {data["Invoices Number"][inv]["Number"].get()}.pdf'
    database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    resource_dir = app.conf["resource_dir"]
    adobe_address = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
    # win32com.client.makepy.GenerateFromTypeLibSpec('Acrobat')
    # adobe = win32com.client.DispatchEx('AcroExch.App')
    # avDoc = win32client.dynamic.Dispatch("AcroExch.AVDoc")

    if len(data["Invoices Number"][inv]["Number"].get()) == 0:
        messagebox.showerror("Error", "You need to generate a invoice number before you generate the Invoice")
        return

    rewrite = True
    if os.path.exists(os.path.join(database_dir, invoice_name)):
        old_pdf_path=os.path.join(database_dir, invoice_name)
        # avDoc.open(old_pdf_path, old_pdf_path)
        def open_pdf():
            subprocess.call([adobe_address, old_pdf_path])
        _thread.start_new_thread(open_pdf, ())
        # webbrowser.open(old_pdf_path)
        rewrite = messagebox.askyesno("Warming", f"Existing file PCE INV {invoice_name}.pdf do you want to rewrite?")
        if not rewrite:
            return
        else:
            #     avDoc.close(0)
            for proc in psutil.process_iter():
                if proc.name() == "Acrobat.exe":
                    proc.kill()
    if rewrite:
        shutil.copy(os.path.join(resource_dir, "xlsx", f"invoice_template.xlsx"),
                    os.path.join(database_dir, excel_name))
        excel = win32client.Dispatch("Excel.Application")
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        work_book = excel.Workbooks.Open(os.path.join(database_dir, excel_name))
        try:
            work_sheets = work_book.Worksheets[0]
            address_to = data["Address_to"].get()
            work_sheets.Cells(4, 1).Value = data["Project Info"][address_to]["Full Name"].get()
            work_sheets.Cells(6, 1).Value = data["Project Info"][address_to]["Company"].get()
            work_sheets.Cells(7, 1).Value = data["Project Info"][address_to]["Address"].get()
            work_sheets.Cells(4, 10).Value = datetime.today().strftime("%d-%b-%Y")
            work_sheets.Cells(5, 10).Value = data["Invoices Number"][inv]["Number"].get()
            work_sheets.Cells(12, 1).Value = data["Project Info"]["Project"]["Project Name"].get()
            cur_row = 14
            total_fee = 0
            total_inGST = 0
            for service in data["Invoices"]["Details"].values():
                if not service["Include"].get() or service["Service"].get()=="Variation":
                    continue
                if service["Expand"].get():
                    work_sheets.Cells(cur_row, 1).Value = service["Service"].get()
                    work_sheets.Cells(cur_row+1, 2).Value = "Payment Instalments"
                    work_sheets.Cells(cur_row+1, 7).Value = "Amount"
                    work_sheets.Cells(cur_row+1, 8).Value = "This pay"
                    cur_row += 2
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
                else:
                    work_sheets.Cells(cur_row, 1).Value = service["Service"].get() + " design and documentation"
                    work_sheets.Cells(cur_row+1, 2).Value = "Payment Instalments"
                    work_sheets.Cells(cur_row+1, 7).Value = "Amount"
                    work_sheets.Cells(cur_row+1, 8).Value = "This pay"
                    work_sheets.Cells(cur_row+2, 1).Value = "The Full amount for design and documentation"
                    work_sheets.Cells(cur_row+2, 7).Value = service["Fee"].get()
                    if service["Number"].get() == f"INV{str(inv+1)}":
                        work_sheets.Cells(cur_row+2, 8).Value = service["Fee"].get()
                        work_sheets.Cells(cur_row+2, 9).Value = service["Fee"].get()
                        work_sheets.Cells(cur_row+2, 10).Value = service["in.GST"].get()
                        total_fee += float(service["Fee"].get()) if len(service["Fee"].get()) !=0 else 0
                        total_inGST += float(service["in.GST"].get()) if len(service["Fee"].get()) !=0 else 0
                    cur_row+=3
                # cur_row+=1
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

            work_sheets.Cells(33, 9).Value = str(total_fee)
            work_sheets.Cells(33, 10).Value = str(total_inGST)
            work_sheets.ExportAsFixedFormat(0, os.path.join(database_dir, invoice_name))
            work_book.Close(True)
        except PermissionError:
            messagebox.showerror("Error", "Please close the preview or file before you use it")
        except Exception as e:
            messagebox.showerror("Error", "Please Close the pdf before you generate a new one")
            print(e)
        else:
            invoice_path = os.path.join(database_dir, invoice_name)

            def open_pdf():
                subprocess.call([adobe_address, invoice_path])

            _thread.start_new_thread(open_pdf, ())
            # webbrowser.open(invoice_path)
            # avDoc.open(invoice_path, invoice_path)
            save(app)
            app.log.log_generate_invoices(app, inv)
            config_state(app)
            config_log(app)
        try:
            excel.ScreenUpdating = True
            excel.DisplayAlerts = True
            excel.EnableEvents = True
            work_book.Close(True)
        except:
            pass

def email_fee_proposal(app, *args):
    data = app.data
    database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    pdf_name = f'Mechanical Fee Proposal for {data["Project Info"]["Project"]["Project Name"].get()} Rev {data["Fee Proposal"]["Reference"]["Revision"].get()}.pdf'

    if not data["State"]["Generate Proposal"].get():
        messagebox.showerror("Error", "Please Generate a pdf first")
        return
    elif not os.path.exists(os.path.join(database_dir, pdf_name)):
        messagebox.showerror("Error",
                             f'Python cant found fee proposal for {data["Project Info"]["Project"]["Quotation Number"].get()} revision {data["Fee Proposal"]["Reference"]["Revision"].get()}')
        return

    if not "OUTLOOK.EXE" in (p.name() for p in psutil.process_iter()):
        os.startfile("outlook")

    # user_email_dic = json.load(open(os.path.join(app.conf["database_dir"], "user_email.json")))

    ol = win32client.Dispatch("Outlook.Application")
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    # newmail.From = user_email_dic[app.user]
    newmail.Subject = f'{data["Project Info"]["Project"]["Quotation Number"].get()}-Mechanical Fee Proposal - {data["Project Info"]["Project"]["Project Name"].get()} Rev {data["Fee Proposal"]["Reference"]["Revision"].get()}'
    newmail.To = f'{data["Project Info"]["Client"]["Contact Email"].get()}; {data["Project Info"]["Main Contact"]["Contact Email"].get()}'
    newmail.CC = "felix@pcen.com.au; admin@pcen.com.au"
    newmail.BCC = "bridge@pcen.com.au"
    newmail.GetInspector()
    index = newmail.HTMLbody.find(">", newmail.HTMLbody.find("<body"))

    address_to = data["Address_to"].get()
    full_name = data["Project Info"][address_to]["Full Name"].get()
    first_name = get_first_name(full_name)

    message = f"""
    Dear {first_name},<br>

    I hope this email finds you well. Please find the attached fee proposal to this email.<br>

    If you have any questions or need more information regarding the proposal, please don't hesitate to reach out. I am happy to provide you with whatever information you need.<br>

    I look forward to hearing from you soon.<br>

    Cheers,<br>
    """
    newmail.HTMLbody = newmail.HTMLbody[:index + 1] + message + newmail.HTMLbody[index + 1:]
    newmail.Attachments.Add(os.path.join(database_dir, pdf_name))
    newmail.Display()
    save(app)
    config_state(app)
    return True
def chase(app, *args):
    data = app.data
    if not data["State"]["Email to Client"].get():
        messagebox.showerror("Error", "Please Sent the fee proposal to client first")
        return

    if not "OUTLOOK.EXE" in (p.name() for p in psutil.process_iter()):
        os.startfile("outlook")
    ol = win32client.Dispatch("Outlook.Application")
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = f'Re: {data["Project Info"]["Project"]["Quotation Number"].get()}-{data["Project Info"]["Project"]["Project Name"].get()}'
    newmail.To = f'{data["Project Info"]["Client"]["Contact Email"].get()}; {data["Project Info"]["Main Contact"]["Contact Email"].get()}'
    newmail.CC = "felix@pcen.com.au; admin@pcen.com.au"
    newmail.BCC = "bridge@pcen.com.au"
    newmail.GetInspector()
    index = newmail.HTMLbody.find(">", newmail.HTMLbody.find("<body"))

    address_to = data["Address_to"].get()
    full_name = data["Project Info"][address_to]["Full Name"].get()
    first_name = get_first_name(full_name)

    message = f"""
    Hi {first_name},<br>
    I wanted to circle back regarding the fee proposal we sent on {data["Email"]["Fee Proposal"].get()}. Do you have any questions or concerns? We're looking forward to working with you and hearing your feedback. <br>

    Thank you for considering our proposal, and we anticipate your response soon.<br>

    Cheers ,<br>
    """
    newmail.HTMLbody = newmail.HTMLbody[:index + 1] + message + newmail.HTMLbody[index + 1:]
    newmail.Display()
    save(app)
    config_state(app)
def email_invoice(app, inv):
    data = app.data
    database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    pdf_name = f'PCE INV {data["Invoices Number"][inv]["Number"].get()}.pdf'

    if not os.path.exists(os.path.join(database_dir, pdf_name)):
        messagebox.showerror("Error", f"Please generate the invoice before you send to client")
        return
    # if not data["State"]["Fee Accepted"].get():
    #     messagebox.showerror("Error", "Please Sent the fee proposal to client first")
    #     return


    if not "OUTLOOK.EXE" in (p.name() for p in psutil.process_iter()):
        os.startfile("outlook")
    ol = win32client.Dispatch("Outlook.Application")
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = f'Invoice {data["Invoices Number"][inv]["Number"].get()}-{data["Project Info"]["Project"]["Project Name"].get()}'
    newmail.To = f'{data["Project Info"]["Client"]["Contact Email"].get()}; {data["Project Info"]["Main Contact"]["Contact Email"].get()}'
    newmail.CC = "felix@pcen.com.au; admin@pcen.com.au"
    newmail.BCC = "bridge@pcen.com.au"
    newmail.GetInspector()
    index = newmail.HTMLbody.find(">", newmail.HTMLbody.find("<body"))

    address_to = data["Address_to"].get()
    full_name = data["Project Info"][address_to]["Full Name"].get()
    first_name = get_first_name(full_name)

    message = f"""
    Hi {first_name},<br>
    I hope this email finds you well and energized for the week ahead.<br>

    Please find the attached invoice for client payment. We appreciate your prompt attention to this matter.<br>

    If you have any questions or concerns regarding the invoice, please do not hesitate to contact Felix and me. We are happy to discuss further.<br>

    Thank you for your time and consideration.<br>
    
    Cheers ,<br>
    """
    newmail.HTMLbody = newmail.HTMLbody[:index + 1] + message + newmail.HTMLbody[index + 1:]
    newmail.Display()
    newmail.Attachments.Add(os.path.join(database_dir, pdf_name))

def get_invoice_item(app):
    data = app.data
    res = [[] for _ in range(6)]
    for i in range(6):
        for key, service in data["Invoices"]["Details"].items():
            if not service["Include"].get():
                continue
            if service["Expand"].get():
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
            else:
                if service["Number"].get() == f"INV{i+1}":
                    if len(service["Fee"].get()) != 0:
                        res[i].append(
                            {
                                "Item": service["Service"].get(),
                                "Fee": service["Fee"].get(),
                                "in.GST": service["in.GST"].get()
                            }
                        )
    return res

def update_app_invoices(app, inv_list):
    data = app.data
    invoices_dir = os.path.join(app.conf["database_dir"], "invoices.json")
    invoices_json = json.load(open(invoices_dir))

    bills_dir = os.path.join(app.conf["database_dir"], "bills.json")
    bills_json = json.load(open(bills_dir))

    for state, inv_dict in inv_list["Invoices"].items():
        for inv_number, value in inv_dict.items():
            if state == "PAID":
                invoices_json[inv_number] = "Paid"
            elif state == "VOIDED":
                invoices_json[inv_number] = "Void"

    for state, bill_dict in inv_list["Bills"].items():
        for bill_number, value in bill_dict.items():
            if state == "SUBMITTED":
                bills_json[bill_number] = "Awaiting approval"
            elif state == "AUTHORISED":
                bills_json[bill_number] = "Awaiting payment"
            elif state == "PAID":
                bills_json[bill_number] = "Paid"
            elif state == "VOIDED":
                bills_json[bill_number] = "Void"

    for value in data["Invoices Number"]:
        if value["Number"].get() in inv_list["Invoices"]["PAID"].keys():
            value["State"].set("Paid")
        elif value["Number"].get() in inv_list["Invoices"]["VOIDED"].keys():
            value["State"].set("Void")

    for value in data["Bills"]["Details"].values():
        for item in value["Content"]:
            bill_number = data["Project Info"]["Project"]["Project Number"].get() + item["Number"].get()
            if bill_number in inv_list["Bills"]["SUBMITTED"].keys():
                item["State"].set("Awaiting approval")
                value = inv_list["Bills"]["SUBMITTED"][bill_number]
                if value["line_amount_types"] == 'NoTax':
                    item["no.GST"].set(True)
                    item["Fee"].set(str(value["sub_total"]))
                elif value["line_amount_types"] == 'Exclusive':
                    item["no.GST"].set(False)
                    item["Fee"].set(str(value["sub_total"]))
                elif value["line_amount_types"] == 'Inclusive':
                    item["no.GST"].set(False)
                    item["Fee"].set(str(value["sub_total"]))
            elif bill_number in inv_list["Bills"]["AUTHORISED"].keys():
                item["State"].set("Awaiting payment")
            elif bill_number in inv_list["Bills"]["PAID"].keys():
                item["State"].set("Paid")
            elif bill_number in inv_list["Bills"]["VOIDED"].keys():
                item["State"].set("Void")

    with open(invoices_dir, "w") as f:
        json.dump(invoices_json, f, indent=4)

    with open(bills_dir, "w") as f:
        json.dump(bills_json, f, indent=4)



def send_email_with_attachment(filename):
    msg = MIMEMultipart()
    msg['From'] = conf["bridge_email"]
    msg['To'] = conf["xero_bill_email"]
    msg['Subject'] = f"Bill Number: {filename.split('-')[0]}"

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
