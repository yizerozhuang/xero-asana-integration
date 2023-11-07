from tkinter import messagebox

from custom_dialog import FileSelectDialog

from win32com import client as win32client
import shutil
import os
import webbrowser
import json
from datetime import date, datetime
import subprocess
import psutil


def reset(app):
    database_dir = os.path.join(app.conf["database_dir"], app.data["Project Info"]["Project"]["Quotation Number"].get())
    if os.path.exists(database_dir) and len(app.data["Project Info"]["Project"]["Quotation Number"].get()) != 0:
        save(app)
    database_dir = os.path.join(app.conf["database_dir"], "data_template.json")
    template_json = json.load(open(database_dir))
    template_json["Fee Proposal"]["Reference"]["Date"] = datetime.today().strftime("%d-%b-%Y")
    convert_to_data(template_json, app.data)
    app.log_text.set("")


def save(app):
    data = app.data
    data_json = convert_to_json(data)
    database_dir = os.path.join(app.conf["database_dir"], data_json["Project Info"]["Project"]["Quotation Number"])
    if not os.path.exists(database_dir):
        os.mkdir(database_dir)
    with open(os.path.join(database_dir, "data.json"), "w") as f:
        json_object = json.dumps(data_json, indent=4)
        f.write(json_object)
    # current_folder_name = [folder for folder in os.listdir(app.conf["working_dir"]) if folder.startswith(data["Project Info"]["Project"]["Quotation Number"])][0]
    # print(current_folder_name)


def load(app):
    data = app.data
    database_dir = os.path.join(app.conf["database_dir"],
                                data["Project Info"]["Project"]["Quotation Number"].get().upper())
    data_json = json.load(open(os.path.join(database_dir, "data.json")))
    convert_to_data(data_json, data)
    config_log(app)


def convert_to_json(obj):
    if isinstance(obj, list):
        return [convert_to_json(x) for x in obj]
    elif isinstance(obj, dict):
        return {k: convert_to_json(v) for k, v in obj.items()}
    else:
        return obj.get()


def convert_to_data(json, data):
    if isinstance(json, list):
        # try:
        [convert_to_data(json[i], data[i]) for i in range(len(data))]
        # except IndexError:
        #     data.append(
        #         {
        #             "Include": tk.BooleanVar(value=True),
        #             "Item": tk.StringVar()
        #         }
        #     )
        #     convert_to_data(json[len(data)-1], data[-1])
    elif isinstance(json, dict):
        [convert_to_data(json[k], data[k]) for k in json.keys()]
    else:
        if data.get() != json:
            data.set(json)


def finish_setup(app):
    data = app.data
    folder_name = data["Project Info"]["Project"]["Quotation Number"].get() + "-" + data["Project Info"]["Project"][
        "Project Name"].get()
    working_dir = os.path.join(app.conf["working_dir"], folder_name)

    if not os.path.exists(working_dir):
        create_folder = messagebox.askyesno("Folder not found",
                                            f"Can not find the folder {folder_name}, do you want to create the folder")
        if create_folder:
            create_new_folder(folder_name, app.conf)
        else:
            return
    data["State"]["Set Up"].set(True)
    data["State"]["Generate Proposal"].set(True)
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
    if data["State"]["Done"] or data["State"]["Quote Unsuccessful"]:
        return
    elif data["State"]["Fee Accepted"]:
        _classify_fee(res, data)
    elif data["State"]["Email to Client"]:
        res["Email to Client"].append(data["Project Info"]["Project"]["Quotation Number"])
    elif data["State"]["Generate Proposal"]:
        res["Generate Proposal"].append(data["Project Info"]["Project"]["Quotation Number"])
    elif data["State"]["Set Up"]:
        res["Set Up"].append(
            data["Project Info"]["Project"]["Quotation Number"] + "-" + data["Project Info"]["Project"]["Project Name"])


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
    log = open(log_file).readlines()
    log.reverse()
    app.log_text.set("".join(log))


def get_quotation_number():
    current_quotation_list = [dir for dir in os.listdir("..") if
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
            app.data["State"]["Set Up"].set(True)
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
        if len(service_fee["Fee"].get()) == 0:
            return False
    return True


def excel_print_pdf(app, *args):
    data = app.data
    folder_name = data["Project Info"]["Project"]["Quotation Number"].get() + "-" + \
                  data["Project Info"]["Project"]["Project Name"].get()
    folder_path = os.path.join(app.conf["working_dir"], folder_name)
    pdf_name = f'Mechanical Fee Proposal revision {data["Fee Proposal"]["Reference"]["Revision"].get()}.pdf'
    database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    resource_dir = app.conf["resource_dir"]
    past_projects_dir = os.path.join(app.conf["database_dir"], "past_projects.json")
    services = [key for key in data["Fee Proposal"]["Scope"].keys()]
    page = len(services) // 2 + 1

    if not data["State"]["Generate Proposal"].get():
        messagebox.showerror("Error", "Please finish Set Up first")
        return
    elif not os.path.exists(folder_path):
        messagebox.showerror("Error", "Can't find the folder, please rename first")
        return
    elif len(services) == 0:
        messagebox.showerror("Error", "Please at least select 1 service")
        return
    elif not _check_fee(app):
        messagebox.showerror("Error", "Please go to fee proposal page to complete fee first")
        return
    elif not page in [1, 2, 3]:
        messagebox.showerror("Error", "Excess the maximum value of service, please contact administrator")
    pdf_list = [file for file in os.listdir(os.path.join(database_dir)) if
                str(file).startswith("Mechanical Fee Proposal")]
    if len(pdf_list) != 0:
        current_revision = str(max([str(pdf).split(" ")[-1].split(".")[0] for pdf in pdf_list]))
        if data["Fee Proposal"]["Reference"]["Revision"].get() == current_revision or data["Fee Proposal"]["Reference"][
            "Revision"].get() == str(int(current_revision) + 1):
            msg = messagebox.askyesno(f"Warming", f"Revision {current_revision} found, do you want to overwrite")
            if not msg:
                return
        else:
            messagebox.showerror("Error",
                                 f'Current revision is {current_revision}, you can not use revision {data["Fee Proposal"]["Reference"]["Revision"].get()}')
            return
    else:
        if not data["Fee Proposal"]["Reference"]["Revision"].get() == "1":
            messagebox.showerror("Error", "There is no other existing fee proposal found, can only have revision 1")
            return


    shutil.copy(os.path.join(resource_dir, "xlsx", f"fee_proposal_template_{page}.xlsx"),
                os.path.join(database_dir, "fee proposal.xlsx"))
    excel = win32client.Dispatch("Excel.Application")
    excel.ScreenUpdating = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    work_book = excel.Workbooks.Open(os.path.join(database_dir, "fee proposal.xlsx"))
    try:
        work_sheets = work_book.Worksheets[0]
        work_sheets.Cells(1, 2).Value = data["Project Info"]["Client"]["Client Full Name"].get()
        work_sheets.Cells(5, 2).Value = data["Project Info"]["Client"]["Client Full Name"].get()
        work_sheets.Cells(2, 2).Value = data["Project Info"]["Client"]["Client Company"].get()
        work_sheets.Cells(1, 8).Value = data["Project Info"]["Project"]["Quotation Number"].get()
        work_sheets.Cells(2, 8).Value = data["Fee Proposal"]["Reference"]["Date"].get()
        work_sheets.Cells(3, 8).Value = data["Fee Proposal"]["Reference"]["Revision"].get()
        work_sheets.Cells(6, 1).Value = "Re: " + data["Project Info"]["Project"]["Project Name"].get()
        work_sheets.Cells(8, 1).Value = f"""Thank you for giving us the opportunity to submit this fee proposal for our 
                                        {', '.join([value['Service'].get() for value in data['Project Info']['Project']['Service Type'] if value['Include'].get()])} 
                                        for the above project."""
        work_sheets.Cells(16, 7).Value = data["Fee Proposal"]["Time"]["Fee Proposal"]["Start"].get() + "-" + \
                                         data["Fee Proposal"]["Time"]["Fee Proposal"]["End"].get()
        work_sheets.Cells(21, 7).Value = data["Fee Proposal"]["Time"]["Pre-design"]["Start"].get() + "-" + \
                                         data["Fee Proposal"]["Time"]["Pre-design"]["End"].get()
        work_sheets.Cells(26, 7).Value = data["Fee Proposal"]["Time"]["Documentation"]["Start"].get() + "-" + \
                                         data["Fee Proposal"]["Time"]["Documentation"]["End"].get()
        cur_row = 52
        cur_index = 1
        for i, _ in enumerate(data['Fee Proposal']['Scope'].items()):
            cur_row = cur_row if i % 2 == 0 else 84 + (i - 1) // 2 * 46
            key, service = _
            extra_list = ["Extend", "Exclusion", "Deliverables"]
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
        cur_row = 102 + (page - 1) * 46
        project_type = data["Project Info"]["Project"]["Project Type"].get()
        past_projects = json.load(open(past_projects_dir, encoding="utf-8"))[project_type]
        for i, project in enumerate(past_projects):
            work_sheets.Cells(cur_row + i, 1).Value = "•"
            work_sheets.Cells(cur_row + i, 2).Value = project
        cur_row += 34
        for i, service in enumerate(data["Invoices"]["Details"].values()):
            work_sheets.Cells(cur_row + i, 2).Value = service["Service"].get() + " design and documentation"
            work_sheets.Cells(cur_row + i, 6).Value = service["Fee"].get()
            work_sheets.Cells(cur_row + i, 7).Value = service["in.GST"].get()
        if page == 3:
            work_sheets.Cells(cur_row + 4, 6).Value = data["Invoices"]["Fee"].get()
            work_sheets.Cells(cur_row + 4, 7).Value = data["Invoices"]["in.GST"].get()
        else:
            work_sheets.Cells(cur_row + 3, 6).Value = data["Invoices"]["Fee"].get()
            work_sheets.Cells(cur_row + 3, 7).Value = data["Invoices"]["in.GST"].get()
        cur_row += 17
        work_sheets.Cells(cur_row, 2).Value = "Re: " + data["Project Info"]["Project"]["Project Name"].get()
        work_sheets.ExportAsFixedFormat(0, os.path.join(database_dir, pdf_name))
        work_book.Close(True)
        # shutil.copyfile(
        #     os.path.join(folder_path, "Plot", pdf_name),
        #     os.path.join(database_dir, pdf_name)
        # )
        # shutil.copyfile(
        #     os.path.join(folder_path, "fee proposal.xlsx"),
        #     os.path.join(database_dir, "fee proposal.xlsx")
        # )
        # os.remove(os.path.join(folder_path, "fee proposal.xlsx"))
    except PermissionError:
        messagebox.showerror("Error", "Please close the preview or file before you use it")
    except FileNotFoundError:
        messagebox.showerror("Error", "Please Contact Administer, the app cant find the file")
    except Exception as e:
        messagebox.showerror("Error", "Please Close the pdf before you generate a new one")
        print(e)
    else:
        app.data["State"]["Email to Client"].set(True)
        webbrowser.open(os.path.join(database_dir, pdf_name))
        save(app)
        config_state(app)
        app.log.log_fee_proposal(app)
        config_log(app)
    try:
        excel.ScreenUpdating = True
        excel.DisplayAlerts = True
        excel.EnableEvents = True
        work_book.Close(True)
    except:
        pass

def excel_print_invoice(app, inv):
    inv = f"INV{str(inv+1)}"
    data = app.data
    excel_name = inv + ".xlsx"
    invoice_name = inv + ".pdf"
    database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    resource_dir = app.conf["resource_dir"]
    # if not data["State"]["Done"].get():
    #     messagebox.showerror("Error", "Please Submit a fee acceptance yet")
    #     return
    # invoice_list = [file for file in os.listdir(os.path.join(database_dir)) if
    #             str(file).startswith("Mechanical Fee Proposal")]
    # if len(pdf_list) != 0:
    #     current_revision = str(max([str(pdf).split(" ")[-1].split(".")[0] for pdf in pdf_list]))
    #     if data["Fee Proposal"]["Reference"]["Revision"].get() == current_revision or data["Fee Proposal"]["Reference"][
    #         "Revision"].get() == str(int(current_revision) + 1):
    #         msg = messagebox.askyesno(f"Warming", f"Revision {current_revision} found, do you want to overwrite")
    #         if not msg:
    #             return
    #     else:
    #         messagebox.showerror("Error",
    #                              f'Current revision is {current_revision}, you can not use revision {data["Fee Proposal"]["Reference"]["Revision"].get()}')
    #         return
    # else:
    #     if not data["Fee Proposal"]["Reference"]["Revision"].get() == "1":
    #         messagebox.showerror("Error", "There is no other existing fee proposal found, can only have revision 1")
    #         return
    shutil.copy(os.path.join(resource_dir, "xlsx", f"invoice_template.xlsx"),
                os.path.join(database_dir, excel_name))
    excel = win32client.Dispatch("Excel.Application")
    excel.ScreenUpdating = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    work_book = excel.Workbooks.Open(os.path.join(database_dir, excel_name))
    try:
        work_sheets = work_book.Worksheets[0]
        work_sheets.Cells(4, 1).Value = data["Project Info"]["Client"]["Client Full Name"].get()
        work_sheets.Cells(6, 1).Value = data["Project Info"]["Client"]["Client Company"].get()
        work_sheets.Cells(7, 1).Value = data["Project Info"]["Client"]["Client Address"].get()
        work_sheets.Cells(4, 10).Value = datetime.today().strftime("%d-%b-%Y")
        work_sheets.Cells(5, 10).Value = data["Financial Panel"]["Invoice Details"][inv]["Number"].get()
        work_sheets.Cells(12, 1).Value = data["Project Info"]["Project"]["Project Name"].get()
        cur_row = 14
        total_fee = 0
        total_inGST = 0
        for service in data["Invoices"]["Details"].values():
            if service["Expand"].get():
                work_sheets.Cells(cur_row, 1).Value = service["Service"].get()
                work_sheets.Cells(cur_row + 1, 2).Value = "Payment Instalments"
                work_sheets.Cells(cur_row + 1, 7).Value = "Amount"
                work_sheets.Cells(cur_row + 1, 8).Value = "This pay"
                cur_row += 2
                for item in service["Content"]:
                    if len(item["Service"].get()) != 0:
                        work_sheets.Cells(cur_row, 1).Value = item["Service"].get()
                        work_sheets.Cells(cur_row, 7).Value = item["Fee"].get()
                        if item["Number"].get() == inv:
                            work_sheets.Cells(cur_row, 8).Value = item["Fee"].get()
                            work_sheets.Cells(cur_row, 9).Value = item["Fee"].get()
                            work_sheets.Cells(cur_row, 10).Value = item["in.GST"].get()
                            total_fee += float(item["Fee"].get())
                            total_inGST += float(item["in.GST"].get())
                        cur_row += 1
            else:
                work_sheets.Cells(cur_row, 1).Value = service["Service"].get() + " design and documentation"
                work_sheets.Cells(cur_row+1, 2).Value = "Payment Instalments"
                work_sheets.Cells(cur_row + 1, 7).Value = "Amount"
                work_sheets.Cells(cur_row + 1, 8).Value = "This pay"
                work_sheets.Cells(cur_row+2, 1).Value = "The Full amount for design and documentation"
                work_sheets.Cells(cur_row+2, 7).Value = service["Fee"].get()
                if service["Number"].get() == inv:
                    work_sheets.Cells(cur_row+2, 8).Value = service["Fee"].get()
                    work_sheets.Cells(cur_row+2, 9).Value = service["Fee"].get()
                    work_sheets.Cells(cur_row+2, 10).Value = service["in.GST"].get()
                    total_fee += float(service["Fee"].get())
                    total_inGST += float(service["in.GST"].get())
                cur_row+=3
            cur_row+=1
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
        webbrowser.open(os.path.join(database_dir, invoice_name))
        save(app)
        config_state(app)
        config_log(app)
    try:
        excel.ScreenUpdating = True
        excel.DisplayAlerts = True
        excel.EnableEvents = True
        work_book.Close(True)
    except:
        pass


def email(app, *args):
    data = app.data
    database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    pdf_name = f'Mechanical Fee Proposal revision {data["Fee Proposal"]["Reference"]["Revision"].get()}.pdf'

    if not data["State"]["Email to Client"].get():
        messagebox.showerror("Error", "Please Generate a pdf first")
        return
    elif not os.path.exists(os.path.join(database_dir, pdf_name)):
        messagebox.showerror("Error",
                             f'Python cant found fee proposal for {data["Project Info"]["Project"]["Quotation Number"].get()} revision {data["Fee Proposal"]["Reference"]["Revision"].get()}')
        return

    if not "OUTLOOK.EXE" in (p.name() for p in psutil.process_iter()):
        os.startfile("outlook")
    ol = win32client.Dispatch("Outlook.Application")
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = f'{data["Project Info"]["Project"]["Quotation Number"].get()}-Mechanical Fee Proposal - {data["Project Info"]["Project"]["Project Name"].get()} Rev {data["Fee Proposal"]["Reference"]["Revision"].get()}'
    newmail.To = f'{data["Project Info"]["Client"]["Contact Email"].get()}; {data["Project Info"]["Main Contact"]["Main Contact Email"].get()}'
    newmail.CC = "felix@pcen.com.au"
    newmail.BCC = "bridge@pcen.com.au"
    newmail.GetInspector()
    index = newmail.HTMLbody.find(">", newmail.HTMLbody.find("<body"))
    first_name = data["Project Info"]["Main Contact"]["Main Contact Full Name"].get().split(" ")[0]
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


def chase(app, *args):
    data = app.data
    if not data["State"]["Fee Accepted"].get():
        messagebox.showerror("Error", "Please Sent the fee proposal to client first")
        return

    if not "OUTLOOK.EXE" in (p.name() for p in psutil.process_iter()):
        os.startfile("outlook")
    ol = win32client.Dispatch("Outlook.Application")
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = f'Re: {data["Project Info"]["Project"]["Quotation Number"].get()}-{data["Project Info"]["Project"]["Project Name"].get()}'
    newmail.To = f'{data["Project Info"]["Client"]["Contact Email"].get()}; {data["Project Info"]["Main Contact"]["Main Contact Email"].get()}'
    newmail.CC = "felix@pcen.com.au"
    newmail.BCC = "bridge@pcen.com.au"
    newmail.GetInspector()
    index = newmail.HTMLbody.find(">", newmail.HTMLbody.find("<body"))
    first_name = data["Project Info"]["Main Contact"]["Main Contact Full Name"].get().split(" ")[0]
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
    inv = f"INV{str(inv+1)}"
    data = app.data
    database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get())
    pdf_name = inv+".pdf"

    # if not data["State"]["Fee Accepted"].get():
    #     messagebox.showerror("Error", "Please Sent the fee proposal to client first")
    #     return

    if not "OUTLOOK.EXE" in (p.name() for p in psutil.process_iter()):
        os.startfile("outlook")
    ol = win32client.Dispatch("Outlook.Application")
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = f'Invoice {data["Financial Panel"]["Invoice Details"][inv]["Number"].get()}-{data["Project Info"]["Project"]["Project Name"].get()}'
    newmail.To = f'{data["Project Info"]["Client"]["Contact Email"].get()}; {data["Project Info"]["Main Contact"]["Main Contact Email"].get()}'
    newmail.CC = "felix@pcen.com.au"
    newmail.BCC = "bridge@pcen.com.au"
    newmail.GetInspector()
    index = newmail.HTMLbody.find(">", newmail.HTMLbody.find("<body"))
    first_name = data["Project Info"]["Client"]["Client Full Name"].get().split(" ")[0]
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
    save(app)
    config_state(app)