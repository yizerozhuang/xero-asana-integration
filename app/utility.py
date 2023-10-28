from tkinter import messagebox

from custom_dialog import FileSelectDialog

from win32com import client as win32client
import shutil
import os
import webbrowser
import json


def save(app):
    data = app.data
    data_json = convert_to_json(data)
    database_dir = os.path.join(app.conf["database_dir"], data_json["Project Info"]["Project"]["Quotation Number"])
    # write = True
    # if not data["State"]["Folder Renamed"].get():
    #     messagebox.showerror("Error", "You should rename the folder before you save it to the database")
    #     return
    if not os.path.exists(database_dir):
        os.mkdir(database_dir)
    # else:
    #     write = messagebox.askyesno("Overwrite", "Existing Data found, do you want to overwrite")
    # if write:
    with open(os.path.join(database_dir, "data.json"), "w") as f:
        json_object = json.dumps(data_json, indent=4)
        f.write(json_object)
    # messagebox.showinfo("Saved", "It is saved into database")


def load(app):
    data = app.data
    database_dir = os.path.join(app.conf["database_dir"], data["Project Info"]["Project"]["Quotation Number"].get().upper())
    data_json = json.load(open(os.path.join(database_dir, "data.json")))
    convert_to_data(data_json, data)


def convert_to_json(obj):
    if isinstance(obj, list):
        return [convert_to_json(x) for x in obj]
    elif isinstance(obj, dict):
        return {k: convert_to_json(v) for k, v in obj.items()}
    else:
        return obj.get()


def convert_to_data(json, data):
    if isinstance(json, list):
        [convert_to_data(json[i], data[i]) for i in range(len(json))]
    elif isinstance(json, dict):
        [convert_to_data(json[k], data[k]) for k in json.keys()]
    else:
        if data.get() != json:
            data.set(json)


def remove_none(obj):
    if isinstance(obj, list):
        return [remove_none(x) for x in obj if x is not None]
    elif isinstance(obj, dict):
        return {k: remove_none(v) for k, v in obj.items() if v is not None}
    else:
        return obj


def rename_new_folder(app):
    data = app.data
    folder_name = app.data["Project Info"]["Project"]["Quotation Number"].get() + "-" + \
                  app.data["Project Info"]["Project"]["Project Name"].get()
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
            shutil.move(os.path.join(app.conf["working_dir"], rename_list[0]), folder_path + "\\External")
            os.mkdir(os.path.join(folder_path, "Photos"), mode)
            os.mkdir(os.path.join(folder_path, "Plot"), mode)
            os.mkdir(os.path.join(folder_path, "SS"), mode)
            shutil.copyfile(os.path.join(app.conf["resource_dir"], "xlsx", "Preliminary Calculation v2.5.xlsx"),
                            os.path.join(folder_path, "Preliminary Calculation v2.5.xlsx"))
            app.data["State"]["Folder Renamed"].set(True)
            messagebox.showinfo(title="Folder renamed", message=f"Rename Folder {rename_list[0]} to {folder_name}")
        else:
            FileSelectDialog(app, rename_list, "Multiple new folders found, please select one")
    except FileExistsError:
        messagebox.showerror("Error", f"Cannot create a file when that file already exists:{folder_name}")
        return
    save(app)


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

    if not data["State"]["Folder Renamed"].get():
        messagebox.showerror("Error", "Please rename a folder first")
        return
    elif not os.path.exists(folder_path):
        messagebox.showerror("Error", "Can't find the folder, please rename first")
        return
    elif len(services) == 0:
        messagebox.showerror("Error", "Please at least select 1 service")
        return
    elif not page in [1, 2, 3]:
        messagebox.showerror("Error", "Excess the maximum value of service, please contact administrator")
    pdf_list = [file for file in os.listdir(os.path.join(database_dir)) if str(file).startswith("Mechanical Fee Proposal")]
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

    excel = win32client.Dispatch("Excel.Application")
    excel.ScreenUpdating = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    try:
        shutil.copyfile(os.path.join(resource_dir, "xlsx", f"fee_proposal_template_{page}.xlsx"),
                        os.path.join(folder_path, "fee proposal.xlsx"))
        work_book = excel.Workbooks.Open(os.path.join(folder_path, "fee proposal.xlsx"))
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
        # app.cur.execute(
        #     f"""
        #         SELECT *
        #         FROM past_project
        #         WHERE project_type='{project_type}'
        #     """
        # )
        # past_project = app.cur.fetchall()
        past_projects = json.load(open(past_projects_dir, encoding="utf-8"))[project_type]
        for i, project in enumerate(past_projects):
            work_sheets.Cells(cur_row + i, 1).Value = "•"
            work_sheets.Cells(cur_row + i, 2).Value = project
        cur_row += 34
        for i, service in enumerate(data['Fee Proposal']['Fee Details']["Details"].values()):
            work_sheets.Cells(cur_row + i, 2).Value = service["Service"].get() + " design and documentation"
            work_sheets.Cells(cur_row + i, 6).Value = service["Fee"].get()
            work_sheets.Cells(cur_row + i, 7).Value = service["in.GST"].get()
        if page == 3:
            work_sheets.Cells(cur_row + 4, 6).Value = data['Fee Proposal']['Fee Details']["Fee"].get()
            work_sheets.Cells(cur_row + 4, 7).Value = data['Fee Proposal']['Fee Details']["in.GST"].get()
        else:
            work_sheets.Cells(cur_row + 3, 6).Value = data['Fee Proposal']['Fee Details']["Fee"].get()
            work_sheets.Cells(cur_row + 3, 7).Value = data['Fee Proposal']['Fee Details']["in.GST"].get()
        cur_row += 17
        work_sheets.Cells(cur_row, 2).Value = "Re: " + data["Project Info"]["Project"]["Project Name"].get()
        work_sheets.ExportAsFixedFormat(0, os.path.join(folder_path, "Plot", pdf_name))
        work_book.Close(True)
        shutil.copyfile(
            os.path.join(folder_path, "Plot", pdf_name),
            os.path.join(database_dir, pdf_name)
        )
        shutil.copyfile(
            os.path.join(folder_path, "fee proposal.xlsx"),
            os.path.join(database_dir, "fee proposal.xlsx")
        )
        os.remove(os.path.join(folder_path, "fee proposal.xlsx"))
    except PermissionError:
        messagebox.showerror("Error", "Please close the preview or file before you use it")
    except FileNotFoundError:
        messagebox.showerror("Error", "Please Contact Administer, the app cant find the file")
    else:
        app.data["State"]["Fee Proposal Issued"].set(True)
        webbrowser.open(os.path.join(folder_path, "Plot", pdf_name))
    finally:
        excel.ScreenUpdating = True
        excel.DisplayAlerts = True
        excel.EnableEvents = True

def email(app, *args):
    data = app.data
    folder_name = data["Project Info"]["Project"]["Quotation Number"].get() + "-" + \
                  data["Project Info"]["Project"]["Project Name"].get()
    folder_path = os.path.join(app.conf["working_dir"], folder_name)
    pdf_name = f'Mechanical Fee Proposal revision {data["Fee Proposal"]["Reference"]["Revision"].get()}.pdf'

    if not data["State"]["Fee Proposal Issued"].get():
        messagebox.showerror("Error", "Please Generate a pdf first")
        return
    elif not os.path.exists(os.path.join(folder_path, "Plot", pdf_name)):
        messagebox.showerror("Error", f'Python cant found fee proposal for {data["Project Info"]["Project"]["Quotation Number"].get()} revision {data["Fee Proposal"]["Reference"]["Revision"].get()}')
        return

    ol = win32client.Dispatch("Outlook.Application")
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = "Mechanical Fee Proposal - " + data["Project Info"]["Project"]["Project Name"].get() + " revision " + data["Fee Proposal"]["Reference"]["Revision"].get()
    newmail.To = f'{data["Project Info"]["Client"]["Contact Email"].get()}; {data["Project Info"]["Main Contact"]["Main Contact Email"].get()}'
    newmail.CC = "felix@pcen.com.au"
    newmail.GetInspector()
    index = newmail.HTMLbody.find(">", newmail.HTMLbody.find("<body"))
    message = f"""
    Dear {data["Project Info"]["Client"]["Client Full Name"].get()},<br>

    I hope this email finds you well. Please find the attached fee proposal to this email.<br>

    If you have any questions or need more information regarding the proposal, please don't hesitate to reach out. I am happy to provide you with whatever information you need.<br>

    I look forward to hearing from you soon.<br>

    Cheers,<br>
    """
    newmail.HTMLbody = newmail.HTMLbody[:index + 1] + message + newmail.HTMLbody[index + 1:]
    newmail.Attachments.Add(os.path.join(folder_path, "Plot", pdf_name))
    newmail.Display()
    data["State"]["Email to Client"].set(True)
