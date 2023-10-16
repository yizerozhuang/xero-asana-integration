import tkinter as tk
from tkinter import messagebox
import pdfkit
from function.utility import *
from function.project_info_page import ProjectInfoPage
from function.fee_proposal_page import FeeProposalPage
from function.invoice_page import InvoicePage
from function.custom_dialog import FileSelectDialog
import psycopg2
import os
import shutil

from win32com import client


class App(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        self.title("Data Entry Form")
        self.state("zoomed")
        self.font = ("Calibri", 11)
        self.data = {}
        # self.attributes("-fullscreen", True)
        # the main frame
        self.main_frame = tk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=1)
        self.conn = psycopg2.connect(
            host="127.0.0.1",
            database="postgres",
            user="postgres",
            password="Zero0929"
        )
        self.cur = self.conn.cursor()

        self.utility_frame()
        self.change_page_frame()
        self.main_context_frame()

        self.show_frame(self.page_info_page)

    def utility_frame(self):
        # Utility Part
        utility_frame = tk.LabelFrame(self.main_frame, text="Utility", font=self.font)
        utility_frame.pack(side=tk.TOP)

        # create_folder = tk.Button(utility_frame, text="Create new Folder", command=create_new_folder, bg="brown",
        #                           fg="white", font=self.font)
        # create_folder.grid(row=0, column=0)

        rename_folder = tk.Button(utility_frame, text="Rename folder", command=self.rename_new_folder, bg="brown",
                                  fg="white", font=self.font)
        rename_folder.grid(row=0, column=0)

        update_asana_button = tk.Button(utility_frame, text="Update Asana", command=lambda: update_asana(self.data), bg="brown", fg="white", font=self.font)
        update_asana_button.grid(row=0, column=1)

        # print_button = tk.Button(utility_frame, text="Print", bg="brown", command=self._print_pdf, fg="white",
        #                          font=self.font)
        print_button = tk.Button(utility_frame, text="Print", bg="brown", command=lambda: excel_print_pdf(self.data), fg="white",
                                 font=self.font)
        print_button.grid(row=0, column=2)

        email_button = tk.Button(utility_frame, text="Email", command=lambda: email(self.data), bg="brown", fg="white", font=self.font)
        email_button.grid(row=0, column=3)

        # xero_button = tk.Button(utility_frame, text="Update xero", command=lambda:update_xero(self.data), bg="brown", fg="white", font=self.font)
        # xero_button.grid(row=0, column=4)

    def change_page_frame(self):
        # change page
        change_page_frame = tk.LabelFrame(self.main_frame, text="Change Page")
        change_page_frame.pack(side=tk.BOTTOM)

        proj_info_button = tk.Button(change_page_frame, text="Project Info",
                                     command=lambda: self.show_frame(self.page_info_page), bg="brown", fg="white",
                                     font=self.font)
        proj_info_button.grid(row=0, column=0)
        fee_proposal_button = tk.Button(change_page_frame, text="Fee Proposal",
                                        command=lambda: self.show_frame(self.fee_proposal_page), bg="brown", fg="white",
                                        font=self.font)
        fee_proposal_button.grid(row=0, column=1)
        invoice_button = tk.Button(change_page_frame, text="Invoice Proposal",
                                        command=lambda: self.show_frame(self.invoice_page), bg="brown", fg="white",
                                        font=self.font)
        invoice_button.grid(row=0, column=2)

    def main_context_frame(self):
        # main frame page
        self.page_info_page = ProjectInfoPage(self.main_frame, self)
        self.fee_proposal_page = FeeProposalPage(self.main_frame, self)
        self.invoice_page = InvoicePage(self.main_frame, self)

    def show_frame(self, cont):
        self.page_info_page.pack_forget()
        self.fee_proposal_page.pack_forget()
        self.invoice_page.pack_forget()
        cont.pack(fill=tk.BOTH, expand=1)

    def update_service_type(self, var):
        self.fee_proposal_page.update_scope_frame(var)
        self.fee_proposal_page.update_fee(var)
        self.invoice_page.update_fee(var)
        self.invoice_page.update_bill(var)
        self.invoice_page.update_profit(var)

    def rename_new_folder(self):
        if self.data["Project Information"]["Project Name"].get()=="":
            messagebox.showerror(title="Error", message="Please Input a project name")
            return
        elif self.data["Project Information"]["Quotation Number"].get()=="":
            messagebox.showerror(title="Error", message="Please Input a quotation number")
            return
        dir_list = os.listdir(".")
        rename_list = [dir for dir in dir_list if dir.startswith("New folder")]
        if len(rename_list)==0:
            messagebox.showerror(title="Error", message="Please create a new folder first")
            return
        elif len(rename_list)==1:
            mode = 0o666
            os.mkdir(rename_list[0] + "/External", mode)
            external_files = os.listdir(rename_list[0])
            for file in external_files:
                shutil.move(rename_list[0] + "/" + file, rename_list[0] + "/External")
            os.mkdir(rename_list[0] + "/Photos", mode)
            os.mkdir(rename_list[0] + "/Plot", mode)
            os.mkdir(rename_list[0] + "/SS", mode)
            shutil.copyfile("resource/xlsx/Preliminary Calculation v2.5.xlsx",
                            rename_list[0] + "/Preliminary Calculation v2.5.xlsx")

            os.rename(rename_list[0],
                      self.data["Project Information"]["Quotation Number"].get() + "-" +
                      self.data["Project Information"]["Project Name"].get())
            messagebox.showinfo(title="Folder renamed", message="Folder renamed")
        else:
            FileSelectDialog(self, self, rename_list, "Multiple new folders found, please select one")

    def _print_pdf(self):
        options = {
            "page-size":"A4",
            "margin-bottom": "2mm",
            'zoom': 2
        }
        with open("resource/html/cur_html.html", "w") as f:
            f.write(open("resource/html/sample_html.html").read().format(
                Client_Full_Name=self.data["Client"]["Client Full Name"].get(),
                Client_Company=self.data["Client"]["Client Company"].get(),
                Reference=self.data["Project Information"]["Quotation Number"].get(),
                Date=self.data["Fee Proposal Page"]["Date"].get(),
                Revision=self.data["Fee Proposal Page"]["Revision"].get(),
                Project_Name=self.data["Project Information"]["Project Name"].get(),
                Fee_Stage_Start=self.data["Fee Proposal Page"]["Time"]["Fee proposal stage Duration"][0].get(),
                Fee_Stage_End=self.data["Fee Proposal Page"]["Time"]["Fee proposal stage Duration"][1].get(),
                Design_Stage_Start=self.data["Fee Proposal Page"]["Time"]["Pre-design information collection Duration"][
                    0].get(),
                Design_Stage_End=self.data["Fee Proposal Page"]["Time"]["Pre-design information collection Duration"][
                    1].get(),
                Issue_Stage_Start=self.data["Fee Proposal Page"]["Time"]["Documentation and issue Duration"][0].get(),
                Issue_Stage_End=self.data["Fee Proposal Page"]["Time"]["Documentation and issue Duration"][1].get(),
                Scope_of_Work=self._format_scope(),
                Past_Project=self._format_experience(),
                Details=self._format_details()
                )
            )
        pdfkit.from_file("resource/html/cur_html.html", "output.pdf", options=options)

        #option choice to convet to a excel first
        # excel = client.Dispatch("Excel.Application")
        # sheets = excel.Workbooks.Open("C:/Users/yeezh/xero-asana-intergration/app/sample.xlsx")
        # work_sheets = sheets.Worksheets[2]
        # work_sheets.ExportAsFixedFormat(0, "output.pdf")

    def _format_details(self):
        res = ""
        for key, value in self.data["Fee Proposal Page"]["Details"].items():
            if value["on"].get():
                ingst = int(float(value["Fee"].get()) * 1.1) if value["Fee"].get() != "" else 0
                res += f"""
                 </tr>
                  <tr height=20 style='mso-height-source:userset;height:15.0pt'>
                  <td height=20 class=xl1525308 style='height:15.0pt'></td>
                  <td colspan=4 class=xl11625308 style='border-right:.5pt solid black'>{key} design and documentation</td>
                  <td class=xl9625308 align=right style='border-top:none;border-left:none'>{value["Fee"].get()}</td>
                  <td class=xl9625308 align=right style='border-top:none;border-left:none'>${ingst}</td>
                  <td class=xl8825308></td>
                  <td class=xl7325308></td>
                 </tr>
                """
        return res

    def _format_scope(self):
        data = self._return_scope()
        res = ""
        i = 1
        for service in data.keys():
            for extra in data[service].keys():
                i += 1
                res += f"""
                <tr height=20 style='mso-height-source:userset;height:15.0pt'>
                <td height=20 class=xl8625308 style='height:15.0pt'>2.{i}</td>
                <td class=xl6925308 colspan=3>{service} - {extra}</td>
                <td class=xl7425308></td>
                <td class=xl7425308></td>
                <td class=xl7425308></td>
                <td class=xl1525308></td>
                <td class=xl8625308></td>
                </tr>
                """
                for context in data[service][extra]:
                    res += f"""
                     <tr height=20 style='mso-height-source:userset;height:15.0pt'>
                      <td height=20 class=xl6825308 style='height:15.0pt'>•</td>
                      <td class=xl1525308 colspan=4>{context}</td>
                      <td class=xl1525308></td>
                      <td class=xl6825308></td>
                      <td class=xl1525308></td>
                      <td class=xl6825308></td>
                     </tr>
                    """
        return res

    def _format_experience(self):
        res = ""
        project_type = self.data["Project Information"]["Project Type"].get()
        self.cur.execute(
            f"""
                SELECT *
                FROM past_project
                WHERE project_type='{project_type}'
            """
        )
        past_project = self.cur.fetchall()
        for project in past_project:
            res += f"""
             <tr height=20 style='mso-height-source:userset;height:15.0pt'>
              <td height=20 class=xl6825308 style='height:15.0pt'>•</td>
              <td class=xl1525308 colspan=4>{project[1]}</td>
              <td class=xl1525308></td>
              <td class=xl6825308></td>
              <td class=xl1525308></td>
              <td class=xl6825308></td>
             </tr>
            """
        return res

    def _return_scope(self):
        res = dict()
        for service_name, service_context in self.data["Fee Proposal Page"]["Scope of Work"].items():
            if service_context["on"]:
                res[service_name] = dict()
                for extra in ["Extend", "Exclusion", "Deliverables"]:
                    res[service_name][extra] = []
                    for context in self.data["Fee Proposal Page"]["Scope of Work"][service_name][extra]:
                        if context[0].get():
                            res[service_name][extra].append(context[1].get())
        return res

    def _sum_update(self, area_list, label, *args):
        sum = 0
        for area in area_list:
            if area is None or area.get() == "":
                continue
            try:
                sum += float(area.get())
            except ValueError:
                sum = "Error"
                if type(label) == tk.Label:
                    label.config(text=str(sum), bg="red")
                elif type(label) ==tk.StringVar:
                    label.set(str(sum))
                return
        if type(label) == tk.Label:
            label.config(text=str(int(sum)), bg=self.cget("bg"))
        elif type(label) == tk.StringVar:
            label.set(str(int(sum)))
    def _ist_update(self, string_variable, label):
        if type(label)==tk.Label:
            if len(string_variable.get()) == 0:
                label.config(text="0", bg=self.cget("bg"))
                return
            try:
                num = int(float(string_variable.get())*1.1)
            except ValueError:
                label.config(text="Error", bg="red")
                return
            label.config(text=str(num), bg=self.cget("bg"))
        elif type(label)==tk.StringVar:
            if len(string_variable.get())==0:
                label.set("")
                return
            try:
                num = int(float(string_variable.get()) * 1.1)
            except ValueError:
                label.set("Error")
                return
            label.set(str(num))

    def update_invoice_sum(self, details, invoice_list):
        new_value=[0]*6
        for service in details.values():
            if not service["on"].get():
                continue
            if service["Expanded"].get():
                for i in range(3):
                    if service["Context"][i]["Invoice"].get() != "None":
                        index = int(service["Context"][i]["Invoice"].get()[3]) - 1
                        try:
                            new_value[index] += int(float(service["Context"][i]["Fee"].get()))
                        except ValueError:
                            continue
            else:
                if service["Invoice"].get() != "None":
                    index = int(service["Invoice"].get()[3]) - 1
                    try:
                        new_value[index] += int(float(service["Fee"].get()))
                    except ValueError:
                        continue
        for j in range(6):
            invoice_list[j].set(str(new_value[j]))



