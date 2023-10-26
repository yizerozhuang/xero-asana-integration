import tkinter as tk
from tkinter import messagebox

from project_info_page import ProjectInfoPage
from fee_proposal_page import FeeProposalPage
from invoice_and_bill import InvoicePage
from utility import rename_new_folder, excel_print_pdf, email, convert_to_json, convert_to_data
from asana_function import update_asana
from custom_dialog import ConfirmDialog

import psycopg2
from PIL import Image, ImageTk
import os
import json

class App(tk.Tk):
    def __init__(self, CONFIUGRATION, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        self.title("Premium Consulting Engineers")
        self.conf = CONFIUGRATION
        logo = Image.open(self.conf["resource_dir"]+"\\jpg\\logo.jpg")
        render = ImageTk.PhotoImage(logo)
        
        self.iconphoto(False, render)
        self.state("zoomed")
        self.data = {}
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

        self.protocol("WM_DELETE_WINDOW", self.confirm)

    def utility_frame(self):
        # Utility Part
        utility_frame = tk.LabelFrame(self.main_frame, text="Utility", font=self.conf["font"])
        utility_frame.pack(side=tk.TOP)

        rename_folder = tk.Button(utility_frame, text="Rename folder", command=lambda: rename_new_folder(self),
                                  bg="brown",
                                  fg="white", font=self.conf["font"])
        rename_folder.grid(row=0, column=0)

        print_button = tk.Button(utility_frame, text="Preview Fee Proposal", bg="brown",
                                 command=lambda: excel_print_pdf(self), fg="white",
                                 font=self.conf["font"])
        print_button.grid(row=0, column=1)

        email_button = tk.Button(utility_frame, text="Email to Client", command=lambda: email(self.data), bg="brown",
                                 fg="white", font=self.conf["font"])
        email_button.grid(row=0, column=2)

        update_asana_button = tk.Button(utility_frame, text="Update Asana", command=lambda: update_asana(self.data),
                                        bg="brown", fg="white", font=self.conf["font"])
        update_asana_button.grid(row=0, column=3)

        update_asana_button = tk.Button(utility_frame, text="Update Xero", bg="brown", fg="white", font=self.conf["font"])
        update_asana_button.grid(row=0, column=4)

        self.data["State"] = {
            "Folder Renamed": tk.BooleanVar(),
            "Fee Proposal Issued": tk.BooleanVar(),
            "Email to Client": tk.BooleanVar(),
            "Update Asana": tk.BooleanVar(),
            "Update Xero": tk.BooleanVar()
        }
        tk.Checkbutton(utility_frame, variable=self.data["State"]["Folder Renamed"], state=tk.DISABLED,
                       text="Folder Renamed").grid(row=1, column=0)
        tk.Checkbutton(utility_frame, variable=self.data["State"]["Fee Proposal Issued"], state=tk.DISABLED,
                       text="Fee Proposal Issued").grid(row=1, column=1)
        tk.Checkbutton(utility_frame, variable=self.data["State"]["Email to Client"], state=tk.DISABLED,
                       text="Email to Client").grid(row=1, column=2)
        tk.Checkbutton(utility_frame, variable=self.data["State"]["Update Asana"], state=tk.DISABLED,
                       text="Update Asana").grid(row=1, column=3)
        tk.Checkbutton(utility_frame, variable=self.data["State"]["Update Xero"], state=tk.DISABLED,
                       text="Update Xero").grid(row=1, column=4)

    def change_page_frame(self):
        # change page
        change_page_frame = tk.LabelFrame(self.main_frame, text="Change Page")
        change_page_frame.pack(side=tk.BOTTOM)

        proj_info_button = tk.Button(change_page_frame, text="Project Info",
                                     command=lambda: self.show_frame(self.page_info_page), bg="brown", fg="white",
                                     font=self.conf["font"])
        proj_info_button.grid(row=0, column=0)
        fee_proposal_button = tk.Button(change_page_frame, text="Fee Proposal",
                                        command=lambda: self.show_frame(self.fee_proposal_page), bg="brown", fg="white",
                                        font=self.conf["font"])
        fee_proposal_button.grid(row=0, column=1)
        invoice_button = tk.Button(change_page_frame, text="Invoice",
                                   command=lambda: self.show_frame(self.invoice_page), bg="brown", fg="white",
                                   font=self.conf["font"])
        invoice_button.grid(row=0, column=2)

    def main_context_frame(self):
        # main frame page
        self.page_info_page = ProjectInfoPage(self.main_frame, self)
        self.fee_proposal_page = FeeProposalPage(self.main_frame, self)
        self.invoice_page = InvoicePage(self.main_frame, self)

    def show_frame(self, page):
        self.page_info_page.pack_forget()
        self.fee_proposal_page.pack_forget()
        self.invoice_page.pack_forget()
        page.pack(fill=tk.BOTH, expand=1)

    def _ist_update(self, fee, ingst):
        tax_rate = self.conf["tax rates"]
        if len(fee.get()) == 0:
            ingst.set("")
            return
        try:
            num = round(float(fee.get()) * tax_rate, 2)
        except ValueError:
            ingst.set("Error")
            return
        ingst.set(str(num))

    def _sum_update(self, fee_list, total, *args):
        sum = 0
        for fee in fee_list:
            if fee.get() == "":
                continue
            try:
                sum += float(fee.get())
            except ValueError:
                total.set("Error")
                return
        total.set(str(round(sum, 2)))

    def save(self):
        data_json = convert_to_json(self.data)
        msg_box = True
        if not self.data["State"]["Folder Renamed"].get():
            messagebox.showerror("Error", "You should rename the folder before you save it to the database")
            return
        elif not os.path.exists(
                os.getcwd() + "\\database\\" + data_json["Project Info"]["Project"]["Quotation Number"]):
            os.mkdir(os.getcwd() + "\\database\\" + data_json["Project Info"]["Project"]["Quotation Number"])
        elif os.path.exists(os.getcwd() + "\\database\\" + data_json["Project Info"]["Project"][
            "Quotation Number"] + "\\data.json"):
            msg_box = messagebox.askyesno("Overwrite", "Existing Data found, do you want to overwrite")
        if msg_box:
            with open(os.getcwd() + "\\database\\" + data_json["Project Info"]["Project"][
                "Quotation Number"] + "\\data.json", "w") as f:
                json_object = json.dumps(data_json, indent=4)
                f.write(json_object)
            messagebox.showinfo("Saved", "It is saved into database")

    def load(self):
        if len(self.data["Project Info"]["Project"]["Quotation Number"].get()) == 0:
            messagebox.showerror("Error", "Please enter a Quotation Number before you load")
            return
        elif not os.path.exists(
                os.getcwd() + "\\database\\" + self.data["Project Info"]["Project"]["Quotation Number"].get()):
            messagebox.showerror("Error", "The quotation number doesn't exist")
            return
        f = open(os.getcwd() + "\\database\\" + self.data["Project Info"]["Project"][
            "Quotation Number"].get() + "\\data.json")
        data_json = json.load(f)
        convert_to_data(data_json, self.data)

    def confirm(self):
        if len(self.data["Project Info"]["Project"]["Quotation Number"].get()) == 0:
            self.destroy()
        else:
            ConfirmDialog(self)

