import tkinter as tk
from tkinter import ttk, messagebox
from tkinterdnd2 import TkinterDnD

from app_log import AppLog
from project_info_page import ProjectInfoPage
from fee_proposal_page import FeeProposalPage
from fee_accpeted_page import FeeAcceptedPage
from financial_panel import FinancialPanelPage
from utility import rename_new_folder, excel_print_pdf, email, chase, save, load, config_state, config_log
from asana_function import update_asana
from xero_function import login_xero, update_xero

from PIL import Image, ImageTk
import time
import _thread
import os


class App(tk.Tk):
    def __init__(self, conf, user, *args, **kwargs):
        TkinterDnD.Tk.__init__(self, *args, **kwargs)
        # tk.Tk.__init__(self, *args, **kwargs)

        self.title("Premium Consulting Engineers")
        self.conf = conf
        self.user = user
        self.log = AppLog()
        self.log_text = tk.StringVar()
        logo = Image.open(os.path.join(self.conf["resource_dir"], "jpg", "logo.jpg"))
        render = ImageTk.PhotoImage(logo)

        self.iconphoto(False, render)
        self.state("zoomed")
        self.data = {}
        # the main frame
        self.main_frame = tk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=1)

        self.utility_frame()
        self.change_page_frame()
        self.main_context_frame()

        self.show_frame(self.page_info_page)

        self.protocol("WM_DELETE_WINDOW", self.confirm)
        config_state(self)
        self.auto_check()

    def utility_frame(self):
        # Utility Part
        utility_frame = tk.LabelFrame(self.main_frame, text="Utility", font=self.conf["font"])
        utility_frame.pack(side=tk.TOP)

        tk.Button(utility_frame, text="Rename Folder", command=lambda: rename_new_folder(self), bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=0)

        tk.Button(utility_frame, text="Preview Fee Proposal", bg="brown", command=lambda: excel_print_pdf(self),
                  fg="white",
                  font=self.conf["font"]).grid(row=0, column=1)

        tk.Button(utility_frame, text="Email to Client", command=lambda: email(self), bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=2)
        tk.Button(utility_frame, text="Chase Client", command=lambda: chase(self), bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=3)

        tk.Button(utility_frame, text="Update Asana", command=lambda: update_asana(self), bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=4)

        tk.Button(utility_frame, text="Login Xero", command=login_xero, bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=1, column=4)

        tk.Button(utility_frame, text="Update Xero", command=self._update_xero, bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=2, column=4)

        self.data["State"] = {
            "Set Up": tk.BooleanVar(),
            "Generate Proposal": tk.BooleanVar(),
            "Email to Client": tk.BooleanVar(),
            "Fee Accepted": tk.BooleanVar(),
            "Done": tk.BooleanVar(),
            "Quote Unsuccessful": tk.BooleanVar()
        }

        tk.Checkbutton(utility_frame, variable=self.data["State"]["Set Up"], state=tk.DISABLED,
                       text="Set Up").grid(row=1, column=0)
        tk.Checkbutton(utility_frame, variable=self.data["State"]["Generate Proposal"], state=tk.DISABLED,
                       text="Generate Proposal").grid(row=1, column=1)
        tk.Checkbutton(utility_frame, variable=self.data["State"]["Email to Client"], state=tk.DISABLED,
                       text="Email to Client").grid(row=1, column=2)
        tk.Checkbutton(utility_frame, variable=self.data["State"]["Fee Accepted"], state=tk.DISABLED,
                       text="Fee Accepted").grid(row=1, column=3)
        self.load_project_quotation = tk.StringVar()
        self.load_project_quotation.trace("w", self.load_project)
        self.state_dict = {
            "Set Up": ttk.Combobox(utility_frame, textvariable=self.load_project_quotation),
            "Generate Proposal": ttk.Combobox(utility_frame, textvariable=self.load_project_quotation),
            "Email to Client": ttk.Combobox(utility_frame, textvariable=self.load_project_quotation),
            "Fee Accepted": ttk.Combobox(utility_frame, textvariable=self.load_project_quotation)
        }

        self.state_dict["Set Up"].grid(row=2, column=0)
        self.state_dict["Generate Proposal"].grid(row=2, column=1)
        self.state_dict["Email to Client"].grid(row=2, column=2)
        self.state_dict["Fee Accepted"].grid(row=2, column=3)

        self.data["Email"] = {
            "Fee Proposal": tk.StringVar(),
            "First Chase": tk.StringVar(),
            "Second Chase": tk.StringVar(),
            "Third Chase": tk.StringVar()
        }

    def change_page_frame(self):
        # change page
        change_page_frame = tk.LabelFrame(self.main_frame, text="Change Page")
        change_page_frame.pack(side=tk.BOTTOM)

        tk.Button(change_page_frame, text="Project Info",
                  command=lambda: self.show_frame(self.page_info_page), bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=0)
        tk.Button(change_page_frame, text="Fee Proposal",
                  command=lambda: self.show_frame(self.fee_proposal_page), bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=1)
        tk.Button(change_page_frame, text="Log Files",
                  command=lambda: self.show_frame(self.fee_accepted_page), bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=2)
        tk.Button(change_page_frame, text="Financial Panel",
                  command=lambda: self.show_frame(self.financial_panel_page), bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=3)

    def main_context_frame(self):
        # main frame page
        self.page_info_page = ProjectInfoPage(self.main_frame, self)
        self.fee_proposal_page = FeeProposalPage(self.main_frame, self)
        self.fee_accepted_page = FeeAcceptedPage(self.main_frame, self)
        self.financial_panel_page = FinancialPanelPage(self.main_frame, self)

    def show_frame(self, page):
        self.page_info_page.pack_forget()
        self.fee_proposal_page.pack_forget()
        self.fee_accepted_page.pack_forget()
        self.financial_panel_page.pack_forget()
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

    def load_project(self, *args):
        if len(self.load_project_quotation.get()) == 0:
            return
        self.data["Project Info"]["Project"]["Quotation Number"].set(self.load_project_quotation.get().split("-")[0])
        load(self)
        self.load_project_quotation.set("")
        for drop_down in self.state_dict.values():
            drop_down.set("")

    def auto_check(self):
        def thread_task():
            while True:
                time.sleep(10)
                # save(self)
                config_state(self)
                if len(self.data["Project Info"]["Project"]["Quotation Number"].get()) != 0:
                    try:
                        config_log(self)
                    except:
                        continue

        _thread.start_new_thread(thread_task, ())

    def confirm(self):
        if len(self.data["Project Info"]["Project"]["Quotation Number"].get()) == 0:
            self.destroy()
        else:
            save(self)
            self.destroy()

    def _update_xero(self):
        if not self.data["State"]["Done"].get():
            messagebox.showerror("Error", "You haven't update a fee accountant yet, please update a fee acceptance before you update xero")
            return
        elif len(self.data["Financial Panel"]["Invoice Details"]["INV1"]["Number"].get())==0:
            messagebox.showerror("Error", "Please Generate An Invoice Number Fist")
            return


        if len(self.data["Project Info"]["Client"]["Client Full Name"].get()) == 0:
            if len(self.data["Project Info"]["Client"]["Client Company"].get()) == 0:
                messagebox.showerror("Error", "You should at least provide client name or client company")
                return
            else:
                contact = self.data["Project Info"]["Client"]["Client Company"].get()
        else:
            if len(self.data["Project Info"]["Client"]["Client Company"].get()) == 0:
                contact = self.data["Project Info"]["Client"]["Client Full Name"].get()
            else:
                contact = self.data["Project Info"]["Client"]["Client Company"].get() + ", " + self.data["Project Info"]["Client"]["Client Full Name"].get()
        try:
            update_xero(self, contact)
        except RuntimeError:
            messagebox.showerror("Error", "You  should login into xero first")