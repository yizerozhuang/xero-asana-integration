import tkinter as tk
from tkinter import ttk
# from tkinterdnd2 import TkinterDnD

from app_log import AppLog
from project_info_page import ProjectInfoPage
from fee_proposal_page import FeeProposalPage
# from fee_accpeted_page import FeeAcceptedPage
from financial_panel import FinancialPanelPage
from utility import *
from asana_function import update_asana, rename_asana_project
from xero_function import login_xero, update_xero
from email_server import email_server

from PIL import Image, ImageTk
import time
import _thread
import os


class App(tk.Tk):
    def __init__(self, conf, user, *args, **kwargs):
        # TkinterDnD.Tk.__init__(self, *args, **kwargs)
        tk.Tk.__init__(self, *args, **kwargs)

        self.title("Premium Consulting Engineers")
        self.conf = conf
        self.user = user
        self.log = AppLog()
        self.log_text = tk.StringVar()
        logo = Image.open(os.path.join(self.conf["resource_dir"], "jpg", "logo.jpg"))
        render = ImageTk.PhotoImage(logo)
        # self.option_add('*Dialog.msg.font', 'Helvetica 24')
        # self.master.option_add('*Dialog.msg.width', 34)
        # self.master.option_add("*Dialog.msg.wrapLength", "6i")
        # self.option_add('*Dialog.msg.font', self.conf[    "msg_font"])
        self.iconphoto(False, render)
        self.state("zoomed")
        self.data = {}
        # the main frame
        self.main_frame = tk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=1)

        self.utility_frame()
        self.change_page_frame()
        self.main_context_frame()

        self.show_frame(self.project_info_page)

        self.protocol("WM_DELETE_WINDOW", self.confirm)
        config_state(self)
        self.auto_check()

    def utility_frame(self):
        # Utility Part
        utility_frame = tk.LabelFrame(self.main_frame, text="Utility", font=self.conf["font"])
        utility_frame.pack(side=tk.TOP)
        
        state_frame = tk.LabelFrame(utility_frame, font=self.conf["font"])
        state_frame.pack(side=tk.LEFT)

        tk.Button(state_frame, text="Set Up Project", command=self._finish_setup, bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=0)
        tk.Button(state_frame, text="Preview Fee Proposal", bg="brown", command=lambda: excel_print_pdf(self), fg="white",
                  font=self.conf["font"]).grid(row=0, column=1)
        tk.Button(state_frame, text="Email to Client", command=lambda: email(self), bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=2)
        tk.Button(state_frame, text="Chase Client", command=lambda: chase(self), bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=3)
        
        function_frame = tk.LabelFrame(utility_frame, font=self.conf["font"])
        function_frame.pack(side=tk.RIGHT, fill=tk.Y)

        tk.Button(function_frame, width=10, text="Open Folder", command=self.open_folder, bg="brown",
                  fg="white",
                  font=self.conf["font"]).grid(row=0, column=0)

        tk.Button(function_frame, width=10, text="Open Database", command=self.open_database, bg="brown",
                  fg="white",
                  font=self.conf["font"]).grid(row=1, column=0)

        tk.Button(function_frame, text="Rename Project", command=self._rename_project, bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=1)

        tk.Button(function_frame, text="Update Asana", command=lambda: update_asana(self), bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=2)

        tk.Button(function_frame, text="Login Xero", command=login_xero, bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=3)

        tk.Button(function_frame, text="Update Xero", command=self._update_xero, bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=1, column=3)

        self.data["State"] = {
            "Set Up": tk.BooleanVar(),
            "Generate Proposal": tk.BooleanVar(),
            "Email to Client": tk.BooleanVar(),
            "Fee Accepted": tk.BooleanVar(),
            "Quote Unsuccessful": tk.BooleanVar()
        }

        tk.Label(state_frame, text="Set Up").grid(row=1, column=0)
        tk.Label(state_frame, text="Generate Proposal").grid(row=1, column=1)
        tk.Label(state_frame, text="Email to Client").grid(row=1, column=2)
        tk.Label(state_frame, text="Fee Accepted").grid(row=1, column=3)

        # tk.Checkbutton(state_frame, variable=self.data["State"]["Set Up"], state=tk.DISABLED,
        #                text="Set Up").grid(row=1, column=0)
        # tk.Checkbutton(state_frame, variable=self.data["State"]["Generate Proposal"], state=tk.DISABLED,
        #                text="Generate Proposal").grid(row=1, column=1)
        # tk.Checkbutton(state_frame, variable=self.data["State"]["Email to Client"], state=tk.DISABLED,
        #                text="Email to Client").grid(row=1, column=2)
        # tk.Checkbutton(state_frame, variable=self.data["State"]["Fee Accepted"], state=tk.DISABLED,
        #                text="Fee Accepted").grid(row=1, column=3)

        self.set_up_state=tk.Label(state_frame, bg="yellow")
        self.set_up_state.grid(row=2, column=0, sticky="ew")
        self.proposal_state=tk.Label(state_frame, bg="red")
        self.proposal_state.grid(row=2, column=1, sticky="ew")
        self.email_state=tk.Label(state_frame, bg="red")
        self.email_state.grid(row=2, column=2, sticky="ew")
        self.accept_state=tk.Label(state_frame, bg="red")
        self.accept_state.grid(row=2, column=3, sticky="ew")

        self.data["State"]["Set Up"].trace("w", self._update_rec)
        self.data["State"]["Generate Proposal"].trace("w", self._update_rec)
        self.data["State"]["Email to Client"].trace("w", self._update_rec)
        self.data["State"]["Fee Accepted"].trace("w", self._update_rec)

        self.load_project_quotation = tk.StringVar()
        self.load_project_quotation.trace("w", self.load_project)
        self.state_dict = {
            "Set Up": ttk.Combobox(state_frame, textvariable=self.load_project_quotation),
            "Generate Proposal": ttk.Combobox(state_frame, textvariable=self.load_project_quotation),
            "Email to Client": ttk.Combobox(state_frame, textvariable=self.load_project_quotation),
            "Fee Accepted": ttk.Combobox(state_frame, textvariable=self.load_project_quotation)
        }

        self.state_dict["Set Up"].grid(row=3, column=0)
        self.state_dict["Generate Proposal"].grid(row=3, column=1)
        self.state_dict["Email to Client"].grid(row=3, column=2)
        self.state_dict["Fee Accepted"].grid(row=3, column=3)

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
                  command=lambda: self.show_frame(self.project_info_page), bg="green", fg="black",
                  font=self.conf["font"]).grid(row=0, column=0)
        tk.Button(change_page_frame, text="Fee Details",
                  command=lambda: self.show_frame(self.fee_proposal_page), bg="green", fg="black",
                  font=self.conf["font"]).grid(row=0, column=1)
        # tk.Button(change_page_frame, text="Log Files",
        #           command=lambda: self.show_frame(self.fee_accepted_page), bg="green", fg="black",
        #           font=self.conf["font"]).grid(row=0, column=2)
        tk.Button(change_page_frame, text="Financial Panel",
                  command=lambda: self.show_frame(self.financial_panel_page), bg="green", fg="black",
                  font=self.conf["font"]).grid(row=0, column=3)

    def main_context_frame(self):
        # main frame page
        self.project_info_page = ProjectInfoPage(self.main_frame, self)
        # self.fee_accepted_page = FeeAcceptedPage(self.main_frame, self)
        self.fee_proposal_page = FeeProposalPage(self.main_frame, self)
        self.financial_panel_page = FinancialPanelPage(self.main_frame, self)
        self._update_variation()
    def _update_variation(self):
        # variation = [
        #     {
        #         "Service": tk.StringVar(),
        #         "Fee": tk.StringVar(),
        #         "in.GST": tk.StringVar(),
        #         "Number": tk.StringVar(value="None")
        #     } for _ in range(self.conf["n_variation"])
        # ]
        # self.data["Variation"] = variation
        # ist_update_fuc = lambda i: lambda a, b, c: self.app._ist_update(variation[i]["Fee"], variation[i]["in.GST"])
        # for i in range(self.conf["n_variation"]):
        #     variation[i]["Fee"].trace("w", ist_update_fuc(i))
        #     variation[i]["Fee"].trace("w", self.update_sum)
        #
        # for i in range(self.conf["n_variation"]):
        #     variation_frame = tk.LabelFrame(self.fee_frame)
        #     variation_frame.pack(side=tk.BOTTOM, fill=tk.X)
        #     tk.Label(variation_frame, width=10, text="", font=self.conf["font"]).grid(row=0, column=0)
        #     tk.Entry(variation_frame, width=50, textvariable=variation[i]["Service"], font=self.conf["font"],
        #              fg="blue").grid(row=0, column=1)
        #     tk.Entry(variation_frame, width=20, textvariable=variation[i]["Fee"], font=self.conf["font"],
        #              fg="blue").grid(row=0, column=2, padx=(40, 0))
        #     tk.Label(variation_frame, width=20, textvariable=variation[i]["in.GST"], font=self.conf["font"]).grid(row=0,
        #                                                                                                           column=3)

        variation_var = {
            "Service": tk.StringVar(value="Variation"),
            "Include": tk.BooleanVar(value=True)
        }

        self.fee_proposal_page.update_fee(variation_var)
        self.fee_proposal_page.fee_dic["Variation"]["Expand"].grid_forget()
        tk.Label(self.fee_proposal_page.fee_frames["Variation"], text="", width=10).grid(row=0, column=0)
        self.fee_proposal_page.fee_dic["Variation"]["Service"].config(text="Variation")
        self.fee_proposal_page.fee_frames["Variation"].pack(side=tk.BOTTOM)

        self.financial_panel_page.update_invoice(variation_var)
        self.financial_panel_page.invoice_dic["Variation"]["Service"].config(text="Variation")

        self.financial_panel_page.update_bill(variation_var)
        self.financial_panel_page.update_profit(variation_var)

        self.data["Invoices"]["Details"]["Variation"]["Expand"].set(True)

        for i in range(self.conf["n_items"]-1, -1, -1):
            self.financial_panel_page.invoice_dic["Variation"]["Expand"][i].pack_forget()
            self.financial_panel_page.invoice_dic["Variation"]["Expand"][i].pack(fill=tk.X, side=tk.BOTTOM)
        self.financial_panel_page.invoice_frames["Variation"].pack_forget()
        self.financial_panel_page.invoice_frames["Variation"].pack(fill=tk.X, side=tk.BOTTOM)
        for i in range(self.conf["n_items"]-1, -1, -1):
            self.financial_panel_page.bill_dic["Variation"]["Expand"][i].pack_forget()
            self.financial_panel_page.bill_dic["Variation"]["Expand"][i].pack(fill=tk.X, side=tk.BOTTOM)
        self.financial_panel_page.bill_frames["Variation"].pack_forget()
        self.financial_panel_page.bill_frames["Variation"].pack(fill=tk.X, side=tk.BOTTOM)
        for i in range(self.conf["n_items"]-1, -1, -1):
            self.financial_panel_page.profit_dic["Variation"]["Expand"][i].pack_forget()
            self.financial_panel_page.profit_dic["Variation"]["Expand"][i].pack(fill=tk.X, side=tk.BOTTOM)
        self.financial_panel_page.profit_frames["Variation"].pack_forget()
        self.financial_panel_page.profit_frames["Variation"].pack(fill=tk.X, side=tk.BOTTOM)
        # self.update_fee(variation_var)
        # self.data["Invoices"]["Details"]["Variation"]["Expand"].set(True)
        # self.fee_dic["Variation"]["Expand"].grid_forget()
        # tk.Label(self.fee_frames["Variation"], text="", width=10).grid(row=0, column=0)
        # self.fee_dic["Variation"]["Service"].config(text="Variation")
        # self.fee_frames["Variation"].pack(side=tk.BOTTOM)
        # self.financial_panel_page.update_invoice(variation_var)
        # self.financial_panel_page.update_bill(variation_var)
        # self.financial_panel_page.update_profit(variation_var)

    def show_frame(self, page):
        self.project_info_page.pack_forget()
        self.fee_proposal_page.pack_forget()
        # self.fee_accepted_page.pack_forget()
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

    def _minus_update(self, fee, bill, total, *args):
        fee_amount = 0 if len(fee.get().strip()) == 0 else fee.get()
        bill_amount = 0 if len(bill.get().strip()) == 0 else bill.get()
        try:
            total.set(str(round(float(fee_amount)-float(bill_amount))))
        except ValueError:
            total.set("Error")

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
        if len(self.data["Project Info"]["Project"]["Quotation Number"].get()) !=0:
            save(self)
        self.data["Project Info"]["Project"]["Quotation Number"].set(self.load_project_quotation.get().split("-")[0])
        load(self)
        self.load_project_quotation.set("")
        for drop_down in self.state_dict.values():
            drop_down.set("")

    def auto_check(self):
        def running_email_server():
            email_server(self)

        _thread.start_new_thread(running_email_server, ())

        # def check_log_and_save():
        #     while True:
        #         print("auto confined")
        #         time.sleep(10)
        #         # save(self)
        #         config_state(self)
        #         if len(self.data["Project Info"]["Project"]["Quotation Number"].get()) != 0:
        #             try:
        #                 config_log(self)
        #             except:
        #                 continue
        #
        # _thread.start_new_thread(check_log_and_save, ())

    def confirm(self):
        if len(self.data["Project Info"]["Project"]["Quotation Number"].get()) == 0:
            self.destroy()
        else:
            save(self)
            self.destroy()

    def _update_xero(self):
        if not self.data["State"]["Fee Accepted"].get():
            messagebox.showerror("Error", "You haven't update a fee acceptant yet, please update a fee acceptance before you update xero")
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

    def _finish_setup(self):
        if len(self.data["Project Info"]["Project"]["Quotation Number"].get()) == 0:
            messagebox.showerror("Error", "Please Create an quotation Number first")
            return

        finish_setup(self)

    def _rename_project(self):
        if len(self.data["Project Info"]["Project"]["Quotation Number"].get()) == 0:
            messagebox.showerror("Error", "Please Create an quotation Number first")
            return

        old_folder = rename_project(self)
        if rename_asana_project(self, old_folder):
            messagebox.showinfo("Renamed", f"Rename the folder and asana from {old_folder} to {self.data['Project Info']['Project']['Quotation Number'].get()}-{self.data['Project Info']['Project']['Project Name'].get()}")
        else:
            messagebox.showinfo("Renamed", f"Rename the folder from {old_folder} to {self.data['Project Info']['Project']['Quotation Number'].get()}-{self.data['Project Info']['Project']['Project Name'].get()}")
        save(self)
        config_log(self)
        config_state(self)
    def _update_rec(self, *args):
        if self.data["State"]["Fee Accepted"].get():
            self._config_color_code(["green"]*4)
        elif self.data["State"]["Email to Client"].get():
            self._config_color_code(["green", "green", "green", "yellow"])
        elif self.data["State"]["Generate Proposal"].get():
            self._config_color_code(["green", "green", "yellow", "red"])
        elif self.data["State"]["Set Up"].get():
            self._config_color_code(["green", "yellow", "red", "red"])
        else:
            self._config_color_code(["yellow", "red", "red", "red"])

    def _config_color_code(self, l):
        self.set_up_state.config(bg=l[0])
        self.proposal_state.config(bg=l[1])
        self.email_state.config(bg=l[2])
        self.accept_state.config(bg=l[3])

    def open_folder(self):
        quotation_number = self.data["Project Info"]["Project"]["Quotation Number"].get().upper()
        folder_name = self.data["Project Info"]["Project"]["Quotation Number"].get() + "-" + \
                      self.data["Project Info"]["Project"]["Project Name"].get()
        folder_path = os.path.join(self.conf["working_dir"], folder_name)
        if len(quotation_number) == 0:
            messagebox.showerror("Error", "Please enter a Quotation Number before you load")
        elif not os.path.exists(folder_path):
            messagebox.showerror("Error", f"Python cannot find the folder {folder_path}")
        else:
            webbrowser.open(folder_path)
    def open_database(self):
        quotation_number = self.data["Project Info"]["Project"]["Quotation Number"].get().upper()
        database_path = os.path.join(self.conf["database_dir"], quotation_number)
        if len(quotation_number) == 0:
            messagebox.showerror("Error", "Please enter a Quotation Number before you load")
        elif not os.path.exists(database_path):
            messagebox.showerror("Error", f"Python cannot find the folder {database_path}")
        else:
            webbrowser.open(database_path)
