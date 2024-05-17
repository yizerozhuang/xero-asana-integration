from tkinter import ttk

from app_log import AppLog
from project_info_page import ProjectInfoPage
from fee_proposal_page import FeeProposalPage
from financial_panel import FinancialPanelPage
from search_bar_page import SearchBarPage
from utility import *
import tkinter as tk
from asana_function import update_asana, update_asana_invoices
from xero_function import update_xero, refresh_token, login_xero
from email_server import email_server
from app_messagebox import AppMessagebox
from text_extension import TextExtension


from PIL import Image, ImageTk
import _thread
import os
import webbrowser
import time

class App(tk.Tk):
    def __init__(self, conf, user, *args, **kwargs):
        # TkinterDnD.Tk.__init__(self, *args, **kwargs)
        tk.Tk.__init__(self, *args, **kwargs)
        self.title("Premium Consulting Engineers")
        self.conf = conf
        self.user = user
        self.log = AppLog()

        self.current_quotation = tk.StringVar()
        self.current_project_name = tk.StringVar()
        self.current_project_number = tk.StringVar()
        # self.project_number = ""

        self.messagebox = AppMessagebox()
        self.log_text = tk.StringVar()
        logo = Image.open(os.path.join(self.conf["resource_dir"], "jpg", "logo.jpg"))
        render = ImageTk.PhotoImage(logo)
        self.iconphoto(False, render)
        self.state("zoomed")
        self.data = {}

        self.main_frame = tk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=1)

        self.default_set_up()
        self.utility_part()
        self.change_page_part()
        self.main_context_part()

        self.current_page = self.project_info_page

        self.show_frame(self.current_page)

        self.role_check()

        self.protocol("WM_DELETE_WINDOW", self.confirm)

        config_state(self)
        if self.user in self.conf["admin_user_list"] and not self.conf["test_mode"]:
            self.auto_check()

        self.timer_thread()

    def default_set_up(self):
        self.data["Login_user"] = tk.StringVar(value=self.user)
        self.data["Asana_id"] = tk.StringVar()
        self.data["Asana_url"] = tk.StringVar()
        self.data["Current_folder_address"] = tk.StringVar()
        self.data["Lock"] = {
            "Proposal": tk.BooleanVar(),
            "Invoices": tk.BooleanVar()
        }
        self.data["State"] = {
            "Set Up": tk.BooleanVar(),
            "Generate Proposal": tk.BooleanVar(),
            "Email to Client": tk.BooleanVar(),
            "Fee Accepted": tk.BooleanVar(),
            "Quote Unsuccessful": tk.BooleanVar(),
            "Asana State": tk.StringVar(value="Fee Proposal")
        }
        self.data["Email"] = {
            "Fee Coming": tk.StringVar(),
            "Fee Proposal": tk.StringVar(),
            "First Chase": tk.StringVar(),
            "Second Chase": tk.StringVar(),
            "Third Chase": tk.StringVar(),
            "Fee Accepted": tk.StringVar()
        }
        self.data["Email_Content"] = tk.StringVar()
        self.data["Address_to"] = tk.StringVar(value="Client")

    def utility_part(self):

        # Utility Part
        self.utility_frame = tk.LabelFrame(self.main_frame, text="Utility", font=self.conf["font"])
        self.utility_frame.pack(side=tk.TOP)
        
        state_frame = tk.LabelFrame(self.utility_frame, font=self.conf["font"])
        state_frame.grid(row=0, column=1)

        tk.Button(state_frame, text="Set Up", command=self._finish_setup, bg="cyan", font=self.conf["font"]).grid(row=0, column=0)
        self.preview_fee_proposal_button = tk.Button(state_frame, text="Gen Fee Proposal", bg="cyan", command=self._preview_fee_proposal,
                                                     font=self.conf["font"])
        self.preview_fee_proposal_button.grid(row=0, column=1)
        self.email_to_client_button = tk.Button(state_frame, text="Email Fee Proposal", command=self._email_fee_proposal, bg="cyan",
                                                font=self.conf["font"])
        self.email_to_client_button.grid(row=0, column=2)
        self.chase_client_button = tk.Button(state_frame, text="Chase Fee Acceptance", command=lambda: chase(self), bg="cyan",
                                             font=self.conf["font"])
        self.chase_client_button.grid(row=0, column=3)
        
        function_frame = tk.LabelFrame(self.utility_frame, font=self.conf["font"])
        function_frame.grid(row=0, column=2, sticky="ns")

        tk.Button(function_frame, width=10, text="Open Folder", command=self.open_folder, bg="brown",
                  fg="white",
                  font=self.conf["font"]).grid(row=0, column=0)

        self.open_database_button = tk.Button(function_frame, width=12, text="Open Database",
                                              command=self.open_database, bg="brown",
                                              fg="white",
                                              font=self.conf["font"])
        self.open_database_button.grid(row=1, column=0)


        self.design_certificate = tk.Button(function_frame, width=14, text="Design Certificate", bg="brown", fg="white",
                                            command=lambda: open_design_certificate(self),
                                            font=self.conf["font"])
        self.design_certificate.grid(row=2, column=0)


        tk.Button(function_frame, text="Rename Project", command=self._rename_project, bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=1)

        tk.Button(function_frame, text="Update Asana", command=self._update_asana, bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=2)

        tk.Button(function_frame, text="Open Asana", command=self._open_asana, bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=1, column=2)

        self.update_xero_button = tk.Button(function_frame, text="Update Xero", command=self._update_xero, bg="brown", fg="white",
                                            font=self.conf["font"])
        self.update_xero_button.grid(row=0, column=3)

        self.refresh_xero_button = tk.Button(function_frame, text="Refresh Xero Token", command=refresh_token, bg="brown", fg="white",
                                             font=self.conf["font"])
        self.refresh_xero_button.grid(row=1, column=3)

        self.login_xero_button = tk.Button(function_frame, text="Login Xero", command=login_xero, bg="brown", fg="white",
                                           font=self.conf["font"])
        self.login_xero_button.grid(row=2, column=3)

        # tk.Button(function_frame, text="Refresh Token", command=refresh_token, bg="brown", fg="white",
        #           font=self.conf["font"]).grid(row=2, column=3)

        tk.Label(state_frame, text="To Set Up").grid(row=1, column=0)
        tk.Label(state_frame, text="To Gen Fee").grid(row=1, column=1)
        tk.Label(state_frame, text="To Email").grid(row=1, column=2)
        tk.Label(state_frame, text="To Win").grid(row=1, column=3)

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

        self.data["State"]["Set Up"].trace("w", self._update_project_state)
        self.data["State"]["Generate Proposal"].trace("w", self._update_project_state)
        self.data["State"]["Email to Client"].trace("w", self._update_project_state)
        self.data["State"]["Fee Accepted"].trace("w", self._update_project_state)
        self.data["State"]["Quote Unsuccessful"].trace("w", self._update_project_state)

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


        self.legend_frame = tk.LabelFrame(self.utility_frame)
        self.legend_frame.grid(row=0, column=3)

        tk.Label(self.legend_frame, text="Invoice States: ").grid(row=0, column=0)

        _ = tk.LabelFrame(self.legend_frame)
        _.grid(row=0, column=1, sticky="ew")
        tk.Label(_, text="Backlog").pack()

        tk.Label(self.legend_frame, text="Sent", bg="red").grid(row=0, column=2, sticky="ew")
        tk.Label(self.legend_frame, text="Paid", bg="green").grid(row=0, column=4, sticky="ew")
        tk.Label(self.legend_frame, text="Voided", bg="purple", fg="white").grid(row=0, column=5, sticky="ew")

        tk.Label(self.legend_frame, text="Bill States: ").grid(row=1, column=0)

        _ = tk.LabelFrame(self.legend_frame)
        _.grid(row=1, column=1, sticky="ew")
        tk.Label(_, text="Draft").pack()

        tk.Label(self.legend_frame, text="Awaiting Approval", bg="red").grid(row=1, column=2, sticky="ew")
        tk.Label(self.legend_frame, text="Awaiting Payment", bg="orange").grid(row=1, column=3, sticky="ew")
        tk.Label(self.legend_frame, text="Paid", bg="green").grid(row=1, column=4, sticky="ew")
        tk.Label(self.legend_frame, text="Voided", bg="purple", fg="white").grid(row=1, column=5, sticky="ew")

    def main_context_part(self):
        # main frame page
        self.project_info_page = ProjectInfoPage(self.main_frame, self)
        self.search_bar_page = SearchBarPage(self.main_frame, self)
        self.fee_proposal_page = FeeProposalPage(self.main_frame, self)
        self.financial_panel_page = FinancialPanelPage(self.main_frame, self)
        self._update_variation()
        self._project_number_page()

    def change_page_part(self):
        # change page
        change_page_frame = tk.LabelFrame(self.main_frame, text="Change Page")
        change_page_frame.pack(side=tk.BOTTOM)

        tk.Button(change_page_frame, text="Search Bar",
                  command=lambda: self.show_frame(self.search_bar_page), bg="DarkOrange1", fg="white",
                  font=self.conf["font"]).grid(row=0, column=0)
        tk.Button(change_page_frame, text="Project Info",
                  command=lambda: self.show_frame(self.project_info_page), bg="DarkOrange1", fg="white",
                  font=self.conf["font"]).grid(row=0, column=1)
        tk.Button(change_page_frame, text="Fee Details",
                  command=lambda: self.show_frame(self.fee_proposal_page), bg="DarkOrange1", fg="white",
                  font=self.conf["font"]).grid(row=0, column=2)
        self.finacial_panel_button = tk.Button(change_page_frame, text="Financial Panel",
                                               command=lambda: self.show_frame(self.financial_panel_page),
                                               bg="DarkOrange1",
                                               fg="white",
                                               font=self.conf["font"])
        self.finacial_panel_button.grid(row=0, column=3)

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

        # variation_var = {
        #     "Service": tk.StringVar(value="Variation"),
        #     "Include": tk.BooleanVar(value=True)
        # }

        self.fee_proposal_page.update_fee("Variation", True)
        # self.fee_proposal_page.fee_dic["Variation"]["Expand"].grid_forget()
        self.fee_proposal_page.fee_dic["Variation"]["Service"].config(text="Variation", bg="cyan")
        self.fee_proposal_page.fee_frames["Variation"].pack(side=tk.BOTTOM)

        self.financial_panel_page.update_invoice("Variation", True)
        self.financial_panel_page.invoice_dic["Variation"]["Service"].config(text="Variation")

        self.financial_panel_page.update_bill("Variation", True)
        self.financial_panel_page.update_profit("Variation", True)

        # self.data["Invoices"]["Details"]["Variation"]["Expand"].set(True)

        for i in range(self.conf["n_bills"], -1, -1):
            self.financial_panel_page.invoice_dic["Variation"]["Expand"][i].pack_forget()
            self.financial_panel_page.invoice_dic["Variation"]["Expand"][i].pack(fill=tk.X, side=tk.BOTTOM)
        self.financial_panel_page.invoice_frames["Variation"].pack_forget()
        self.financial_panel_page.invoice_frames["Variation"].pack(fill=tk.X, side=tk.BOTTOM)

        for i in range(self.conf["n_bills"], -1, -1):
            self.financial_panel_page.bill_dic["Variation"]["Expand"][i].pack_forget()
            self.financial_panel_page.bill_dic["Variation"]["Expand"][i].pack(fill=tk.X, side=tk.BOTTOM)
        self.financial_panel_page.bill_frames["Variation"].pack_forget()
        self.financial_panel_page.bill_frames["Variation"].pack(fill=tk.X, side=tk.BOTTOM)

        for i in range(self.conf["n_bills"], -1, -1):
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

    def utility_search(self, *args):
        data = self.search_bar_page.check(None, self.utility_search_string_var.get().strip())
        if len(set(data)) == 1:
            load_data(self, data[0][0])
            self.show_frame(self.project_info_page)
        else:
            self.search_bar_page.entry.delete(0, "end")
            self.search_bar_page.entry.insert(0, self.utility_search_string_var.get().strip())
            self.search_bar_page.check(None)
            self.search_bar_page.refresh()
            self.show_frame(self.search_bar_page)
        self.utility_search_string_var.set("")



    def utility_reset(self):
        reset(self)
        self.utility_search_string_var.set("")

    def _project_number_page(self):
        project_number_frame = tk.LabelFrame(self.utility_frame)
        project_number_frame.grid(row=0, column=0, sticky="ns")
        tk.Label(project_number_frame, text="Project Number: ", font=self.conf["font"]).grid(row=0, column=0)
        tk.Label(project_number_frame, text="Quotation Number: ", font=self.conf["font"]).grid(row=1, column=0)
        tk.Label(project_number_frame, text="Project Name: ", font=self.conf["font"]).grid(row=2, column=0)


        tk.Label(project_number_frame, textvariable=self.data["Project Info"]["Project"]["Project Number"], font=self.conf["font"]).grid(row=0, column=1)
        self.quotation_number_label = tk.Label(project_number_frame, textvariable=self.current_quotation, font=self.conf["font"])
        self.quotation_number_label.grid(row=1, column=1)
        # self.data["Project Info"][]
        # self.current_quotation.trace("w", self._update_quotation_number_label)
        tk.Label(project_number_frame, textvariable=self.data["Project Info"]["Project"]["Project Name"], font=self.conf["font"]).grid(row=2, column=1)
        # tk.Entry(project_number_frame, textvariable=self.data["Project Info"]["Project"]["Project Name"], width=30, font=self.conf["font"], state=tk.DISABLED).grid(row=2, column=1)
        search_frame = tk.Frame(project_number_frame)
        search_frame.grid(row=3, column=0, columnspan=2)
        self.utility_search_string_var = tk.StringVar()
        tk.Entry(search_frame, fg="blue", font=self.conf["font"], textvariable=self.utility_search_string_var).grid(row=0, column=0)
        tk.Button(search_frame, text="Clear Up", command=self.utility_reset, bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=1)
        tk.Button(search_frame, text="Search", command=self.utility_search, bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=2)
        self.bind("<Control-Key-space>", self.utility_search)
        # self.project_number_label = tk.Label(project_number_frame, font=self.conf["font"])
        # self.project_number_label.grid(row=0, column=1)
        # self.data["Project Info"]["Project"]["Quotation Number"].trace("w",self._update_project_number_label)
        # self.data["Project Info"]["Project"]["Project Number"].trace("w",self._update_project_number_label)
        #
        # # tk.Label(project_number_frame, textvariable=self.data["Project Info"]["Project"]["Quotation Number"], font=self.conf["font"]).grid(row=0, column=1)
        #
        # tk.Label(project_number_frame, textvariable=self.data["Project Info"]["Project"]["Project Name"], font=self.conf["font"]).grid(row=1, column=1)

    def role_check(self):
        if not self.user in conf["admin_user_list"]:
            # self.project_info_page.client_frame.grid_forget()
            # self.project_info_page.contact_frame.grid_forget()
            self.project_info_page.finish_frame.grid_forget()
            self.fee_proposal_page.preview_installation_proposal_button.grid_forget()
            self.fee_proposal_page.email_installation_proposal_buttion.grid_forget()
            self.fee_proposal_page.fee_frame.grid_forget()
            self.finacial_panel_button.grid_forget()
            self.preview_fee_proposal_button.grid_forget()
            self.email_to_client_button.grid_forget()
            self.chase_client_button.grid_forget()
            self.open_database_button.grid_forget()
            self.design_certificate.grid_forget()
            self.update_xero_button.grid_forget()
            self.refresh_xero_button.grid_forget()
            self.login_xero_button.grid_forget()
            self.legend_frame.grid_forget()


    # def _update_quotation_number_label(self, *args):
    #     if len(self.data["Project Info"]["Project"]["Project Number"].get()) == 0:
    #         self.quotation_number_label.config(text=self.data["Project Info"]["Project"]["Quotation Number"].get())
    #     else:
    #         self.quotation_number_label.config(text=self.data["Project Info"]["Project"]["Project Number"].get())

    def show_frame(self, page):
        self.search_bar_page.pack_forget()
        self.project_info_page.pack_forget()
        self.fee_proposal_page.pack_forget()
        # self.fee_accepted_page.pack_forget()
        self.financial_panel_page.pack_forget()

        if type(page) == SearchBarPage:
            self.search_bar_page.refresh()

        page.pack(fill=tk.BOTH, expand=1)

    def _finish_setup(self):
        if len(self.data["Project Info"]["Project"]["Quotation Number"].get()) == 0:
            self.messagebox.show_error("Please Create an quotation Number first")
            # messagebox.showerror("Error", "\n Please Create an quotation Number first\n  \n")
            return
        if self._validate_project_name():
            try:
                finish_setup(self)
            except PermissionError as e:
                print(e)
                self.messagebox.show_error("Permission Denied, Please Close the file before you finish set up")

    def _validate_project_name(self):
        project_name = self.data["Project Info"]["Project"]["Project Name"].get()
        if project_name[0] == " " or project_name[-1] == " ":
            self.messagebox.show_error("You cant have empty space in front or end of the project name")
            return False
        special_char_list = ["\\", "/", ":", "*", "?", '"', "<",">","|"]
        for char in project_name:
            if char in special_char_list:
                self.messagebox.show_error("You cant have special character in your project name")
                return False
        return True


    def _preview_fee_proposal(self):
        preview_fee_proposal(self)

    def _ist_update(self, fee, ingst, no_gst=None):
        if not no_gst is None and no_gst.get():
            ingst.set(fee.get())
            return

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
        # if len(self.data["Project Info"]["Project"]["Quotation Number"].get()) !=0:
        #     save(self)

        # load(self, self.load_project_quotation.get().split("-")[0])
        load_quotation = self.load_project_quotation.get().split("-")[0]
        # self.data["Project Info"]["Project"]["Quotation Number"].set(load_quotation)

        load_data(self, load_quotation)

        self.load_project_quotation.set("")

        for drop_down in self.state_dict.values():
            drop_down.set("")

    def auto_check(self):
        def running_email_server():
            email_server(self)

        _thread.start_new_thread(running_email_server, ())

    def timer_thread(self):
        self.timer = self.conf["timer"]
        def project_timer():
            while True:
                time.sleep(1)
                if len(self.current_quotation.get())!=0:
                    self.timer -=1
                    if self.timer == 0:
                        reset(self)
                        self.messagebox.show_error(f"The Session is exceeding {str(int(self.conf['timer']/60))} mins, please login this project again")
                else:
                    self.timer = self.conf["timer"]

        _thread.start_new_thread(project_timer, ())


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
            try:
                if len(self.current_quotation.get())!=0:
                    self.data["Project Info"]["Project"]["Quotation Number"].set(self.current_quotation.get())
                self.data["Login_user"].set("")
                save(self)
            except Exception as e:
                print(e)
            self.destroy()

    def _update_asana(self):
        # update = change_type_of_calculator(self)
        # if update is None:
        #     return
        # if update:
        try:
            update_asana(self)
            update_asana_invoices(self)
        except Exception as e:
            print(e)
            self.messagebox.show_error("Can not find Asana id")
            return
        self.messagebox.show_update("Successful Update Asana")
        # else:
        #     self.messagebox.show_error("Please Close the calculation excel sheet before you Update Asana")


    def _open_asana(self):
        if len(self.data["Asana_id"].get()) == 0:
            self.messagebox.show_error("You should update Asana first before you open asana")
            return
        elif len(self.data["Asana_url"].get()) == 0:
            self.messagebox.show_error("Can not find the asana link, please contact Admin")
            return
        edge_address = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
        def open_asana():
            subprocess.call([edge_address, self.data["Asana_url"].get()+"/f"])

        _thread.start_new_thread(open_asana, ())
        # webbrowser.open(self.data["Asana_url"].get())
    # def _open_asana(self):
    #     if len(self.data["Asana_id"]) !=0:
    #         self.messagebox.show_error("Please Update Asana Before you Open")
    #         return

    def _update_xero(self):
        if not self.data["State"]["Fee Accepted"].get():
            self.messagebox.show_error("Please upload a fee acceptance before you update xero")
            return
        elif len(self.data["Project Info"]["Project"]["Project Number"].get())==0:
            self.messagebox.show_error("Cant Find the Project Number, Please sent first invoice to the client.")
            return

        # true = False
        refresh_token()
        true=update_xero(self)
        # self.messagebox.show_error("You Haven't login xero yet")
        if true:
            self.messagebox.show_update("Update Xero and Asana Successful")
        else:
            self.messagebox.show_error("Can not update Xero and Asana")

    def _rename_project(self):
        current_folder_address = os.path.join(self.conf["working_dir"], self.data["Current_folder_address"].get())
        if len(self.data["Project Info"]["Project"]["Project Number"].get()) != 0:
            folder_name = self.data["Project Info"]["Project"]["Project Number"].get() + "-" + self.data["Project Info"]["Project"]["Project Name"].get()
        else:
            folder_name = self.data["Project Info"]["Project"]["Quotation Number"].get() + "-" + self.data["Project Info"]["Project"]["Project Name"].get()
        folder_address = os.path.join(self.conf["working_dir"], folder_name)


        if len(self.data["Project Info"]["Project"]["Quotation Number"].get()) == 0:
            self.messagebox.show_error("Please Create an Quotation Number first")
            return
        elif os.path.exists(folder_address):
            self.messagebox.show_error(f"{folder_address} exists, you dont neet to rename it")
            return

        if self._validate_project_name():
            try:
                os.rename(current_folder_address, folder_address)
                # old_dir, new_dir = rename_project(self)
            except Exception as e:
                print(e)
                self.messagebox.show_error("Bridge wont able to rename the folder, Please close anything related to the folder")
                return

            if len(self.data["Asana_id"].get()) == 0:
                self.messagebox.file_info("Rename Folder", current_folder_address, folder_address)
            else:
                update_asana(self)
                self.messagebox.file_info("Rename Folder and Asana", current_folder_address, folder_address)
            self.data["Current_folder_address"].set(folder_name)
            webbrowser.open(folder_address)
            save(self)
            config_log(self)
            config_state(self)

    def _email_fee_proposal(self):
        try:
            def email():
                email_fee_proposal(self)
            _thread.start_new_thread(email, ())
            update = self.messagebox.ask_yes_no("Do you want to update Asana?")
            if update:
                update_asana(self)

        except Exception as e:
            self.messagebox.show_error("Unable to Create Email")
            return

    def _update_project_state(self, *args):
        if self.data["State"]["Fee Accepted"].get() or self.data["State"]["Quote Unsuccessful"].get():
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
        current_folder_address = os.path.join(self.conf["working_dir"], self.data["Current_folder_address"].get())
        # if len(self.data["Project Info"]["Project"]["Project Number"].get()) ==0:
        #     folder_name = self.data["Project Info"]["Project"]["Quotation Number"].get() + "-" + \
        #                   self.data["Project Info"]["Project"]["Project Name"].get()
        # else:
        #     folder_name = self.data["Project Info"]["Project"]["Project Number"].get() + "-" + \
        #                   self.data["Project Info"]["Project"]["Project Name"].get()
        # folder_path = os.path.join(self.conf["working_dir"], folder_name)

        if len(quotation_number) == 0:
            self.messagebox.show_error("Please login to a project first")
        elif not os.path.exists(current_folder_address):
            self.messagebox.show_error(f"Python cannot find the folder {current_folder_address}")
        else:
            webbrowser.open(current_folder_address)

    def open_database(self):
        quotation_number = self.data["Project Info"]["Project"]["Quotation Number"].get().upper()
        # database_path = os.path.join(self.conf["database_dir"], quotation_number)
        accounting_dir = os.path.join(self.conf["accounting_dir"], quotation_number)
        if len(quotation_number) == 0:
            self.messagebox.show_error("Please enter a Quotation Number before you load")
        elif not os.path.exists(accounting_dir):
            self.messagebox.show_error(f"Python cannot find the folder {accounting_dir}")
        else:
            webbrowser.open(accounting_dir)

    def _config_frame(self, tk_object, state):
        if type(tk_object)==tk.Frame or type(tk_object)==tk.LabelFrame or type(tk_object)==TextExtension:
            for child in tk_object.winfo_children():
                self._config_frame(child, state)
        else:
            tk_object.config(state=state)