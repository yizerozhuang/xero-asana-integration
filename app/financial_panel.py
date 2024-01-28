import tkinter as tk
from tkinter import ttk, filedialog

from utility import *
from asana_function import update_asana, update_asana_invoices

import os
import json
from datetime import date
import shutil

class FinancialPanelPage(tk.Frame):
    def __init__(self, parent, app):
        tk.Frame.__init__(self, parent)
        self.update_invoice_number = None
        self.parent = parent
        self.app = app
        self.data = app.data
        self.conf = app.conf
        self.messagebox = app.messagebox
        
        self.main_frame = tk.Frame(self, bg="blue")
        self.main_frame.pack(fill=tk.BOTH, expand=1, side=tk.LEFT, anchor="nw")

        self.main_canvas = tk.Canvas(self.main_frame)
        self.main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1, anchor="nw")

        self.scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.main_canvas.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.main_canvas.config(yscrollcommand=self.scrollbar.set)
        self.main_canvas.bind("<Configure>", lambda e:self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")))
        self.app.bind("<MouseWheel>", self._on_mousewheel, add="+")
        self.main_context_frame = tk.LabelFrame(self.main_canvas, text="Main Context")
        self.main_canvas.create_window((0, 0), window=self.main_context_frame, anchor="nw")
        self.main_context_frame.bind("<Configure>", self.reset_scrollregion)

        self.invoice_part()
        self.invoice_function_part()
        self.bill_part()
        self.profit_part()
        self.fee_acceptance_part()

    def invoice_part(self):

        invoice = []
        self.data["Invoices Number"] = invoice

        self.invoice_dic = dict()
        self.invoice_frames= dict()

        self.invoice_frame = tk.LabelFrame(self.main_context_frame, text="Invoice Details", font=self.conf["font"])
        self.invoice_frame.grid(row=0, column=0)
        self.invoice_label_list = []
        title_frame = tk.LabelFrame(self.invoice_frame)
        title_frame.pack(fill=tk.X)
        tk.Label(title_frame, width=35, text="Items", font=self.conf["font"]).grid(row=0, column=0)
        tk.Label(title_frame, width=15, text="ex.GST/in.GST", font=self.conf["font"]).grid(row=0, column=1)
        invoice_number_frame = tk.LabelFrame(self.invoice_frame)
        invoice_number_frame.pack(fill=tk.X)
        tk.Label(invoice_number_frame, width=35, text="Invoice Number", font=self.conf["font"]).grid(row=0, column=0)
        _ = tk.Frame(invoice_number_frame, width=15)
        _.grid(row=0, column=1)
        tk.Button(_, width=7, text="Unlock", command=self.unlock, bg="brown", fg="white", font=self.conf["font"]).grid(row=0, column=0)
        tk.Button(_, width=7, text="Gen", command=self.generate_invoice_number, bg="brown", fg="white", font=self.conf["font"]).grid(row=0, column=1)

        for i in range(6):
            invoice.append(
                {
                    "Number": tk.StringVar(),
                    "Fee": tk.StringVar(),
                    "in.GST": tk.StringVar(),
                    "State": tk.StringVar(value="Backlog"),
                    "Asana_id": tk.StringVar(),
                    "Xero_id": tk.StringVar()
                }
            )
            tk.Label(title_frame, width=8, text="INV"+str(i+1), font=self.conf["font"]).grid(row=0, column=i+2, sticky="ew", padx=(2,0))
            self.invoice_label_list.append([])
            self.invoice_label_list[i].append(tk.Entry(invoice_number_frame, state=tk.DISABLED, textvariable=invoice[i]["Number"], width=8, font=self.conf["font"], fg="blue"))
            self.invoice_label_list[i][0].grid(row=0, column=i + 2, padx=(10, 0), sticky="ew")

        total_frame = tk.LabelFrame(self.invoice_frame)
        total_frame.pack(side=tk.BOTTOM, fill=tk.X)
        tk.Label(total_frame, width=35, text="Invoice Total", font=self.conf["font"]).grid(row=0, column=0)
        # self.total_label = tk.Label(total_frame, width=15, text="$0", font=self.conf["font"])
        tk.Label(total_frame, width=15, textvariable=self.data["Invoices"]["Fee"], font=self.conf["font"]).grid(row=0, column=1)

        tk.Label(total_frame, width=15, textvariable=self.data["Invoices"]["in.GST"], font=self.conf["font"]).grid(row=1, column=1)

        ist_fuc = lambda i: lambda a, b, c: self.app._ist_update(invoice[i]["Fee"], invoice[i]["in.GST"])

        for i in range(6):
            self.invoice_label_list[i].append(tk.Label(total_frame, textvariable=invoice[i]["Fee"], width=8, font=self.conf["font"]))
            self.invoice_label_list[i][1].grid(row=0, column=2 + i, padx=(2, 0))
            self.invoice_label_list[i].append(tk.Label(total_frame, textvariable=invoice[i]["in.GST"], width=8, font=self.conf["font"]))
            self.invoice_label_list[i][2].grid(row=1, column=2 + i, padx=(2, 0))
            invoice[i]["Fee"].trace("w", ist_fuc(i))

        # details = self.data["Invoices"]["Details"]
        # invoice_details = self.data["Invoices Number"]

        # self.variation_label_list = []
        # update_label_fuc = lambda i: lambda a, b, c: self.update_fee_label(self.data["Variation"][i]["Fee"], self.variation_label_list[i])
        # for i in range(self.conf["n_variation"]):
        #     variation_frame = tk.LabelFrame(self.invoice_frame)
        #     variation_frame.pack(side=tk.BOTTOM, fill=tk.X)
        #     tk.Label(variation_frame, textvariable=self.data["Variation"][i]["Service"], width=35, font=self.conf["font"]).grid(row=0, column=0)
        #     self.variation_label_list.append(tk.Label(variation_frame, text="$0/$0.0", width=15, font=self.conf["font"]))
        #     self.variation_label_list[i].grid(row=0, column=1)
        #     self.data["Variation"][i]["Fee"].trace("w", update_label_fuc(i))
        #     for j in range(6):
        #         tk.Radiobutton(variation_frame, width=6, variable=self.data["Variation"][i]["Number"], value="INV" + str(j + 1)).grid(row=0, column=j+2, padx=(2,0))
        #     self.data["Variation"][i]["Fee"].trace("w", lambda a, b, c: self.update_invoice_sum(details, invoice_details))
        #     self.data["Variation"][i]["Number"].trace("w", lambda a, b, c: self.update_invoice_sum(details, invoice_details))
    def update_invoice(self, var):
        details = self.data["Invoices"]["Details"]
        invoice_details = self.data["Invoices Number"]
        service = var["Service"].get()
        include = var["Include"].get()
        if include:
            if not service in self.invoice_frames:
            # if not service in self.invoice_dic:
                self.invoice_frames[service] = tk.LabelFrame(self.invoice_frame)
                self.invoice_dic[service] = {
                    "Service": tk.Label(self.invoice_frames[service],
                                        text=service + " design and documentation",
                                        width=35,
                                        font=self.conf["font"]),
                    # "Fee": tk.Label(self.invoice_frames[service],
                    #                 text="$0/$0.0",
                    #                 width=15,
                    #                 font=self.conf["font"]),
                    "Fee": tk.Frame(self.invoice_frames[service]),
                    "Invoice": [tk.Radiobutton(self.invoice_frames[service],
                                               width=6,
                                               variable=self.data["Invoices"]["Details"][service]["Number"],
                                               value="INV" + str(i + 1)) for i in range(6)],
                    "Expand": [tk.LabelFrame(self.invoice_frame, height=30) for _ in range(self.conf["n_items"])]
                }
                tk.Label(self.invoice_dic[service]["Fee"], text="$").grid(row=0, column=0)
                tk.Label(self.invoice_dic[service]["Fee"], textvariable=details[service]["Fee"], width=6).grid(row=0, column=1)
                tk.Label(self.invoice_dic[service]["Fee"], text="/").grid(row=0, column=2)
                tk.Label(self.invoice_dic[service]["Fee"], text="$").grid(row=0, column=3)
                tk.Label(self.invoice_dic[service]["Fee"], textvariable=details[service]["in.GST"], width=6).grid(row=0, column=4)

                self.invoice_dic[service]["Content"] = [
                    {
                        "Service": tk.Label(self.invoice_dic[service]["Expand"][i],
                                            width=35,
                                            font=self.conf["font"],
                                            textvariable=details[service]["Content"][i]["Service"]),
                        "Fee": tk.Frame(self.invoice_dic[service]["Expand"][i],
                                        width=15),
                        "Invoice": [tk.Radiobutton(self.invoice_dic[service]["Expand"][i],
                                                   width=6,
                                                   variable=
                                                   self.app.data["Invoices"]["Details"][service]["Content"][i]["Number"],
                                                   value="INV" + str(j + 1)) for j in range(6)]
                    } for i in range(self.conf["n_items"])
                ]
                self.invoice_dic[service]["Service"].grid(row=0, column=0)
                self.invoice_dic[service]["Fee"].grid(row=0, column=1)

                for i in range(self.conf["n_items"]):
                    tk.Label(self.invoice_dic[service]["Content"][i]["Fee"], text="$").grid(row=0, column=0)
                    tk.Label(self.invoice_dic[service]["Content"][i]["Fee"], textvariable=details[service]["Content"][i]["Fee"], width=6).grid(row=0,
                                                                                                                   column=1)
                    tk.Label(self.invoice_dic[service]["Content"][i]["Fee"], text="/").grid(row=0, column=2)
                    tk.Label(self.invoice_dic[service]["Content"][i]["Fee"], text="$").grid(row=0, column=3)
                    tk.Label(self.invoice_dic[service]["Content"][i]["Fee"], textvariable=details[service]["Content"][i]["in.GST"], width=6).grid(row=0,
                                                                                                                      column=4)

                for i in range(6):
                    self.invoice_dic[service]["Invoice"][i].grid(row=0, column=2 + i, padx=(2, 0))


                details[service]["Expand"].trace("w", lambda a, b, c: self._expand_invoice(service))

                details[service]["Fee"].trace("w", lambda a, b, c: self.update_invoice_sum(invoice_details))
                # details[service]["Fee"].trace("w", lambda a, b, c: self.update_fee_label(details[service]["Fee"], self.invoice_dic[service]["Fee"]))
                details[service]["Number"].trace("w", lambda a, b, c: self.update_invoice_sum(invoice_details))
                # update_label_fuc = lambda i: lambda a, b, c: self.update_fee_label(details[service]["Content"][i]["Fee"], self.invoice_dic[service]["Content"][i]["Fee"])
                for i in range(self.conf["n_items"]):
                    details[service]["Content"][i]["Fee"].trace("w", lambda a, b, c: self.update_invoice_sum(invoice_details))
                    # details[service]["Content"][i]["Fee"].trace("w", update_label_fuc(i))

                    details[service]["Content"][i]["Number"].trace("w", lambda a, b, c: self.update_invoice_sum(invoice_details))
            self.invoice_frames[service].pack(fill=tk.X)
            for i in range(self.conf["n_items"]):
                self.invoice_dic[service]["Expand"][i].pack(fill=tk.X)
        else:
            self.invoice_frames[service].pack_forget()
            for i in range(self.conf["n_items"]):
                self.invoice_dic[service]["Expand"][i].pack_forget()
    def _expand_invoice(self, service):
        details = self.data["Invoices"]["Details"]
        if details[service]["Expand"].get():
            for i in range(self.conf["n_items"]):
                self.invoice_dic[service]["Content"][i]["Service"].grid(row=0, column=0, pady=1)
                self.invoice_dic[service]["Content"][i]["Fee"].grid(row=0, column=1)
                for j in range(6):
                    self.invoice_dic[service]["Content"][i]["Invoice"][j].grid(row=0, column=2+j, padx=(2, 0))
            for i in range(6):
                self.invoice_dic[service]["Invoice"][i].grid_forget()
            # self.invoice_dic[service]["Fee"].config(font=self.conf["font"]+["bold"])
        else:
            for i in range(self.conf["n_items"]):
                self.invoice_dic[service]["Content"][i]["Service"].grid_forget()
                self.invoice_dic[service]["Content"][i]["Fee"].grid_forget()
                for j in range(6):
                    self.invoice_dic[service]["Content"][i]["Invoice"][j].grid_forget()

            # self.invoice_dic[service]["Fee"].config(font=self.conf["font"])
            for i in range(6):
                self.invoice_dic[service]["Invoice"][i].grid(row=0, column=2+i)
        details[service]["Number"].set("None")
        for i in range(self.conf["n_items"]):
            details[service]["Content"][i]["Number"].set("None")
    def invoice_function_part(self):
        remittance = [
            {
                "Type": tk.StringVar(),
                "Preview_Upload": tk.BooleanVar(),
                "Email_Upload": tk.BooleanVar(),
                "Full_Upload": tk.BooleanVar(),
                "Part1": tk.StringVar(),
                "Part1_Upload": tk.BooleanVar(),
                "Part2": tk.StringVar(),
                "Part2_Upload": tk.BooleanVar(),
                "Part3": tk.StringVar(),
                "Part3_Upload": tk.BooleanVar(),
            } for _ in range(6)
        ]
        self.data["Remittances"] = remittance

        invoice_function_frame = tk.LabelFrame(self.main_context_frame, text="Invoice Functions", font=self.conf["font"])
        invoice_function_frame.grid(row=1, column=0, sticky="ew")
        print_invoice_function = lambda i: lambda: self._excel_print_invoice(i)
        email_invoice_function = lambda i: lambda: self._email_invoice(i)
        upload_remittance_func = lambda part, i: lambda: self.upload_remittance(invoice_number=self.data["Invoices Number"][i]["Number"].get(), part=part, i=i)
        button_config_func = lambda var, button: lambda a,b,c: self.config_button(var, button)

        for i in range(6):
            tk.Label(invoice_function_frame, text="Invoice Number: ", font=self.conf["font"]).grid(row=i, column=0)
            self.invoice_label_list[i].append(tk.Label(invoice_function_frame, width=10, textvariable=self.data["Invoices Number"][i]["Number"], font=self.conf["font"]))
            self.invoice_label_list[i][3].grid(row=i, column=1)
            self.data["Invoices Number"][i]["State"].trace("w", self.invoice_color_code)
            preview_button = tk.Button(invoice_function_frame, text="Preview", font=self.conf["font"], bg="cyan",
                                       command=print_invoice_function(i))
            preview_button.grid(row=i, column=2)
            remittance[i]["Preview_Upload"].trace("w", button_config_func(remittance[i]["Preview_Upload"], preview_button))

            email_button = tk.Button(invoice_function_frame, text="Email", font=self.conf["font"], bg="cyan",
                                     command=email_invoice_function(i))
            email_button.grid(row=i, column=3)
            remittance[i]["Email_Upload"].trace("w", button_config_func(remittance[i]["Email_Upload"], email_button))

            full_button = tk.Button(invoice_function_frame, text="REMIT FULL", font=self.conf["font"], bg="cyan",
                                    command=upload_remittance_func("Full", i))
            full_button.grid(row=i, column=4)

            remittance[i]["Full_Upload"].trace("w", button_config_func(remittance[i]["Full_Upload"], full_button))

            tk.Entry(invoice_function_frame, width=10, font=self.conf["font"], fg="blue",
                     textvariable=remittance[i]["Part1"]).grid(row=i, column=5)
            button1 = tk.Button(invoice_function_frame, text="REMIT P1", bg="cyan", font=self.conf["font"],
                                command=upload_remittance_func("Part1", i))
            button1.grid(row=i, column=6)

            remittance[i]["Part1_Upload"].trace("w", button_config_func(remittance[i]["Part1_Upload"], button1))

            tk.Entry(invoice_function_frame, width=10, font=self.conf["font"], fg="blue",
                     textvariable=remittance[i]["Part2"]).grid(row=i, column=7)
            button2 = tk.Button(invoice_function_frame, text="REMIT P2", bg="cyan", font=self.conf["font"],
                                command=upload_remittance_func("Part2", i))
            button2.grid(row=i, column=8)

            remittance[i]["Part2_Upload"].trace("w", button_config_func(remittance[i]["Part2_Upload"], button2))

            tk.Entry(invoice_function_frame, width=10, font=self.conf["font"], fg="blue",
                     textvariable=remittance[i]["Part3"]).grid(row=i, column=9)
            button3 = tk.Button(invoice_function_frame, text="REMIT P3", bg="cyan", font=self.conf["font"],
                                command=upload_remittance_func("Part3", i))
            button3.grid(row=i, column=10)
            remittance[i]["Part3_Upload"].trace("w", button_config_func(remittance[i]["Part3_Upload"], button3))

    def bill_part(self):
        bills = {
            "Details": dict(),
            "Fee": tk.StringVar(),
            "in.GST": tk.StringVar()
        }
        self.data["Bills"] = bills

        for service in self.conf["invoice_list"]:
            bills["Details"][service] = {
                    "Include": tk.BooleanVar(),
                    "Service": tk.StringVar(),
                    "Fee": tk.StringVar(),
                    "in.GST": tk.StringVar(),
                    "no.GST": tk.BooleanVar(),
                    "Number": tk.StringVar(),
                    "Origin": tk.StringVar(),
                    "Origin_Upload": tk.BooleanVar(),
                    "V1": tk.StringVar(),
                    "V1_Upload": tk.BooleanVar(),
                    "V2": tk.StringVar(),
                    "V2_Upload": tk.BooleanVar(),
                    "V3": tk.StringVar(),
                    "V3_Upload": tk.BooleanVar(),
                    "Paid": tk.StringVar(),
                    "Paid.GST": tk.StringVar(),
                    "Content": [
                        {
                            "Service": tk.StringVar(),
                            "Fee": tk.StringVar(),
                            "in.GST": tk.StringVar(),
                            "no.GST": tk.BooleanVar(),
                            "Description": tk.StringVar(),
                            "State": tk.StringVar(value="Draft"),
                            "Upload": tk.BooleanVar(),
                            "Number": tk.StringVar(),
                            "Asana_id": tk.StringVar(),
                            "Xero_id": tk.StringVar()
                        } for _ in range(self.conf["n_bills"])
                    ]
                }
            ist_update_fun = lambda service : lambda a, b, c: self.app._ist_update(bills["Details"][service]["Fee"],
                                                                                   bills["Details"][service]["in.GST"],
                                                                                   bills["Details"][service]["no.GST"])

            bills["Details"][service]["Fee"].trace("w", ist_update_fun(service))
            bills["Details"][service]["no.GST"].trace("w", ist_update_fun(service))

            ist_update_fuc = lambda i, service: lambda a, b, c: self.app._ist_update(bills["Details"][service]["Content"][i]["Fee"],
                                                                            bills["Details"][service]["Content"][i]["in.GST"],
                                                                            bills["Details"][service]["Content"][i]["no.GST"])
            bill_color_code_fuc = lambda i, service: lambda a, b, c: self.bill_color_code(bills["Details"][service]["Content"][i]["State"], self.bill_dic[service]["Content"][i]["in.GST"])
            sum_update_fun = lambda service : lambda a, b, c: self.app._sum_update(
                    [bills["Details"][service]["Content"][j]["Fee"] for j in range(self.conf["n_bills"])],
                    bills["Details"][service]["Paid"])
            gst_sum_update_fun = lambda service: lambda a, b, c:self.app._sum_update(
                [bills["Details"][service]["Content"][j]["in.GST"] for j in range(self.conf["n_bills"])],
                bills["Details"][service]["Paid.GST"])
            bill_number_update_fun = lambda i, service :lambda a,b,c: self._update_bill_number(i, service)
            for i in range(self.conf["n_bills"]):
                bills["Details"][service]["Content"][i]["Number"].trace("w", bill_number_update_fun(i, service))
                bills["Details"][service]["Content"][i]["Fee"].trace("w", ist_update_fuc(i, service))
                bills["Details"][service]["Content"][i]["no.GST"].trace("w", ist_update_fuc(i, service))
                bills["Details"][service]["Content"][i]["Fee"].trace("w", sum_update_fun(service))
                bills["Details"][service]["Content"][i]["in.GST"].trace("w", gst_sum_update_fun(service))
                bills["Details"][service]["Content"][i]["State"].trace("w", bill_color_code_fuc(i, service))

            extra_list = ["Origin", "V1", "V2", "V3"]

            extra_update = lambda service: lambda a, b, c: self.app._sum_update(
                    [bills["Details"][service][e] for e in extra_list], bills["Details"][service]["Fee"])
            for extra in extra_list:
                bills["Details"][service][extra].trace("w", extra_update(service))

            fee_update = lambda service: lambda a, b, c: self.app._sum_update(
                [value["Fee"] for value in bills["Details"].values()], self.data["Bills"]["Fee"])
            bills["Details"][service]["Fee"].trace("w", fee_update(service))

            ist_update = lambda service: lambda a, b, c: self.app._sum_update(
                [value["in.GST"] for value in bills["Details"].values()], self.data["Bills"]["in.GST"])

            bills["Details"][service]["in.GST"].trace("w", ist_update(service))



        self.bill_frame = tk.LabelFrame(self.main_context_frame, text="Bill Details", font=self.conf["font"])
        self.bill_frame.grid(row=0, column=1, sticky="ns")

        self.bill_dic = dict()
        self.bill_frames= dict()


        title_frame = tk.LabelFrame(self.bill_frame)
        title_frame.pack(fill=tk.X)
        tk.Label(title_frame, width=10, text="Bill Number", font=self.conf["font"]).grid(row=0, column=0)
        tk.Label(title_frame, width=15, text="Description", font=self.conf["font"]).grid(row=0, column=1)
        tk.Label(title_frame, width=10, text="ex.GST", font=self.conf["font"]).grid(row=0, column=2)
        tk.Label(title_frame, width=10, text="in.GST", font=self.conf["font"]).grid(row=0, column=3)

        sub_title = tk.LabelFrame(self.bill_frame, height=36)
        sub_title.pack(fill=tk.X)
        tk.Label(sub_title, width=10, text="Origin", font=self.conf["font"]).grid(row=0, column=0, padx=(290, 0), pady=4)
        tk.Label(sub_title, width=10, text="V1", font=self.conf["font"]).grid(row=0, column=1, padx=(60, 0))
        tk.Label(sub_title, width=10, text="V2", font=self.conf["font"]).grid(row=0, column=2, padx=(70, 0))
        tk.Label(sub_title, width=10, text="V3", font=self.conf["font"]).grid(row=0, column=3, padx=(70, 0))

        total_frame = tk.LabelFrame(self.bill_frame)
        total_frame.pack(side=tk.BOTTOM, fill=tk.X)
        tk.Label(total_frame, width=35, text="Bill Total", font=self.conf["font"]).grid(row=0, column=0)
        self.total_label = tk.Label(total_frame, width=15, textvariable=bills["Fee"], font=self.conf["font"])
        self.total_label.grid(row=0, column=1)
        self.total_ingst_label = tk.Label(total_frame, width=15, textvariable=bills["in.GST"], font=self.conf["font"])
        self.total_ingst_label.grid(row=1, column=1)
    def update_bill(self, var):
        details = self.data["Bills"]["Details"]
        service = var["Service"].get()
        include = var["Include"].get()
        if include:
            details[service]["Include"].set(True)
            if not service in self.bill_frames:
                self.bill_frames[service] = tk.LabelFrame(self.bill_frame)
                self.bill_dic[service] = {
                    "Service": tk.Entry(self.bill_frames[service],
                                        textvariable=details[service]["Service"],
                                        width=20,
                                        font=self.conf["font"],
                                        fg="blue"),
                    # "Fee": tk.Label(self.bill_frames[service],
                    #                 textvariable=details[service]["Fee"],
                    #                 width=10,
                    #                 font=self.conf["font"]),
                    "Fee": tk.Frame(self.bill_frames[service]),
                    "no.GST": tk.Checkbutton(self.bill_frames[service],
                                             variable=details[service]["no.GST"],
                                             text="No GST",
                                             width=6),
                    "Origin": {
                        "Fee": tk.Entry(self.bill_frames[service],
                                        textvariable=details[service]["Origin"], width=10, font=self.conf["font"], fg="blue"),
                        "Upload": tk.Button(self.bill_frames[service],
                                            command=lambda: self.upload_bills(service=service, bill_description=details[service]["Service"],
                                                                              origin="Origin"), width=10, text="Upload", bg="cyan")
                    },
                    "V1": {
                        "Fee": tk.Entry(self.bill_frames[service],
                                        textvariable=details[service]["V1"], width=10, font=self.conf["font"], fg="blue"),
                        "Upload": tk.Button(self.bill_frames[service],
                                            command=lambda: self.upload_bills(service=service, bill_description=details[service]["Service"],
                                                                              origin="V1"), width=10, text="Upload", bg="cyan")
                    },
                    "V2": {
                        "Fee": tk.Entry(self.bill_frames[service],
                                        textvariable=details[service]["V2"], width=10, font=self.conf["font"], fg="blue"),
                        "Upload": tk.Button(self.bill_frames[service],
                                            command=lambda: self.upload_bills(service=service,bill_description=details[service]["Service"],
                                                                              origin="V2"), width=10, text="Upload", bg="cyan")
                    },
                    "V3": {
                        "Fee": tk.Entry(self.bill_frames[service],
                                        textvariable=details[service]["V3"], width=10, font=self.conf["font"], fg="blue"),
                        "Upload": tk.Button(self.bill_frames[service],
                                            command=lambda: self.upload_bills(service=service,bill_description=details[service]["Service"],
                                                                              origin="V3"), width=10, text="Upload", bg="cyan")
                    },
                    "Expand": [tk.LabelFrame(self.bill_frame) for _ in range(self.conf["n_items"])]

                }
                upload_file_func = lambda i: lambda : self.upload_sub_fee(service, details[service]["Content"][i]["Number"], i)
                self.bill_dic[service]["Content"] = [
                    {
                        "Number": tk.Entry(self.bill_dic[service]["Expand"][i],
                                           textvariable=details[service]["Content"][i]["Number"],
                                           width=2,
                                           font=self.conf["font"],
                                           fg="blue"
                                           ),
                        "Service": tk.Entry(self.bill_dic[service]["Expand"][i],
                                            textvariable=details[service]["Content"][i]["Service"],
                                            font=self.conf["font"],
                                            width=18,
                                            fg="blue"),
                        "Fee": tk.Entry(self.bill_dic[service]["Expand"][i],
                                        width=10,
                                        textvariable=details[service]["Content"][i]["Fee"],
                                        font=self.conf["font"],
                                        fg="blue",
                                        state=tk.DISABLED),
                        "in.GST": tk.Label(self.bill_dic[service]["Expand"][i],
                                           width=10,
                                           textvariable=details[service]["Content"][i]["in.GST"],
                                           font=self.conf["font"]),
                        "no.GST": tk.Checkbutton(self.bill_dic[service]["Expand"][i],
                                                 variable=details[service]["Content"][i]["no.GST"],
                                                 width=6,
                                                 text="No GST"),
                        "Description": tk.Entry(self.bill_dic[service]["Expand"][i],
                                                width=30,
                                                textvariable=details[service]["Content"][i]["Description"],
                                                font=self.conf["font"],
                                                fg="blue"),
                        "Upload": tk.Button(self.bill_dic[service]["Expand"][i],
                                            command=upload_file_func(i),
                                            width=8,
                                            text="Upload", bg="cyan")
                    } for i in range(self.conf["n_bills"])
                ]

                # self.bill_dic[service]["Paid"] = tk.Label(self.bill_dic[service]["Expand"][-1],
                #                                           textvariable=details[service]["Paid"],
                #                                           width=10,
                #                                           font=self.conf["font"])
                self.bill_dic[service]["Paid"] = tk.Frame(self.bill_dic[service]["Expand"][-1])



                tk.Label(self.bill_dic[service]["Paid"], text="$").grid(row=0, column=0)
                tk.Label(self.bill_dic[service]["Paid"], textvariable=details[service]["Paid"], width=6).grid(row=0, column=1)
                tk.Label(self.bill_dic[service]["Paid"], text="/").grid(row=0, column=2)
                tk.Label(self.bill_dic[service]["Paid"], text="$").grid(row=0, column=3)
                tk.Label(self.bill_dic[service]["Paid"], textvariable=details[service]["Paid.GST"], width=6).grid(row=0, column=4)


                self.bill_dic[service]["Service"].grid(row=0, column=0)
                self.bill_dic[service]["no.GST"].grid(row=0, column=1)
                self.bill_dic[service]["Origin"]["Fee"].grid(row=0, column=2)
                self.bill_dic[service]["Origin"]["Upload"].grid(row=0, column=3)
                self.data["Bills"]["Details"][service]["Origin_Upload"].trace("w", lambda a, b, c: self.config_button(self.data["Bills"]["Details"][service]["Origin_Upload"], self.bill_dic[service]["Origin"]["Upload"]))
                self.bill_dic[service]["V1"]["Fee"].grid(row=0, column=4)
                self.bill_dic[service]["V1"]["Upload"].grid(row=0, column=5)
                self.data["Bills"]["Details"][service]["V1_Upload"].trace("w", lambda a, b, c: self.config_button(
                    self.data["Bills"]["Details"][service]["V1_Upload"],
                    self.bill_dic[service]["V1"]["Upload"]))
                self.bill_dic[service]["V2"]["Fee"].grid(row=0, column=6)
                self.bill_dic[service]["V2"]["Upload"].grid(row=0, column=7)
                self.data["Bills"]["Details"][service]["V2_Upload"].trace("w", lambda a, b, c: self.config_button(
                    self.data["Bills"]["Details"][service]["V2_Upload"],
                    self.bill_dic[service]["V2"]["Upload"]))
                self.bill_dic[service]["V3"]["Fee"].grid(row=0, column=8)
                self.bill_dic[service]["V3"]["Upload"].grid(row=0, column=9)
                self.data["Bills"]["Details"][service]["V3_Upload"].trace("w", lambda a, b, c: self.config_button(
                    self.data["Bills"]["Details"][service]["V3_Upload"],
                    self.bill_dic[service]["V3"]["Upload"]))
                self.bill_dic[service]["Fee"].grid(row=0, column=10)

                tk.Label(self.bill_dic[service]["Fee"], text="$").grid(row=0, column=0)
                tk.Label(self.bill_dic[service]["Fee"], textvariable=details[service]["Fee"], width=6).grid(row=0, column=1)
                tk.Label(self.bill_dic[service]["Fee"], text="/").grid(row=0, column=2)
                tk.Label(self.bill_dic[service]["Fee"], text="$").grid(row=0, column=3)
                tk.Label(self.bill_dic[service]["Fee"], textvariable=details[service]["in.GST"], width=6).grid(row=0, column=4)


                self.bill_dic[service]["Paid"].pack(side=tk.RIGHT)
                config_button_func = lambda i : lambda a, b, c: self.config_button(self.data["Bills"]["Details"][service]["Content"][i]["Upload"], self.bill_dic[service]["Content"][i]["Upload"])
                for i in range(self.conf["n_bills"]):
                    self.bill_dic[service]["Content"][i]["Number"].grid(row=0, column=0, padx=(0, 3))
                    self.bill_dic[service]["Content"][i]["Service"].grid(row=0, column=1)
                    self.bill_dic[service]["Content"][i]["Fee"].grid(row=0, column=2)
                    self.bill_dic[service]["Content"][i]["in.GST"].grid(row=0, column=3)
                    self.bill_dic[service]["Content"][i]["no.GST"].grid(row=0, column=4)
                    self.bill_dic[service]["Content"][i]["Description"].grid(row=0, column=5)
                    self.bill_dic[service]["Content"][i]["Upload"].grid(row=0, column=6, padx=(250, 0))
                    self.data["Bills"]["Details"][service]["Content"][i]["Upload"].trace("w", config_button_func(i))
                    # self.bill_dic[service]["Content"][i]["Upload"].grid(row=0, column=6, sticky="w")
            self.bill_frames[service].pack(fill=tk.X)
            for i in range(self.conf["n_items"]):
                self.bill_dic[service]["Expand"][i].pack(fill=tk.X)

        else:
            details[service]["Include"].set(False)
            self.bill_frames[service].pack_forget()
            for i in range(self.conf["n_items"]):
                self.bill_dic[service]["Expand"][i].pack_forget()
    def profit_part(self):
        profits = {
            "Details": dict(),
            "Fee": tk.StringVar(),
            "in.GST": tk.StringVar()
        }
        self.data["Profits"] = profits
        invoices_details = self.data["Invoices"]["Details"]
        bills_details = self.data["Bills"]["Details"]
        profits_details = self.data["Profits"]["Details"]
        for service in self.conf["invoice_list"]:
            profits["Details"][service] = {
                "Include": tk.BooleanVar(),
                "Fee": tk.StringVar(),
                "in.GST": tk.StringVar()
            }
            _ = lambda service: lambda a, b, c: self.app._minus_update(
                invoices_details[service]["Fee"],
                bills_details[service]["Fee"],
                profits_details[service]["Fee"])

            invoices_details[service]["Fee"].trace("w", _(service))


            _ = lambda service :lambda a, b, c: self.app._minus_update(invoices_details[service]["Fee"],
                                                                                       bills_details[service]["Fee"],
                                                                                       profits_details[service]["Fee"])

            bills_details[service]["Fee"].trace("w", _(service))


            _ = lambda service: lambda a, b, c: self.app._minus_update(
                invoices_details[service]["in.GST"],
                bills_details[service]["in.GST"],
                profits_details[service]["in.GST"])

            invoices_details[service]["in.GST"].trace("w", _(service))

            _ = lambda service: lambda a, b, c: self.app._minus_update(
                invoices_details[service]["in.GST"],
                bills_details[service]["in.GST"],
                profits_details[service]["in.GST"])

            bills_details[service]["in.GST"].trace("w", _(service))


            _ = lambda service: lambda a, b, c: self.app._sum_update(
                [service["Fee"] for service in profits_details.values()], self.data["Profits"]["Fee"])

            profits_details[service]["Fee"].trace("w", _(service))

            _ = lambda service: lambda a, b, c: self.app._sum_update(
                [service["in.GST"] for service in profits_details.values()], self.data["Profits"]["in.GST"])

            profits_details[service]["in.GST"].trace("w", _(service))

        self.profit_frame = tk.LabelFrame(self.main_context_frame, text="Profit", font=self.conf["font"])
        self.profit_frame.grid(row=0, column=2, sticky="ns")

        title_frame = tk.LabelFrame(self.profit_frame)
        tk.Label(title_frame, width=10, text="ex.GST/in.GST", font=self.conf["font"]).grid(row=0, column=0)
        title_frame.pack()

        _ = tk.LabelFrame(self.profit_frame)
        tk.Label(_, text="").grid(row=0, column=0, pady=6)
        _.pack(fill=tk.X)

        self.profit_dic = dict()
        self.profit_frames = dict()

        total_frame = tk.LabelFrame(self.profit_frame)
        tk.Label(total_frame, width=10, textvariable=profits["Fee"], font=self.conf["font"]).grid(row=0, column=0)
        tk.Label(total_frame, width=10, textvariable=profits["in.GST"], font=self.conf["font"]).grid(row=1, column=0)

        total_frame.pack(side=tk.BOTTOM)
    def update_profit(self, var):
        profits_details = self.data["Profits"]["Details"]
        service = var["Service"].get()
        include = var["Include"].get()
        if include:
            # if service in archive:
            #     profits_details[service] = archive.pop(service)
            # else:
            #     profits_details[service] = {
            #         "Fee": tk.StringVar(),
            #         "in.GST": tk.StringVar(),
            #         "Context": [
            #             {
            #                 "Fee": tk.StringVar(),
            #                 "in.GST": tk.StringVar(),
            #             } for _ in range(self.conf["n_items"]-1)
            #         ]
            #     }
            profits_details[service]["Include"].set(True)
            if not service in self.profit_frames:
                self.profit_frames[service] = tk.LabelFrame(self.profit_frame)
                self.profit_dic[service] = {
                    "Fee": tk.Label(self.profit_frames[service],
                                    textvariable=profits_details[service]["Fee"],
                                    width=10,
                                    font=self.conf["font"]),
                    "in.GST": tk.Label(self.profit_frames[service],
                                       textvariable=profits_details[service]["in.GST"],
                                       width=10,
                                       font=self.conf["font"]),
                    "Expand": [tk.LabelFrame(self.profit_frame) for _ in range(self.conf["n_items"])]
                }

            # self.data["Invoices"]["Fee"].trace("w", lambda a, b, c: )
                for i in range(self.conf["n_items"]):
                    tk.Label(self.profit_dic[service]["Expand"][i], text="").pack(pady=2)
                self.profit_dic[service]["Fee"].pack(pady=1)

            self.profit_frames[service].pack()
            for i in range(self.conf["n_items"]):
                self.profit_dic[service]["Expand"][i].pack(fill=tk.X)
            self.app._sum_update(
                [service["Fee"] for service in profits_details.values()],
                self.data["Profits"]["Fee"])
            # self.profit_dic[service]["in.GST"].grid(row=1, column=0)

        else:
            profits_details[service]["Include"].set(False)
            self.profit_frames[service].pack_forget()
            for i in range(self.conf["n_items"]):
                self.profit_dic[service]["Expand"][i].pack_forget()
            self.app._sum_update(
                [service["Fee"] for service in profits_details.values()],
                self.data["Profits"]["Fee"])
    def fee_acceptance_part(self):
        self.data["Verbal Acceptance Note"] = tk.StringVar()
        self.data["Fee_Acceptance_Upload"] = tk.BooleanVar()
        self.data["Verbal_Acceptance_Upload"] = tk.BooleanVar()
        fee_acceptance_frame = tk.LabelFrame(self.main_context_frame, text="Upload Fee Acceptance", font=self.conf["font"])
        fee_acceptance_frame.grid(row=1, column=1, sticky="news", columnspan=2)

        tk.Label(fee_acceptance_frame, text="Fee Acceptance", font=self.conf["font"]).grid(row=0, column=0)
        fee_upload_button =tk.Button(fee_acceptance_frame, text="Upload", font=self.conf["font"], bg="cyan",
                                     command=self.upload_fee_acceptance)
        fee_upload_button.grid(row=0, column=1)
        self.data["Fee_Acceptance_Upload"].trace("w", self.config_button(self.data["Fee_Acceptance_Upload"], fee_upload_button))

        tk.Label(fee_acceptance_frame, text="Verbal Acceptance", font=self.conf["font"]).grid(row=1, column=0)
        verbal_acceptance_button = tk.Button(fee_acceptance_frame, text="Confirm", font=self.conf["font"], bg="cyan",
                                             command=self.upload_verbal_acceptance)
        verbal_acceptance_button.grid(row=1, column=1)
        self.data["Verbal_Acceptance_Upload"].trace("w", self.config_button(self.data["Verbal_Acceptance_Upload"], verbal_acceptance_button))


        tk.Label(fee_acceptance_frame, text="Note: ", font=self.conf["font"]).grid(row=1, column=2)
        tk.Entry(fee_acceptance_frame, width=50, textvariable=self.data["Verbal Acceptance Note"], font=self.conf["font"], fg="Blue").grid(row=1, column=3)

    def _on_mousewheel(self, event):
        self.main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    def reset_scrollregion(self, event):
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
    def update_invoice_sum(self, invoice_list):
        invoice_items = get_invoice_item(self.app)
        for i in range(6):
            invoice_list[i]["Fee"].set(sum([float(inv["Fee"])  for inv in invoice_items[i]]))
    # def update_fee_label(self, fee, label):
    #     if len(fee.get().strip())==0:
    #         label.config(text="$0/$0.0")
    #         return
    #     else:
    #         fee = fee.get()
    #     try:
    #         ingst = str(round(float(fee)*self.conf["tax rates"], 2))
    #         label.config(text=f"${fee}/${ingst}")
    #     except:
    #         label.config(text="Error")
    # def update_dollar_sign(self, fee, label):
    #     if len(fee.get()) == 0:
    #         label.config(text="")
    #         return
    #     else:
    #         fee = fee.get()
    #     label.config(text=f"${fee}")
    def generate_invoice_number(self):
        if not self.data["State"]["Fee Accepted"].get():
            self.messagebox.show_error("You need to upload a fee acceptance first")
            return

        generate = self.messagebox.ask_yes_no("Are you sure you want to generate the invoice number?")
        if generate:
            inv = None
            for i, value in enumerate(self.data["Invoices Number"]):
                if len(value["Number"].get()) == 0:
                    inv = i
                    break

            if inv is None:
                self.messagebox.show_error("You cant generate more than 6 invoices")
                return

            current_inv_number = self._get_current_invoice_number()
            self.data["Invoices Number"][inv]["Number"].set(current_inv_number)

    def unlock(self):
        update = self.messagebox.ask_yes_no("Do you want to unlock the all the invoice")
        if update:
            unlock_invoice(self.app)

    def _excel_print_invoice(self, i):
        self.data["Remittances"][i]["Preview_Upload"].set(True)
        excel_print_invoice(self.app, i)

    def _email_invoice(self, i):
        self.data["Remittances"][i]["Email_Upload"].set(True)
        email_invoice(self.app, i)
        if i == 0 and len(self.data["Project Info"]["Project"]["Project Number"].get()) == 0:
            number = self.data["Invoices Number"][0]["Number"].get()
            old_folder, new_folder, update = change_quotation_number(self.app, number)
            if update:
                self.data["Project Info"]["Project"]["Project Number"].set(number)
                project_quotation_dir = os.path.join(self.conf["database_dir"], "project_quotation_number_map.json")
                project_quotation_json = json.load(open(project_quotation_dir))
                project_quotation_json[number] = self.data["Project Info"]["Project"]["Quotation Number"].get()
                with open(project_quotation_dir, "w") as f:
                    json_object = json.dumps(project_quotation_json, indent=4)
                    f.write(json_object)
                # if len(self.data["Asana_id"].get()) != 0:
                update_asana(self.app)
                update_asana_invoices(self.app)
                # self.messagebox.file_info("Folder and Asana Renamed", old_folder, new_folder)
                # yes_asana = self.messagebox.file_ask_yes_no("Rename", old_folder, new_folder, "Asana")
                self.messagebox.file_info("Folder and Asana Renamed", old_folder, new_folder, "Please wait until the invoice turn red and Update Xero")
                # if update_xero:
                #     self.app._update_xero()
            else:
                self.messagebox.show_error(
                    f"Fail to rename the folder from {old_folder} to {new_folder}, Please close all the file relate in the folder")
                return
        else:
            update_asana(self.app)
            update_asana_invoices(self.app)
            update_xero = self.messagebox.ask_yes_no("The Invoices has upload to Asana, do you want to update Xero")
            if update_xero:
                self.app._update_xero()

        save(self.app)
        config_state(self.app)

    def _get_current_invoice_number(self):
        invoice_dir = os.path.join(self.conf["database_dir"], "invoices.json")
        invoices = json.load(open(invoice_dir))
        current_inv_number = list(invoices)[-1]
        today = date.today().strftime("%y%m")[1:]
        res = str(int(current_inv_number)+1) if current_inv_number.startswith(today) else today + "001"
        invoices[res] = "Backlog"
        with open(invoice_dir, "w", encoding='utf-8') as f:
            json.dump(invoices, f, ensure_ascii=False, indent=4)
        return res


    def config_button(self, var, button, *args):
        if var.get() == True:
            button.config(bg="grey")
        else:
            button.config(bg="cyan")


    def upload_remittance(self, invoice_number, part, i):
        database_dir = os.path.join(self.conf["database_dir"], self.data["Project Info"]["Project"]["Quotation Number"].get())
        remittance_dir = os.path.join(self.conf["remittances_address"], date.today().strftime("%Y%m"))
        if part == "Full":
            if self.data['Remittances'][i]["Type"].get() == "Part":
                self.messagebox.show_error("This invoice already has partial amount remittance")
                return
            filename = f"{invoice_number}, ${self.data['Invoices Number'][i]['in.GST'].get()}-${self.data['Invoices Number'][i]['in.GST'].get()}"
        else:
            if self.data['Remittances'][i]["Type"].get() == "Full":
                self.messagebox.show_error("This invoice already has full amount remittance")
                return
            part1_amount = self.data['Remittances'][i]['Part1'].get()
            part2_amount = self.data['Remittances'][i]['Part2'].get()
            part3_amount = self.data['Remittances'][i]['Part3'].get()
            full_amount = self.data['Invoices Number'][i]['in.GST'].get()
            if part == "Part1":
                if not isfloat(part1_amount) or float(part1_amount) > float(full_amount):
                    self.messagebox.show_error("The amount exceed the full amount")
                    return
                amount = part1_amount
            elif part == "Part2":
                if not isfloat(part1_amount) or not isfloat(part2_amount) or float(part1_amount) + float(part2_amount) > float(full_amount):
                    self.messagebox.show_error("The amount exceed the full amount")
                    return
                amount = str(float(part1_amount)+float(part2_amount))
            else:
                if not isfloat(part1_amount) or not isfloat(part2_amount) or not isfloat(part3_amount) or float(
                        part1_amount) + float(part2_amount) + float(part3_amount) > float(full_amount):
                    self.messagebox.show_error("The amount exceed the full amount")
                    return
                amount = str(float(part1_amount)+float(part2_amount)+float(part3_amount))
                fee = self.data['Invoices Number'][i]['in.GST'].get()
                if fee != amount:
                    self.messagebox.show_error(f"The Amount does not match, input amount {amount}, invoice ingst {fee}")
                    return
            filename = f"{invoice_number}, ${amount}-${self.data['Invoices Number'][i]['in.GST'].get()}"

        if self.data["Remittances"][i][part+"_Upload"].get():
            rewrite = self.messagebox.ask_yes_no(f"Existing file found, Do you want to rewrite")
            if not rewrite:
                return
        file = filedialog.askopenfilename()
        if file == "":
            return
        try:
            folder_dir = os.path.join(database_dir, filename + os.path.splitext(file)[1])
            shutil.copy(file, folder_dir)
        except PermissionError:
            self.messagebox.show_error("Please Close the file before you upload it")
            return
        except Exception as e:
            print(e)
            self.messagebox.show_error("Some error occurs, please contact Administrator")
            return

        try:
            if not os.path.exists(remittance_dir):
                os.makedirs(remittance_dir)
            file_path = os.path.join(remittance_dir, filename + os.path.splitext(file)[1])
            shutil.copy(file, file_path)
        except PermissionError:
            self.messagebox.show_error("Please Close the file before you upload it")
            return
        except Exception as e:
            print(e)
            self.messagebox.show_error("Some error occurs, please contact Administrator")
            return

        if part == "Full":
            self.data['Remittances'][i]['Type'].set("Full")
        else:
            self.data['Remittances'][i]['Type'].set("Part")
        self.data["Remittances"][i][part+"_Upload"].set(True)
        self.messagebox.file_info("Upload", file, folder_dir)

    def upload_bills(self, service, bill_description, origin):
        database_dir = os.path.join(self.conf["database_dir"], self.data["Project Info"]["Project"]["Quotation Number"].get())
        if len(bill_description.get()) == 0:
            self.messagebox.show_error("You need to upload the description")
            return
        filename = f"{bill_description}-{origin}"
        if self.data["Bills"]["Details"][service][origin+"_Upload"].get():
            rewrite = self.messagebox.ask_yes_no(f"Existing file found, Do you want to rewrite")
            if not rewrite:
                return
        file = filedialog.askopenfilename()
        if file == "":
            return
        try:
            folder_dir = os.path.join(database_dir, filename + os.path.splitext(file)[1])
            shutil.copy(file, folder_dir)
        except PermissionError:
            self.messagebox.show_error("Please Close the file before you upload it")
            return
        except Exception as e:
            print(e)
            self.messagebox.show_error("Some error occurs, please contact Administrator")
            return
        self.data["Bills"]["Details"][service][origin + "_Upload"].set(True)
        self.messagebox.file_info("Upload", file, folder_dir)

    def upload_sub_fee(self, service, bill_number, i):
        database_dir = os.path.join(self.conf["database_dir"], self.data["Project Info"]["Project"]["Quotation Number"].get())
        if len(bill_number.get()) == 0:
            self.messagebox.show_error("You need to enter a bill number")
            return
        elif len(self.data["Project Info"]["Project"]["Project Number"].get()) == 0:
            self.messagebox.show_error("Please generate first invoice number before you upload bill file")
            return
        if self.data["Bills"]["Details"][service]["Content"][i]["Upload"].get():
            rewrite = self.messagebox.ask_yes_no(f"Existing file found, Do you want to rewrite")
            if not rewrite:
                return
        file = filedialog.askopenfilename()
        if file == "":
            return
        try:
            filename = self.data["Project Info"]["Project"]["Project Number"].get()+ bill_number.get() + "-" + os.path.basename(file)
            folder_dir = os.path.join(database_dir, filename)
            shutil.copy(file, folder_dir)
        except PermissionError:
            self.messagebox.show_error("Please Close the file before you upload it")
            return
        except Exception as e:
            print(e)
            self.messagebox.show_error("Some error occurs, please contact Administrator")
            return

        try:
            send_email_with_attachment(folder_dir)
        except Exception as e:
            print(e)
            self.messagebox.show_error("Unable to upload the File to xero")
            return
        self.messagebox.file_info("Upload", file, folder_dir, "And the bill has been sent to Xero")
        self.data["Bills"]["Details"][service]["Content"][i]["Upload"].set(True)

    def upload_fee_acceptance(self):
        self.data["Fee_Acceptance_Upload"].set(True)
        database_dir = os.path.join(self.conf["database_dir"],
                                    self.data["Project Info"]["Project"]["Quotation Number"].get())

        if not self.data["State"]["Email to Client"].get():
            self.messagebox.show_error("You need to send the Email to Client first")
            return
        elif self.data["State"]["Fee Accepted"].get():
            update = self.messagebox.ask_yes_no("Fee Already Accepted, do you want to update again?")
            if not update:
                return

        acceptance_list = [file for file in os.listdir(os.path.join(database_dir)) if
                           str(file).startswith("Fee Acceptance")]

        if len(acceptance_list) != 0:
            current_revision = str(max([str(pdf).split(" ")[-1].split(".")[0] for pdf in acceptance_list]))
            overwrite = self.messagebox.ask_yes_no(f"Current Fee Acceptance {current_revision} found, do you want to Overwrite")
            if overwrite:
                filename = f"Fee Acceptance Rev {current_revision}"
            else:
                filename = f"Fee Acceptance Rev {str(int(current_revision) + 1)}"
        else:
            filename = "Fee Acceptance Rev 1"
        rewrite = True
        if rewrite:
            file = filedialog.askopenfilename()
            if file == "":
                return
            try:
                folder_dir = os.path.join(database_dir, filename + os.path.splitext(file)[1])
                shutil.copy(file, folder_dir)
            except PermissionError:
                self.messagebox.show_error("Please Close the file before you upload it")
                return
            except Exception as e:
                print(e)
                self.messagebox.show_error("Some error occurs, please contact Administrator")
                return
            self.data["State"]["Fee Accepted"].set(True)
            self.app.log.log_fee_accept_file(self.app.user, self.data["Project Info"]["Project"]["Quotation Number"].get())
            update = self.messagebox.file_ask_yes_no("Upload", file, folder_dir, "Asana")
            if update:
                update_asana(self.app)
                update_asana_invoices(self.app)
                self.messagebox.show_info("Success update Asana")
            save(self.app)
            config_state(self.app)
            config_log(self.app)


    def upload_verbal_acceptance(self):
        self.data["Verbal_Acceptance_Upload"].set(True)
        database_dir = os.path.join(self.conf["database_dir"], self.data["Project Info"]["Project"]["Quotation Number"].get())
        resource_dir = os.path.join(self.conf["resource_dir"], "txt", "Verbal Fee Acceptance.txt")

        if not self.data["State"]["Email to Client"].get():
            self.messagebox.show_error("You need to send the Email to Client first")
            return
        elif self.data["State"]["Fee Accepted"].get():
            update = self.messagebox.ask_yes_no("Fee Accepted Already, Do you want to Upload Fee Acceptance Again?")
            if not update:
                return

        update = self.messagebox.ask_yes_no("Do you want to verbal confirm and update Asana ?")
        if update:
            self.data["State"]["Fee Accepted"].set(True)
            shutil.copy(resource_dir, database_dir)
            with open(os.path.join(database_dir, "Verbal Fee Acceptance.txt"), "w") as f:
                f.write(self.data["Verbal Acceptance Note"].get())
            f.close()
            update_asana(self.app)
            update_asana_invoices(self.app)
            self.app.log.log_fee_accept_file(self.app.user, self.data["Project Info"]["Project"]["Quotation Number"].get())
            save(self.app)
            config_state(self.app)
            config_log(self.app)
            self.messagebox.show_info("Verbal Fee Acceptance logged and asana updated.")

    def _update_bill_number(self, i, service, *args):
        number = self.data["Bills"]["Details"][service]["Content"][i]["Number"].get()
        if len(number) == 0:
            return
        self.data["Bills"]["Details"][service]["Content"][i]["Number"].set(number[-1].upper())

    def invoice_color_code(self, *args):
        for i in range(6):
            state = self.data["Invoices Number"][i]["State"].get()
            if state=="Backlog":
                for label in self.invoice_label_list[i]:
                    label.config(bg=self.app.cget('bg'))
            elif state=="Sent":
                for label in self.invoice_label_list[i]:
                    label.config(bg="red")
                self._config_radiobutton(i, tk.DISABLED)
            elif state=="Paid":
                for label in self.invoice_label_list[i]:
                    label.config(bg="green")
                self._config_radiobutton(i, tk.DISABLED)
            elif state=="Void":
                for label in self.invoice_label_list[i]:
                    label.config(bg="purple")
                self._config_radiobutton(i, tk.DISABLED)
            else:
                raise ValueError

    def _config_radiobutton(self, i, config):
        for service, value in self.invoice_dic.items():
            value["Invoice"][i].config(state=config)
            for j, item in enumerate(value["Content"]):
                item["Invoice"][i].config(state=config)
        for service, value in self.invoice_dic.items():
            if self.data["Invoices"]["Details"][service]["Number"].get() == f"INV{str(i+1)}":
                for inv in value["Invoice"]:
                    inv.config(state=config)
            for j, item in enumerate(value["Content"]):
                if self.data["Invoices"]["Details"][service]["Content"][j]["Number"].get() == f"INV{str(i+1)}":
                    for inv in value["Content"][j]["Invoice"]:
                        inv.config(state=config)

    def bill_color_code(self, state, label, *args):
        state = state.get()
        if state=="Draft":
            label.config(bg=self.app.cget('bg'))
        elif state=="Awaiting approval":
            label.config(bg="red")
        elif state=="Awaiting payment":
            label.config(bg="orange")
        elif state=="Paid":
            label.config(bg="green")
        elif state=="Void":
            label.config(bg="Purple")
        else:
            raise ValueError

    # def no_gst_update(self,nogst, fee, ingst, *args):
    #     if nogst.get():
    #         fee.trace("w", lambda a, b, c: self.no_gst(fee, ingst))
    #         self.no_gst(fee, ingst)
    #     else:
    #         fee.trace("w", lambda a, b, c: self.app._ist_update(fee, ingst))
    #         self.app._ist_update(fee, ingst)
    #
    # def no_gst(self, fee, ingst, *args):
    #     ingst.set(fee.get())

    # def _expand_bill(self, service):
    #     details = self.data["Invoices"]["Details"]
    #     bill_fee = self.data["Invoices Number"]["Bill Details"]
    #     if details[service]["Expand"].get():
    #         bill_fee[service]["Expand"].set(True)
    #         for i in range(self.conf["n_items"]):
    #             self.bill_dic[service]["Content"]["Details"][i]["Service"].grid(row=2*i+2, column=0, pady=(2, 0))
    #             self.bill_dic[service]["Content"]["Details"][i]["Fee"].grid(row=2*i+2, column=1)
    #             self.bill_dic[service]["Content"]["Details"][i]["in.GST"].grid(row=2*i+3, column=1)
    #             for j in range(6):
    #                 self.bill_dic[service]["Content"]["Details"][i]["Number"][j].grid(row=2*i+2, column=2+j, rowspan=2, padx=(2, 0))
    #         self.bill_dic[service]["Content"]["Service"].grid(row=2*(self.conf["n_items"]+1)+1, column=0)
    #         self.bill_dic[service]["Content"]["Fee"].grid(row=2*(self.conf["n_items"]+1)+1, column=1)
    #         self.bill_dic[service]["Content"]["in.GST"].grid(row=2*(self.conf["n_items"]+1)+2, column=1)
    #
    #         self.bill_dic[service]["Service"].config(state=tk.DISABLED)
    #         self.bill_dic[service]["Fee"].grid_forget()
    #         self.bill_dic[service]["in.GST"].grid_forget()
    #         for i in range(6):
    #             self.bill_dic[service]["Number"][i].grid_forget()
    #     else:
    #         bill_fee[service]["Expand"].set(False)
    #         for i in range(self.conf["n_items"]):
    #             self.bill_dic[service]["Content"]["Details"][i]["Service"].grid_forget()
    #             self.bill_dic[service]["Content"]["Details"][i]["Fee"].grid_forget()
    #             self.bill_dic[service]["Content"]["Details"][i]["in.GST"].grid_forget()
    #             for j in range(6):
    #                 self.bill_dic[service]["Content"]["Details"][i]["Number"][j].grid_forget()
    #         self.bill_dic[service]["Content"]["Service"].grid_forget()
    #         self.bill_dic[service]["Content"]["Fee"].grid_forget()
    #         self.bill_dic[service]["Content"]["in.GST"].grid_forget()
    #
    #         self.bill_dic[service]["Service"].config(state=tk.NORMAL)
    #         self.bill_dic[service]["Fee"].grid(row=0, column=1)
    #         self.bill_dic[service]["in.GST"].grid(row=1, column=1)
    #         for i in range(6):
    #             self.bill_dic[service]["Invoice"][i].grid(row=0, column=2+i, rowspan=2, padx=(2, 0))
    #     self.data["Invoices Number"]["Bill Fee"][service]["Invoice"].set("None")
    #     for i in range(3):
    #         self.data["Invoices Number"]["Bill Fee"][service]["Content"][i]["Invoice"].set("None")

    # def profit_expand(self, service):
    #     if self.app.data["Fee Proposal Page"]["Details"][service]["Expanded"].get():
    #         self.profit_data_dic[service]["Expanded"].set(True)
    #         for i in range(4):
    #             self.profit_dic[service]["expand frames"][i].grid(row=int(self.profit_dic[service]["Position"]) + i + 1,
    #                                                               column=0, sticky="ew")
    #         self.profit_data_dic[service]["Fee"].set("")
    #         self.profits_frame[service].config(height=28)
    #         self.profit_dic[service]["Fee"].grid_forget()
    #         self.profit_dic[service]["in.GST"].grid_forget()
    #     else:
    #         self.profit_data_dic[service]["Expanded"].set(False)
    #         self.profits_frame[service].config(height=56)
    #         for i in range(4):
    #             self.profit_dic[service]["expand frames"][i].grid_forget()
    #         self.profit_dic[service]["Fee"].grid(row=0, column=1)
    #         self.profit_dic[service]["in.GST"].grid(row=1, column=1, padx=(0, 2))
    #         for j in range(3):
    #             self.profit_data_dic[service]["Context"][j]["Fee"].set("")


    # def _update_profit(self, *args):
    #     Total = 0
    #     inGST_Total = 0
    #     for service_name, service in self.app.data["Fee Proposal Page"]["Details"].items():
    #         if not service["on"].get():
    #             continue
    #         if service["Expanded"].get():
    #             for i in range(3):
    #                 try:
    #                     invoice = service["Context"][i]["Fee"].get() if service["Context"][i]["Fee"].get() != "" else 0
    #                     bill = self.bill_data_dic[service_name]["Context"][i]["Fee"].get() if self.bill_data_dic[service_name]["Context"][i]["Fee"].get() != "" else 0
    #                     self.profit_data_dic[service_name]["Context"][i]["Fee"].set(int(invoice) - int(bill))
    #                     self.profit_data_dic[service_name]["Context"][i]["in.GST"].set(int(float(invoice)*1.1) - int(float(bill)*1.1))
    #                 except ValueError:
    #                     self.profit_data_dic[service_name]["Context"][i]["Fee"].set("Error")
    #                     self.profit_data_dic[service_name]["Context"][i]["in.GST"].set("Error")
    #         try:
    #             invoice = service["Fee"].get() if service["Fee"].get() != "" else 0
    #             bill = self.bill_data_dic[service_name]["Fee"].get() if self.bill_data_dic[service_name]["Fee"].get() != "" else 0
    #             self.profit_data_dic[service_name]["Fee"].set(int(invoice) - int(bill))
    #             self.profit_data_dic[service_name]["in.GST"].set(int(float(invoice)*1.1) - int(float(bill)*1.1))
    #             Total += int(invoice) - int(bill)
    #             inGST_Total += int(float(invoice)*1.1) - int(float(bill)*1.1)
    #         except ValueError:
    #             self.profit_data_dic[service_name]["Fee"].set("Error")
    #             self.profit_data_dic[service_name]["in.GST"].set("Error")
    #     self.total_profit_var.set(str(Total))
    #     self.total_in_profit_var.set(str(inGST_Total))

    # def update_radiobutton(self, var):
    #     service, index = var._name.split(" item ")
    #     index = int(index)
    #     if len(var.get()) == 0:
    #         for i in range(6):
    #             self.fee_dic[service]["expand frame"][index]["Invoice"][i].grid_forget()
    #     else:
    #         for i in range(6):
    #             self.fee_dic[service]["expand frame"][index]["Invoice"][i].grid(row=0, column=i + 2, padx=(0, 2),rowspan=2)
    #
    # def update_bill_radiobutton(self, var):
    #     service, index = var._name.split(" bill ")
    #     index = int(index)
    #     if len(var.get()) == 0:
    #         for i in range(6):
    #             self.bill_dic[service]["expand frame"][index]["Invoice"][i].grid_forget()
    #     else:
    #         for i in range(6):
    #             self.bill_dic[service]["expand frame"][index]["Invoice"][i].grid(row=0, column=i + 2, padx=(0, 2),rowspan=2)
