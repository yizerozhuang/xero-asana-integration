import shutil
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from utility import excel_print_invoice, email_invoice, save, config_log, config_state, change_quotation_number
from asana_function import update_asana, change_asana_quotation

import os
import json
from datetime import date

class FinancialPanelPage(tk.Frame):
    def __init__(self, parent, app):
        tk.Frame.__init__(self, parent)
        self.update_invoice_number = None
        self.parent = parent
        self.app = app
        self.data = app.data
        self.conf = app.conf
        
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
        self.data["Financial Panel"] = dict()

        self.invoice_part()
        self.invoice_function_part()
        self.bill_part()
        self.profit_part()
        self.fee_acceptance_part()

    def invoice_part(self):

        invoice = dict()
        self.data["Financial Panel"]["Invoice Details"] = invoice

        self.invoice_dic = dict()
        self.invoice_frames= dict()
        self.invoice_dic["Number"] = dict()

        self.invoice_frame = tk.LabelFrame(self.main_context_frame, text="Invoice Details", font=self.conf["font"])
        self.invoice_frame.grid(row=0, column=0)

        title_frame = tk.LabelFrame(self.invoice_frame)
        title_frame.pack(fill=tk.X)
        tk.Label(title_frame, width=35, text="Items", font=self.conf["font"]).grid(row=0, column=0)
        tk.Label(title_frame, width=15, text="ex.GST/in.GST", font=self.conf["font"]).grid(row=0, column=1)
        invoice_number_frame = tk.LabelFrame(self.invoice_frame)
        invoice_number_frame.pack(fill=tk.X)
        tk.Label(invoice_number_frame, width=35, text="Invoice Number", font=self.conf["font"]).grid(row=0, column=0)
        tk.Button(invoice_number_frame, width=15, text="Gen", command=self.generate_invoice_number, bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=1)

        for i in range(6):
            invoice[f"INV{str(i+1)}"] = {
                "Number": tk.StringVar(),
                "Fee": tk.StringVar(),
                "in.GST": tk.StringVar()
            }
            tk.Label(title_frame, width=8, text="INV"+str(i+1), font=self.conf["font"]).grid(row=0, column=i+2, sticky="ew", padx=(2,0))
            self.invoice_dic["Number"][f"INV{str(i+1)}"] = tk.Entry(invoice_number_frame, state=tk.DISABLED, textvariable=invoice[f"INV{str(i+1)}"]["Number"], width=8, font=self.conf["font"], fg="blue")
            self.invoice_dic["Number"][f"INV{str(i+1)}"].grid(row=0, column=i + 2, padx=(10, 0), sticky="ew")

        total_frame = tk.LabelFrame(self.invoice_frame)
        total_frame.pack(side=tk.BOTTOM, fill=tk.X)
        tk.Label(total_frame, width=35, text="Invoice Total", font=self.conf["font"]).grid(row=0, column=0)
        self.total_label = tk.Label(total_frame, width=15, text="$0", font=self.conf["font"])
        self.total_label.grid(row=0, column=1)
        self.data["Invoices"]["Fee"].trace("w", lambda a, b, c: self.update_dollar_sign(self.data["Invoices"]["Fee"], self.total_label))
        self.total_ingst_label = tk.Label(total_frame, width=15, text="$0", font=self.conf["font"])
        self.total_ingst_label.grid(row=1, column=1)
        self.data["Invoices"]["in.GST"].trace("w", lambda a,b,c: self.update_dollar_sign(self.data["Invoices"]["in.GST"], self.total_ingst_label))
        ist_fuc = lambda i: lambda a, b, c: self.app._ist_update(invoice[f"INV{str(i+1)}"]["Fee"], invoice[f"INV{str(i+1)}"]["in.GST"])
        self.invoice_total_list = []
        update_dollar_sign_fuc = lambda i: lambda a, b, c: self.update_dollar_sign(invoice[f"INV{str(i+1)}"]["Fee"], self.invoice_total_list[i])
        self.invoice_ingst_list = []
        update_ingst_dollar_sign_fuc = lambda i: lambda a, b, c: self.update_dollar_sign(invoice[f"INV{str(i + 1)}"]["in.GST"], self.invoice_ingst_list[i])

        for i in range(6):
            self.invoice_total_list.append(tk.Label(total_frame, width=8, font=self.conf["font"]))
            self.invoice_total_list[i].grid(row=0, column=2 + i, padx=(2,0))
            invoice[f"INV{str(i + 1)}"]["Fee"].trace("w", update_dollar_sign_fuc(i))
            self.invoice_ingst_list.append(tk.Label(total_frame, width=8, font=self.conf["font"]))
            self.invoice_ingst_list[i].grid(row=1, column=2 + i, padx=(2,0))
            invoice[f"INV{str(i+1)}"]["in.GST"].trace("w", update_ingst_dollar_sign_fuc(i))
            invoice[f"INV{str(i+1)}"]["Fee"].trace("w", ist_fuc(i))

        details = self.data["Invoices"]["Details"]
        invoice_details = self.data["Financial Panel"]["Invoice Details"]

        self.variation_label_list = []
        update_label_fuc = lambda i: lambda a, b, c: self.update_fee_label(self.data["Variation"][i]["Fee"], self.variation_label_list[i])
        for i in range(self.conf["n_variation"]):
            variation_frame = tk.LabelFrame(self.invoice_frame)
            variation_frame.pack(side=tk.BOTTOM, fill=tk.X)
            tk.Label(variation_frame, textvariable=self.data["Variation"][i]["Service"], width=35, font=self.conf["font"]).grid(row=0, column=0)
            self.variation_label_list.append(tk.Label(variation_frame, text="$0/$0.0", width=15, font=self.conf["font"]))
            self.variation_label_list[i].grid(row=0, column=1)
            self.data["Variation"][i]["Fee"].trace("w", update_label_fuc(i))
            for j in range(6):
                tk.Radiobutton(variation_frame, width=6, variable=self.data["Variation"][i]["Number"], value="INV" + str(j + 1)).grid(row=0, column=j+2, padx=(2,0))
            self.data["Variation"][i]["Fee"].trace("w", lambda a, b, c: self.update_invoice_sum(details, invoice_details))
            self.data["Variation"][i]["Number"].trace("w", lambda a, b, c: self.update_invoice_sum(details, invoice_details))

    def update_invoice(self, var):
        details = self.data["Invoices"]["Details"]
        invoice_details = self.data["Financial Panel"]["Invoice Details"]
        service = var["Service"].get()
        include = var["Include"].get()
        if include:
            if not service in self.invoice_dic:
                self.invoice_frames[service] = tk.LabelFrame(self.invoice_frame)
                self.invoice_dic[service] = {
                    "Service": tk.Label(self.invoice_frames[service],
                                        text=service + " design and documentation",
                                        width=35,
                                        font=self.conf["font"]),
                    "Fee": tk.Label(self.invoice_frames[service],
                                    text="$0/$0.0",
                                    width=15,
                                    font=self.conf["font"]),
                    "Invoice": [tk.Radiobutton(self.invoice_frames[service],
                                               width=6,
                                               variable=self.data["Invoices"]["Details"][service]["Number"],
                                               value="INV" + str(i + 1)) for i in range(6)],
                    "Expand": [tk.LabelFrame(self.invoice_frame, height=30) for _ in range(self.conf["n_items"])]
                }
                self.invoice_dic[service]["Content"] = [
                    {
                        "Service": tk.Label(self.invoice_dic[service]["Expand"][i],
                                            width=35,
                                            font=self.conf["font"],
                                            textvariable=details[service]["Content"][i]["Service"]),
                        "Fee": tk.Label(self.invoice_dic[service]["Expand"][i],
                                        width=15,
                                        font=self.conf["font"]),
                        "Invoice": [tk.Radiobutton(self.invoice_dic[service]["Expand"][i],
                                                   width=6,
                                                   variable=
                                                   self.app.data["Invoices"]["Details"][service]["Content"][i]["Number"],
                                                   value="INV" + str(j + 1)) for j in range(6)]
                    } for i in range(self.conf["n_items"])
                ]
                self.invoice_dic[service]["Service"].grid(row=0, column=0)
                self.invoice_dic[service]["Fee"].grid(row=0, column=1)
                for i in range(6):
                    self.invoice_dic[service]["Invoice"][i].grid(row=0, column=2 + i, padx=(2, 0))


                details[service]["Expand"].trace("w", lambda a, b, c: self._expand_invoice(service))

                details[service]["Fee"].trace("w", lambda a, b, c: self.update_invoice_sum(details, invoice_details))
                details[service]["Fee"].trace("w", lambda a, b, c: self.update_fee_label(details[service]["Fee"], self.invoice_dic[service]["Fee"]))
                details[service]["Number"].trace("w", lambda a, b, c: self.update_invoice_sum(details, invoice_details))
                update_label_fuc = lambda i: lambda a, b, c: self.update_fee_label(details[service]["Content"][i]["Fee"], self.invoice_dic[service]["Content"][i]["Fee"])
                for i in range(self.conf["n_items"]):
                    details[service]["Content"][i]["Fee"].trace("w", lambda a, b, c: self.update_invoice_sum(details,
                                                                                                             invoice_details))
                    details[service]["Content"][i]["Fee"].trace("w", update_label_fuc(i))

                    details[service]["Content"][i]["Number"].trace("w", lambda a, b, c: self.update_invoice_sum(details,
                                                                                                                invoice_details))
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
            self.invoice_dic[service]["Fee"].config(font=self.conf["font"]+["bold"])
        else:
            for i in range(self.conf["n_items"]):
                self.invoice_dic[service]["Content"][i]["Service"].grid_forget()
                self.invoice_dic[service]["Content"][i]["Fee"].grid_forget()
                for j in range(6):
                    self.invoice_dic[service]["Content"][i]["Invoice"][j].grid_forget()

            self.invoice_dic[service]["Fee"].config(font=self.conf["font"])
            for i in range(6):
                self.invoice_dic[service]["Invoice"][i].grid(row=0, column=2+i)
        details[service]["Number"].set("None")
        for i in range(self.conf["n_items"]):
            details[service]["Content"][i]["Number"].set("None")

    def invoice_function_part(self):
        remittance = {
            "INV1": {
                "Part1": tk.StringVar(),
                "Part2": tk.StringVar()
            },
            "INV2": {
                "Part1": tk.StringVar(),
                "Part2": tk.StringVar()
            },
            "INV3": {
                "Part1": tk.StringVar(),
                "Part2": tk.StringVar()
            },
            "INV4": {
                "Part1": tk.StringVar(),
                "Part2": tk.StringVar()
            },
            "INV5": {
                "Part1": tk.StringVar(),
                "Part2": tk.StringVar()
            },
            "INV6": {
                "Part1": tk.StringVar(),
                "Part2": tk.StringVar()
            }
        }
        self.data["Remittances"] = remittance

        invoice_function_frame = tk.LabelFrame(self.main_context_frame, text="Invoice Functions", font=self.conf["font"])
        invoice_function_frame.grid(row=1, column=0, sticky="ew")
        print_invoice_function = lambda i: lambda: excel_print_invoice(self.app, i)
        email_invoice_function = lambda i: lambda: email_invoice(self.app, i)
        upload_invoice_function = lambda i, part: lambda: self.upload_file(invoice=self.data["Financial Panel"]["Invoice Details"]["INV"+str(i+1)]["Number"].get(), part=part)
        for i in range(6):
            tk.Label(invoice_function_frame, width=20, text="Invoice Number: ", font=self.conf["font"]).grid(row=i, column=0)
            tk.Label(invoice_function_frame, width=10, textvariable=self.data["Financial Panel"]["Invoice Details"][f"INV{str(i+1)}"]["Number"], font=self.conf["font"]).grid(row=i, column=1)
            tk.Button(invoice_function_frame, text="Preview", font=self.conf["font"], bg="brown", fg="white",
                      command=print_invoice_function(i)).grid(row=i, column=2)
            tk.Button(invoice_function_frame, text="Email", font=self.conf["font"], bg="brown", fg="white",
                      command=email_invoice_function(i)).grid(row=i, column=3)
            tk.Button(invoice_function_frame, text="Upload full amount", font=self.conf["font"], bg="brown", fg="white",
                      command=upload_invoice_function(i, "Full Amount")).grid(row=i, column=4)
            tk.Entry(invoice_function_frame, width=10, font=self.conf["font"], fg="blue",
                     textvariable=remittance[f"INV{str(i+1)}"]["Part1"]).grid(row=i, column=5)
            tk.Button(invoice_function_frame, text="Upload Part1", bg="brown", fg="white", font=self.conf["font"],
                      command=upload_invoice_function(i, "Part 1")).grid(row=i, column=6)
            tk.Entry(invoice_function_frame, width=10, font=self.conf["font"], fg="blue",
                     textvariable=remittance[f"INV{str(i+1)}"]["Part2"]).grid(row=i, column=7)
            tk.Button(invoice_function_frame, text="Upload Part2", bg="brown", fg="white", font=self.conf["font"],
                      command=upload_invoice_function(i, "Part 2")).grid(row=i, column=8)
    def bill_part(self):
        bills = {
            "Details": dict(),
            "Fee": tk.StringVar(),
            "in.GST": tk.StringVar()
        }
        self.data["Bills"] = bills

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
    def update_bill(self, var):
        details = self.data["Bills"]["Details"]
        archive = self.data["Archive"]["Bill"]
        service = var["Service"].get()
        include = var["Include"].get()
        if include:
            if service in archive:
                details[service] = archive.pop(service)
            else:
                details[service] = {
                    "Service": tk.StringVar(),
                    "Fee": tk.StringVar(),
                    "in.GST": tk.StringVar(),
                    "Number": tk.StringVar(),
                    "Ori": tk.StringVar(),
                    "V1": tk.StringVar(),
                    "V2": tk.StringVar(),
                    "V3": tk.StringVar(),
                    "Paid": tk.StringVar(),
                    "Content": [
                        {
                            "Service": tk.StringVar(),
                            "Fee": tk.StringVar(),
                            "in.GST": tk.StringVar(),
                            "Number": tk.StringVar()
                        } for _ in range(self.conf["n_items"]-1)
                    ]
                }
                self.bill_frames[service] = tk.LabelFrame(self.bill_frame)
                self.bill_dic[service] = {
                    "Service": tk.Entry(self.bill_frames[service],
                                        textvariable=details[service]["Service"],
                                        width=30,
                                        font=self.conf["font"],
                                        fg="blue"),
                    "Fee": tk.Label(self.bill_frames[service],
                                    textvariable=details[service]["Fee"],
                                    width=10,
                                    font=self.conf["font"]),
                    "Ori": {
                        "Fee": tk.Entry(self.bill_frames[service],
                                        textvariable=details[service]["Ori"], width=10, font=self.conf["font"], fg="blue"),
                        "Upload": tk.Button(self.bill_frames[service],
                                            command=lambda: self.upload_file(bill_description=details[service]["Service"].get(),
                                                                             origin="Origin"), width=10, text="Upload", bg="brown", fg="white")
                    },
                    "V1": {
                        "Fee": tk.Entry(self.bill_frames[service],
                                        textvariable=details[service]["V1"], width=10, font=self.conf["font"], fg="blue"),
                        "Upload": tk.Button(self.bill_frames[service],
                                            command=lambda: self.upload_file(bill_description=details[service]["Service"].get(),
                                                                             origin="Sub Variation 1"), width=10, text="Upload", bg="brown", fg="white")
                    },
                    "V2": {
                        "Fee": tk.Entry(self.bill_frames[service],
                                        textvariable=details[service]["V2"], width=10, font=self.conf["font"], fg="blue"),
                        "Upload": tk.Button(self.bill_frames[service],
                                            command=lambda: self.upload_file(bill_description=details[service]["Service"].get(),
                                                                             origin="Sub Variation 2"), width=10, text="Upload", bg="brown", fg="white")
                    },
                    "V3": {
                        "Fee": tk.Entry(self.bill_frames[service],
                                        textvariable=details[service]["V3"], width=10, font=self.conf["font"], fg="blue"),
                        "Upload": tk.Button(self.bill_frames[service],
                                            command=lambda: self.upload_file(bill_description=details[service]["Service"].get(),
                                                                             origin="Sub Variation 3"), width=10, text="Upload", bg="brown", fg="white")
                    },
                    "Expand": [tk.LabelFrame(self.bill_frame) for _ in range(self.conf["n_items"])]

                }
                upload_file_fuc = lambda i: lambda : self.upload_file(bill_number=details[service]["Content"][i]["Number"].get(),
                                                                      bill_description=details[service]["Content"][i]["Service"].get())
                self.bill_dic[service]["Content"] = [
                    {
                        "Number": tk.Entry(self.bill_dic[service]["Expand"][i],
                                           textvariable=details[service]["Content"][i]["Number"],
                                           width=11,
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
                                        fg="blue"),
                        "in.GST": tk.Label(self.bill_dic[service]["Expand"][i],
                                           width=10,
                                           textvariable=details[service]["Content"][i]["in.GST"],
                                           font=self.conf["font"]),
                        "Upload": tk.Button(self.bill_dic[service]["Expand"][i],
                                            command=upload_file_fuc(i),
                                            width=8,
                                            text="Upload",
                                            bg="brown",
                                            fg="white")
                    } for i in range(self.conf["n_items"]-1)
                ]

                self.bill_dic[service]["Paid"] = tk.Label(self.bill_dic[service]["Expand"][-1],
                                                          textvariable=details[service]["Paid"],
                                                          width=10,
                                                          font=self.conf["font"])

                self.bill_dic[service]["Service"].grid(row=0, column=0)
                self.bill_dic[service]["Ori"]["Fee"].grid(row=0, column=1)
                self.bill_dic[service]["Ori"]["Upload"].grid(row=0, column=2)
                self.bill_dic[service]["V1"]["Fee"].grid(row=0, column=3)
                self.bill_dic[service]["V1"]["Upload"].grid(row=0, column=4)
                self.bill_dic[service]["V2"]["Fee"].grid(row=0, column=5)
                self.bill_dic[service]["V2"]["Upload"].grid(row=0, column=6)
                self.bill_dic[service]["V3"]["Fee"].grid(row=0, column=7)
                self.bill_dic[service]["V3"]["Upload"].grid(row=0, column=8)
                self.bill_dic[service]["Fee"].grid(row=0, column=9)
                self.bill_dic[service]["Paid"].pack(side=tk.RIGHT)
                for i in range(self.conf["n_items"]-1):
                    self.bill_dic[service]["Content"][i]["Number"].grid(row=0, column=0, padx=(0, 3))
                    self.bill_dic[service]["Content"][i]["Service"].grid(row=0, column=1)
                    self.bill_dic[service]["Content"][i]["Fee"].grid(row=0, column=2)
                    self.bill_dic[service]["Content"][i]["in.GST"].grid(row=0, column=3)
                    self.bill_dic[service]["Content"][i]["Upload"].grid(row=0, column=4)


                ist_update_fuc = lambda i: lambda a, b, c: self.app._ist_update(details[service]["Content"][i]["Fee"], details[service]["Content"][i]["in.GST"])

                details[service]["Fee"].trace("w", lambda a, b, c: self.app._ist_update(details[service]["Fee"], details[service]["in.GST"]))
                for i in range(self.conf["n_items"]-1):
                    details[service]["Content"][i]["Fee"].trace("w", ist_update_fuc(i))
                    details[service]["Content"][i]["Fee"].trace("w", lambda a, b, c: self.app._sum_update([details[service]["Content"][j]["Fee"] for j in range(self.conf["n_items"]-1)], details[service]["Paid"]))
                extra_list = ["Ori", "V1", "V2", "V3"]

                for extra in extra_list:
                    details[service][extra].trace("w", lambda a, b, c: self.app._sum_update([details[service][e] for e in extra_list], details[service]["Fee"]))
            self.bill_frames[service].pack(fill=tk.X)
            for i in range(self.conf["n_items"]):
                self.bill_dic[service]["Expand"][i].pack(fill=tk.X)

        else:
            archive[service] = details.pop(service)
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
        invoices_details = self.data["Invoices"]["Details"]
        bills_details = self.data["Bills"]["Details"]
        profits_details = self.data["Profits"]["Details"]
        archive = self.data["Archive"]["Profit"]
        variation = self.data["Variation"]
        service = var["Service"].get()
        include = var["Include"].get()
        if include:
            if service in archive:
                profits_details[service] = archive.pop(service)
            else:
                self.profit_frames[service] = tk.LabelFrame(self.profit_frame)
                profits_details[service] = {
                    "Fee": tk.StringVar(),
                    "in.GST": tk.StringVar(),
                    "Context": [
                        {
                            "Fee": tk.StringVar(),
                            "in.GST": tk.StringVar(),
                        } for _ in range(self.conf["n_items"])
                    ]
                }

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


                invoices_details[service]["Fee"].trace("w", lambda a, b, c: self.app._minus_update(invoices_details[service]["Fee"],
                                                                                                   bills_details[service]["Fee"], profits_details[service]["Fee"]))
                bills_details[service]["Fee"].trace("w", lambda a, b, c: self.app._minus_update(
                    invoices_details[service]["Fee"], bills_details[service]["Fee"], profits_details[service]["Fee"]))

                invoices_details[service]["in.GST"].trace("w", lambda a, b, c: self.app._minus_update(invoices_details[service]["in.GST"], bills_details[service]["in.GST"], profits_details[service]["in.GST"]))
                bills_details[service]["in.GST"].trace("w", lambda a, b, c: self.app._minus_update(
                    invoices_details[service]["in.GST"], bills_details[service]["in.GST"], profits_details[service]["in.GST"]))

                profits_details[service]["Fee"].trace("w", lambda a, b, c: self.app._sum_update([service["Fee"] for service in profits_details.values()] + [var["Fee"] for var in variation], self.data["Profits"]["Fee"]))
                profits_details[service]["in.GST"].trace("w", lambda a, b, c: self.app._sum_update([service["in.GST"] for service in profits_details.values()] + [var["in.GST"] for var in variation], self.data["Profits"]["in.GST"]))
                # self.data["Invoices"]["Fee"].trace("w", lambda a, b, c: )
                for i in range(self.conf["n_items"]):
                    tk.Label(self.profit_dic[service]["Expand"][i], text="").pack(pady=2)
                self.profit_dic[service]["Fee"].pack(pady=1)
            self.profit_frames[service].pack()
            for i in range(self.conf["n_items"]):
                self.profit_dic[service]["Expand"][i].pack(fill=tk.X)
            self.app._sum_update(
                [service["Fee"] for service in profits_details.values()] + [var["Fee"] for var in variation],
                self.data["Profits"]["Fee"])
            # self.profit_dic[service]["in.GST"].grid(row=1, column=0)

        else:
            self.profit_frames[service].pack_forget()
            for i in range(self.conf["n_items"]):
                self.profit_dic[service]["Expand"][i].pack_forget()
            archive[service] = profits_details.pop(service)
            self.app._sum_update(
                [service["Fee"] for service in profits_details.values()] + [var["Fee"] for var in variation],
                self.data["Profits"]["Fee"])
    def fee_acceptance_part(self):
        fee_acceptance_frame = tk.LabelFrame(self.main_context_frame, text="Update Fee Acceptance", font=self.conf["font"])
        fee_acceptance_frame.grid(row=1, column=1, sticky="news", columnspan=2)
        tk.Label(fee_acceptance_frame, text="Fee Acceptance", font=self.conf["font"]).grid(row=0, column=0)
        tk.Label(fee_acceptance_frame, text="Rev1", font=self.conf["font"]).grid(row=0, column=1)
        tk.Button(fee_acceptance_frame, text="Upload", font=self.conf["font"], bg="brown", fg="white",
                  command=lambda: self.upload_file(fee_acceptance="Fee Acceptance Rev1")).grid(row=0, column=2)
        tk.Label(fee_acceptance_frame, text="Rev2", font=self.conf["font"]).grid(row=0, column=3)
        tk.Button(fee_acceptance_frame, text="Upload", font=self.conf["font"], bg="brown", fg="white",
                  command=lambda: self.upload_file(fee_acceptance="Fee Acceptance Rev2")).grid(row=0, column=4)
        tk.Label(fee_acceptance_frame, text="Rev3", font=self.conf["font"]).grid(row=0, column=5)
        tk.Button(fee_acceptance_frame, text="Upload", font=self.conf["font"], bg="brown", fg="white",
                  command=lambda: self.upload_file(fee_acceptance="Fee Acceptance Rev3")).grid(row=0, column=6)
    def _on_mousewheel(self, event):
        self.main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def reset_scrollregion(self, event):
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))

    def update_invoice_sum(self, details, invoice_list):
        fee_value = {
            "INV1": 0,
            "INV2": 0,
            "INV3": 0,
            "INV4": 0,
            "INV5": 0,
            "INV6": 0,
        }
        for service in details.values():
            if service["Expand"].get():
                for i in range(self.conf["n_items"]):
                    if service["Content"][i]["Number"].get() != "None":
                        index = service["Content"][i]["Number"].get()
                        try:
                            fee_value[index] += round(float(service["Content"][i]["Fee"].get()), 2)
                        except ValueError:
                            continue
            else:
                if service["Number"].get() != "None":
                    index = service["Number"].get()
                    try:
                        fee_value[index] += round(float(service["Fee"].get()), 2)
                    except ValueError:
                        continue
        for variation in self.data["Variation"]:
            if variation["Number"].get() != "None":
                index = variation["Number"].get()
                try:
                    fee_value[index] += round(float(variation["Fee"].get()), 2)
                except ValueError:
                    continue
        for i in range(6):
            invoice_list[f"INV{str(i+1)}"]["Fee"].set(str(fee_value[f"INV{str(i+1)}"]))

    def update_fee_label(self, fee, label):
        if len(fee.get().strip())==0:
            label.config(text="$0/$0.0")
            return
        else:
            fee = fee.get()
        try:
            ingst = str(round(float(fee)*self.conf["tax rates"], 2))
            label.config(text=f"${fee}/${ingst}")
        except:
            label.config(text="Error")

    def update_dollar_sign(self, fee, label):
        if len(fee.get()) == 0:
            label.config(text="")
            return
        else:
            fee = fee.get()
        label.config(text=f"${fee}")

    def generate_invoice_number(self):
        if not self.data["State"]["Fee Accepted"].get():
            messagebox.showerror("Error", "You need to upload a fee acceptance first")
            return

        generate = messagebox.askyesno("Warming", "Are you sure you want to generate the invoice number?")
        if generate:
            inv = None
            for key, value in self.data["Financial Panel"]["Invoice Details"].items():
                if len(value["Number"].get()) == 0:
                    inv = key
                    break
            if inv is None:
                messagebox.showerror("Error", "You cant generate more than 6 invoices")
                return

            current_inv_number = self._get_current_invoice_number()
            self.data["Financial Panel"]["Invoice Details"][inv]["Number"].set(current_inv_number)
            if inv == "INV1":
                change_asana_quotation(self.app, current_inv_number)
                change_quotation_number(self.app, current_inv_number)
                self.data["Project Info"]["Project"]["Quotation Number"].set(current_inv_number)
                self.data["Project Info"]["Project"]["Project Number"].set(current_inv_number)
                messagebox.showinfo("updated", "folder name and asana updated")

    def _get_current_invoice_number(self):
        invoice_dir = os.path.join(self.conf["database_dir"], "invoices.json")
        invoices = json.load(open(invoice_dir))
        current_inv_number = list(invoices)[-1]
        today = date.today().strftime("%y%m")[1:]
        res = str(int(current_inv_number)+1) if current_inv_number.startswith(today) else today + "000"
        invoices[res] = "Backlog"
        with open(invoice_dir, "w", encoding='utf-8') as f:
            json.dump(invoices, f, ensure_ascii=False, indent=4)
        return res

    def upload_file(self, **kwargs):
        filename_list = []
        if "bill_number" in kwargs.keys():
            if len(self.data["Financial Panel"]["Invoice Details"]["INV1"]["Number"].get()) ==0:
                messagebox.showerror("Error", "Please generate first invoice number before you upload bill file")
                return
            filename_list.append(self.data["Financial Panel"]["Invoice Details"]["INV1"]["Number"].get())

        for key, value in kwargs.items():
            if len(value.strip()) == 0:
                messagebox.showerror("Warming", f"You need to enter {key.replace('_', ' ')} before you upload the file")
                return
            filename_list.append(value)
        filename = "-".join(filename_list)
        database_dir = os.path.join(self.conf["database_dir"], self.data["Project Info"]["Project"]["Quotation Number"].get())
        folder_dir = os.path.join(database_dir, filename)

        if len(filename.strip()) == 0:
            messagebox.showerror("Error", "Please input a file name first")
            return

        rewrite = True
        for file in os.listdir(database_dir):
            if os.path.splitext(file)[0] == filename:
                rewrite = messagebox.askyesno("Warming", "Existing folder found, do you want to rewrite")
                if not rewrite:
                    return
                break
        if rewrite:
            file = filedialog.askopenfilename()
            if file == "":
                return
            try:
                shutil.copy(file.replace("/", "\\"), folder_dir+os.path.splitext(file)[1])
            except PermissionError:
                messagebox.showerror("Error", "Please close the file before you upload it")
            except Exception as e:
                print(e)
                messagebox.showerror("Error", "Some error occurs, please contact Administrator")
            else:
                if "fee_acceptance" in kwargs.keys():
                    self.data["State"]["Fee Accepted"].set(True)
                    update_asana(self.app)
                    self.app.log.log_fee_accept_file(self.app)
                    save(self.app)
                    config_state(self.app)
                    config_log(self.app)
                messagebox.showinfo("upload", f"the file {file} has upload to database")
    # def _expand_bill(self, service):
    #     details = self.data["Invoices"]["Details"]
    #     bill_fee = self.data["Financial Panel"]["Bill Details"]
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
    #     self.data["Financial Panel"]["Bill Fee"][service]["Invoice"].set("None")
    #     for i in range(3):
    #         self.data["Financial Panel"]["Bill Fee"][service]["Content"][i]["Invoice"].set("None")

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
