import tkinter as tk
from tkinter import ttk


class InvoicePage(tk.Frame):
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
        self.data["Invoice Page"] = dict()

        self.invoice_part()
        # self.bill_part()
        # self.profit_main_frame()

    def invoice_part(self):

        #for demonstration purpose
        # tk.Button(self.invoice_frame,width=10, text="Generate Invoice", command= self.generate_invoice_number, bg="brown", fg="white",
        #           font=self.conf["font"]).grid(row=1, column=2)
        # tk.Button(self.invoice_frame, width=10, text="Login Xero", command= lambda: login_xero(), bg="brown",
        #           fg="white",
        #           font=self.conf["font"]).grid(row=1, column=3)
        # tk.Button(self.invoice_frame, width=10, text="Update Xero", command=lambda: update_xero(self.app.data, self.invoice_list), bg="brown",
        #           fg="white",
        #           font=self.conf["font"]).grid(row=1, column=4)
        invoice = dict()
        self.data["Invoice Page"]["Invoice Details"] = invoice

        self.invoice_frames= dict()
        self.invoice_dic = dict()
        self.invoice_dic["Number"] = dict()

        self.invoice_frame = tk.LabelFrame(self.main_context_frame, text="Invoice Details", font=self.conf["font"])
        self.invoice_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

        title_frame = tk.LabelFrame(self.invoice_frame)
        title_frame.pack()
        tk.Label(title_frame, width=35, text="Items", font=self.conf["font"]).grid(row=0, column=0)
        tk.Label(title_frame, width=10, text="ex.GST", font=self.conf["font"]).grid(row=0, column=1)
        tk.Label(title_frame, width=10, text="in.GST", font=self.conf["font"]).grid(row=1, column=1)
        invoice_number_frame = tk.LabelFrame(self.invoice_frame)
        invoice_number_frame.pack()
        tk.Label(invoice_number_frame, width=10, text="", font=self.conf["font"]).grid(row=0, column=1)
        tk.Label(invoice_number_frame, width=35, text="Invoice Number", font=self.conf["font"]).grid(row=0, column=0)
        for i in range(6):
            invoice[f"INV{str(i+1)}"] = {
                "Number": tk.StringVar(),
                "Fee": tk.StringVar(),
                "in.GST": tk.StringVar()
            }
            tk.Label(title_frame, width=10, text="INV"+str(i+1), font=self.conf["font"]).grid(row=0, column=i+2, sticky="ew")
            self.invoice_dic["Number"][f"INV{str(i+1)}"] = tk.Entry(invoice_number_frame, textvariable=invoice[f"INV{str(i+1)}"]["Number"], width=11, font=self.conf["font"], fg="blue")
            self.invoice_dic["Number"][f"INV{str(i + 1)}"].grid(row=0, column=i + 2, padx=(0, 5), sticky="ew")

        total_frame = tk.LabelFrame(self.invoice_frame)
        total_frame.pack(side=tk.BOTTOM)
        tk.Label(total_frame, width=35, text="Invoice Total", font=self.conf["font"]).grid(row=0, column=0)
        tk.Label(total_frame, width=10, textvariable=self.data["Fee Proposal"]["Fee Details"]["Fee"], font=self.conf["font"]).grid(row=0, column=1)
        tk.Label(total_frame, width=10, textvariable=self.data["Fee Proposal"]["Fee Details"]["in.GST"], font=self.conf["font"]).grid(row=1, column=1)
        ist_fuc = lambda i: lambda a, b, c: self.app._ist_update(invoice[f"INV{str(i+1)}"]["Fee"], invoice[f"INV{str(i+1)}"]["in.GST"])
        for i in range(6):
            tk.Label(total_frame, width=10, textvariable=invoice[f"INV{str(i+1)}"]["Fee"], font=self.conf["font"]).grid(row=0, column=2+i)
            tk.Label(total_frame, width=10, textvariable=invoice[f"INV{str(i+1)}"]["in.GST"], font=self.conf["font"]).grid(row=1, column=2+i)
            invoice[f"INV{str(i+1)}"]["Fee"].trace("w", ist_fuc(i))


    def bill_part(self):
        bill_details = dict()
        self.data["Invoice Page"]["Bill Details"] = bill_details
        bill_fee = dict()
        self.data["Invoice Page"]["Bill Fee"] = bill_fee
        self.bill_dic = dict()
        self.bill_dic["Number"] = dict()
        self.bill_frames = dict()

        self.bill_frame = tk.LabelFrame(self.main_context_frame, text="Bill Details", font=self.conf["font"])
        self.bill_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

        title_frame = tk.LabelFrame(self.bill_frame)
        title_frame.pack()
        tk.Label(title_frame, width=35, text="Bills", font=self.conf["font"]).grid(row=0, column=0)
        tk.Label(title_frame, width=10, text="ex.GST", font=self.conf["font"]).grid(row=0, column=1)
        tk.Label(title_frame, width=10, text="in.GST", font=self.conf["font"]).grid(row=1, column=1)

        bill_number_frame = tk.LabelFrame(self.bill_frame)
        bill_number_frame.pack()
        tk.Label(bill_number_frame, width=10, text="", font=self.conf["font"]).grid(row=0, column=1)
        tk.Label(bill_number_frame, width=35, text="Bill Number", font=self.conf["font"]).grid(row=0, column=0)
        for i in range(6):
            bill_details[f"INV{str(i + 1)}"] = {
                "Number": tk.StringVar(),
                "Fee": tk.StringVar(),
                "in.GST": tk.StringVar()
            }
            tk.Label(title_frame, width=10, text="BILL" + str(i + 1), font=self.conf["font"]).grid(row=0, column=i + 2)
            self.bill_dic["Number"][f"INV{str(i+1)}"] = tk.Entry(bill_number_frame, width=11, textvariable=bill_details[f"INV{str(i + 1)}"]["Number"], font=self.conf["font"], fg="blue")
            self.bill_dic["Number"][f"INV{str(i+1)}"].grid(row=0, column=i + 2, padx=(0, 5), sticky="ew")

        total_frame = tk.LabelFrame(self.bill_frame)
        total_frame.pack(side=tk.BOTTOM)
        bill_fee["Fee"] = tk.StringVar()
        bill_fee["in.GST"] = tk.StringVar()
        tk.Label(total_frame, width=35, text=" Bill Total", font=self.conf["font"]).grid(row=0, column=0)
        tk.Label(total_frame, width=10, textvariable=bill_fee["Fee"], font=self.conf["font"]).grid(row=0, column=1)
        tk.Label(total_frame, width=10, textvariable=bill_fee["in.GST"], font=self.conf["font"]).grid(row=1, column=1)
        ist_fuc = lambda i: lambda a, b, c: self.app._ist_update(bill_details[f"INV{str(i+1)}"]["Fee"], bill_details[f"INV{str(i+1)}"]["in.GST"])
        for i in range(6):
            tk.Label(total_frame, width=10, textvariable=bill_details[f"INV{str(i+1)}"]["Fee"], font=self.conf["font"]).grid(row=0, column=2+i)
            tk.Label(total_frame, width=10, textvariable=bill_details[f"INV{str(i+1)}"]["in.GST"], font=self.conf["font"]).grid(row=1, column=2+i)
            bill_details[f"INV{str(i+1)}"]["Fee"].trace("w", ist_fuc(i))
    # def profit_main_frame(self):
    #     self.profit_frame = tk.LabelFrame(self.main_context_frame, text="Profit Details", font=self.conf["font"])
    #     self.profit_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
    #
    #     self.profit_title_frame = tk.LabelFrame(self.profit_frame)
    #     tk.Label(self.profit_title_frame, width=10, text="ex.GST", font=self.conf["font"]).grid(row=0, column=0)
    #     tk.Label(self.profit_title_frame, width=10, text="in.GST", font=self.conf["font"]).grid(row=1, column=0)
    #     self.profit_title_frame.grid(row=0, column=0, sticky="ew")
    #
    #     _ = tk.LabelFrame(self.profit_frame)
    #     tk.Label(_, text="").grid(row=0, column=0, pady=(0, 3))
    #     _.grid(row=1, column=0, sticky="ew")
    #
    #     self.profit_data_dic = dict()
    #     self.profit_dic = dict()
    #     self.profits_frame = dict()
    #     self.profit_frame_number = 2
    #
    #     self.total_profit_frame = tk.LabelFrame(self.profit_frame)
    #     self.total_profit_var = tk.StringVar(name="Total Profit", value = "")
    #     self.total_in_profit_var = tk.StringVar(name="Total ingst Profit", value="")
    #     tk.Label(self.total_profit_frame, width=10, textvariable=self.total_profit_var, font=self.conf["font"]).grid(row=0, column=0)
    #     tk.Label(self.total_profit_frame, width=10, textvariable=self.total_in_profit_var, font=self.conf["font"]).grid(row=1, column=0)
    #
    #     self.total_profit_frame.grid(row=999, column=0, sticky="ew")
    def update_fee(self, var):
        details = self.data["Fee Proposal"]["Fee Details"]["Details"]
        invoice_details = self.data["Invoice Page"]["Invoice Details"]
        service = var["Service"].get()
        include = var["Include"].get()
        if include:
            self.invoice_frames[service] = tk.LabelFrame(self.invoice_frame)
            self.invoice_frames[service].pack()
            self.invoice_dic[service] = {
                "Service": tk.Label(self.invoice_frames[service],
                                    text=service + " design and documentation",
                                    width=35,
                                    font=self.conf["font"]),
                "Fee": tk.Label(self.invoice_frames[service],
                                textvariable=self.app.data["Fee Proposal"]["Fee Details"]["Details"][service]["Fee"],
                                width=10,
                                font=self.conf["font"]),
                "in.GST": tk.Label(self.invoice_frames[service],
                                   textvariable=self.app.data["Fee Proposal"]["Fee Details"]["Details"][service]["in.GST"],
                                   width=10,
                                   font=self.conf["font"]),
                "Invoice": [tk.Radiobutton(self.invoice_frames[service],
                                           width=8,
                                           variable=self.app.data["Fee Proposal"]["Fee Details"]["Details"][service]["Invoice"],
                                           value="INV" + str(i + 1)) for i in range(6)],
                "Content": {
                    "Details": [
                        {
                            "Service": tk.Label(self.invoice_frames[service],
                                                width=35,
                                                font=self.conf["font"],
                                                textvariable=details[service]["Content"][i]["Service"]),
                            "Fee": tk.Label(self.invoice_frames[service],
                                            width=10,
                                            font=self.conf["font"],
                                            textvariable=details[service]["Content"][i]["Fee"]),
                            "in.GST": tk.Label(self.invoice_frames[service],
                                               width=10,
                                               textvariable=details[service]["Content"][i]["in.GST"],
                                               font=self.conf["font"]),
                            "Invoice": [tk.Radiobutton(self.invoice_frames[service],
                                                       width=8,
                                                       variable=self.app.data["Fee Proposal"]["Fee Details"]["Details"][service]["Content"][i]["Invoice"],
                                                       value="INV" + str(j + 1)) for j in range(6)]
                        } for i in range(self.conf["n_items"])
                    ],
                    "Service": tk.Label(self.invoice_frames[service], width=35, text=service + " Total",
                                        font=self.conf["font"]),
                    "Fee": tk.Label(self.invoice_frames[service],
                                    width=10,
                                    textvariable=details[service]["Fee"],
                                    font=self.conf["font"]),
                    "in.GST": tk.Label(self.invoice_frames[service],
                                       width=10,
                                       textvariable=details[service]["in.GST"],
                                       font=self.conf["font"])
                }
            }
            self.invoice_dic[service]["Service"].grid(row=0, column=0)
            self.invoice_dic[service]["Fee"].grid(row=0, column=1)
            self.invoice_dic[service]["in.GST"].grid(row=1, column=1)
            for i in range(6):
                self.invoice_dic[service]["Invoice"][i].grid(row=0, column=2+i, rowspan=2, padx=(2, 0))


            details[service]["Expand"].trace("w",lambda a, b, c: self._expand_invoice(service))

            details[service]["Fee"].trace("w", lambda a, b, c: self.update_invoice_sum(details, invoice_details))
            details[service]["Invoice"].trace("w", lambda a, b, c: self.update_invoice_sum(details, invoice_details))
            for i in range(self.conf["n_items"]):
                details[service]["Content"][i]["Fee"].trace("w", lambda a, b, c: self.update_invoice_sum(details, invoice_details))
                details[service]["Content"][i]["Invoice"].trace("w", lambda a, b, c: self.update_invoice_sum(details, invoice_details))
        else:
            self.invoice_frames[service].destroy()
    
    def update_bill(self, var):
        bill_fee = self.data["Invoice Page"]["Bill Fee"]
        details = self.data["Fee Proposal"]["Fee Details"]["Details"]
        bill_details = self.data["Invoice Page"]["Bill Details"]
        service = var["Service"].get()
        include = var["Include"].get()
        if include:
            self.bill_frames[service] = tk.LabelFrame(self.bill_frame)
            self.bill_frames[service].pack(fill='x')
            bill_fee[service] = {
                "Service": tk.StringVar(),
                "Fee": tk.StringVar(),
                "in.GST": tk.StringVar(),
                "Expand": tk.BooleanVar(),
                "Invoice": tk.StringVar(value="None"),
                "Content":[
                    {
                        "Service":tk.StringVar(),
                        "Fee":tk.StringVar(),
                        "in.GST":tk.StringVar(),
                        "Invoice": tk.StringVar(value="None")
                    } for _ in range(self.conf["n_items"])
                ]
            }
            self.bill_dic[service] = {
                "Service": tk.Entry(self.bill_frames[service],
                                    width=35,
                                    font=self.conf["font"],
                                    fg="blue"),
                "Fee": tk.Entry(self.bill_frames[service],
                                textvariable=bill_fee[service]["Fee"],
                                width=10,
                                font=self.conf["font"],
                                fg="blue"),
                "in.GST": tk.Label(self.bill_frames[service],
                                   textvariable=bill_fee[service]["in.GST"],
                                   width=10,
                                   font=self.conf["font"]),
                "Invoice":[tk.Radiobutton(self.bill_frames[service],
                                          width=8,
                                          variable=bill_fee[service]["Invoice"],
                                          value="INV"+str(i+1)) for i in range(6)],
                "Content": {
                    "Details": [
                        {
                            "Service": tk.Entry(self.bill_frames[service],
                                                width=35,
                                                font=self.conf["font"],
                                                textvariable=bill_fee[service]["Content"][i]["Service"]),
                            "Fee": tk.Entry(self.bill_frames[service],
                                            width=10,
                                            font=self.conf["font"],
                                            textvariable=bill_fee[service]["Content"][i]["Fee"]),
                            "in.GST": tk.Label(self.bill_frames[service],
                                               width=10,
                                               textvariable=bill_fee[service]["Content"][i]["in.GST"],
                                               font=self.conf["font"]),
                            "Invoice": [tk.Radiobutton(self.bill_frames[service],
                                                       width=8,
                                                       variable=bill_fee[service]["Content"][i]["Invoice"],
                                                       value="INV" + str(j + 1)) for j in range(6)]
                        } for i in range(self.conf["n_items"])
                    ],
                    "Service": tk.Label(self.bill_frames[service], width=35, text=service + " Total",
                                        font=self.conf["font"]),
                    "Fee": tk.Label(self.bill_frames[service],
                                    width=10,
                                    textvariable=bill_fee[service]["Fee"],
                                    font=self.conf["font"]),
                    "in.GST": tk.Label(self.bill_frames[service],
                                       width=10,
                                       textvariable=bill_fee[service]["in.GST"],
                                       font=self.conf["font"])
                }
            }
            self.bill_dic[service]["Service"].grid(row=0, column=0, pady=(2, 0))
            self.bill_dic[service]["Fee"].grid(row=0, column=1)
            self.bill_dic[service]["in.GST"].grid(row=1, column=1)
            for i in range(6):
                self.bill_dic[service]["Invoice"][i].grid(row=0, column=2+i, rowspan=2, padx=(2, 0))
            
            details[service]["Expand"].trace("w",lambda a, b, c: self._expand_bill(service))


            bill_fee[service]["Fee"].trace("w", lambda a, b, c: self.update_invoice_sum(bill_fee, bill_details))
            bill_fee[service]["Invoice"].trace("w", lambda a, b, c: self.update_invoice_sum(bill_fee, bill_details))
            for i in range(self.conf["n_items"]):
                bill_fee[service]["Content"][i]["Fee"].trace("w", lambda a, b, c: self.update_invoice_sum(bill_fee, bill_details))
                bill_fee[service]["Content"][i]["Invoice"].trace("w", lambda a, b, c: self.update_invoice_sum(bill_fee, bill_details))
            # #trace function start
            #
            # self.bill_data_dic[var._name]["Fee"].trace("w", lambda a,b,c:self.app._ist_update(
            #     self.bill_data_dic[var._name]["Fee"], self.bill_data_dic[var._name]["in.GST"]))
            #
            # self.bill_data_dic[var._name]["Fee"].trace("w", lambda a,b,c: self.app._sum_update(
            #     [value["Fee"] for value in self.bill_data_dic.values()], self.bill_fee_total))
            #
            # self.bill_data_dic[var._name]["in.GST"].trace("w", lambda a,b,c: self.app._sum_update(
            #     [value["in.GST"] for value in self.bill_data_dic.values()], self.bill_fee_isgst_total))
            #
            #
            #
            # #garbar collect need to fix
            # ist_fun_0=lambda a, b, c: self.app._ist_update(
            #     self.bill_data_dic[var._name]["Context"][0]["Fee"],
            #     self.bill_data_dic[var._name]["Context"][0]["in.GST"])
            # self.bill_data_dic[var._name]["Context"][0]["Fee"].trace("w", ist_fun_0)
            # ist_fun_1=lambda a, b, c: self.app._ist_update(
            #     self.bill_data_dic[var._name]["Context"][1]["Fee"],
            #     self.bill_data_dic[var._name]["Context"][1]["in.GST"])
            # self.bill_data_dic[var._name]["Context"][1]["Fee"].trace("w", ist_fun_1)
            # ist_fun_2=lambda a, b, c: self.app._ist_update(
            #     self.bill_data_dic[var._name]["Context"][2]["Fee"],
            #     self.bill_data_dic[var._name]["Context"][2]["in.GST"])
            # self.bill_data_dic[var._name]["Context"][2]["Fee"].trace("w", ist_fun_2)
            # for i in range(3):
            #     sum_fun = lambda a, b, c: self.app._sum_update([item["Fee"] for item in self.bill_data_dic[var._name]["Context"]],
            #                                                    self.bill_data_dic[var._name]["Fee"])
            #     self.bill_data_dic[var._name]["Context"][i]["Fee"].trace("w", sum_fun)
            #
            #
            # self.bill_data_dic[var._name]["Invoice"].trace("w", lambda a,b,c:self.app.update_invoice_sum(self.bill_data_dic, self.bill_list))
            # self.bill_data_dic[var._name]["Fee"].trace("w", lambda a,b,c:self.app.update_invoice_sum(self.bill_data_dic, self.bill_list))
            #
            # #garbage collect
            # self.bill_data_dic[var._name]["Context"][0]["Service"].trace("w", lambda a,b,c: self.update_bill_radiobutton(self.bill_data_dic[var._name]["Context"][0]["Service"]))
            # self.bill_data_dic[var._name]["Context"][1]["Service"].trace("w", lambda a,b,c: self.update_bill_radiobutton(self.bill_data_dic[var._name]["Context"][1]["Service"]))
            # self.bill_data_dic[var._name]["Context"][2]["Service"].trace("w", lambda a,b,c: self.update_bill_radiobutton(self.bill_data_dic[var._name]["Context"][2]["Service"]))
            #
            # for i in range(3):
            #     self.bill_data_dic[var._name]["Context"][i]["Invoice"].trace("w", lambda a, b,c: self.app.update_invoice_sum(self.bill_data_dic, self.bill_list))
            #
            #
            # for i in range(3):
            #     self.bill_dic[var._name]["expand frame"][i]["Service"].grid(row=0, column=0, padx=20, pady=(2, 0))
            #     self.bill_dic[var._name]["expand frame"][i]["Fee"].grid(row=0, column=1)
            #     self.bill_dic[var._name]["expand frame"][i]["in.GST"].grid(row=1, column=1, padx=(0, 2))
            # self.bill_dic[var._name]["Total"].grid(row=0, column=0)
            # self.bill_dic[var._name]["Total Fee"].grid(row=0, column=1)
            # self.bill_dic[var._name]["Total in.GST"].grid(row=1, column=1)
            #
            # self.bill_dic[var._name]["Service"].grid(row=0, column=0, padx=20, pady=(2, 0))
            # self.bill_dic[var._name]["Fee"].grid(row=0, column=1)
            # self.bill_dic[var._name]["in.GST"].grid(row=1, column=1, padx=(0, 2))
            #
            # for i in range(6):
            #     self.bill_dic[var._name]["Invoice"][i].grid(row=0, column=3+i, padx=(2, 0), rowspan=2)
            # self.bill_dic[var._name]["Position"] = str(self.bill_frame_number)
            # self.bills_frame[var._name].grid(row=self.bill_frame_number, column=0, sticky="ew")
            # self.bill_frame_number += 5
        else:
            self.bill_frames[service].destroy()
            bill_fee[service]["Fee"].set("")
            bill_fee.pop(service)
            # self.bill_data_dic[var._name]["Fee"].set("")
            # if self.bill_data_dic[var._name]["Expanded"].get():
            #     self.bill_data_dic[var._name]["Expanded"].set(False)
            # self.bills_frame[var._name].grid_forget()
    
    def update_profit(self, var):
        if var.get() == True:
            if not var._name in self.profit_dic.keys():
                self.profits_frame[var._name] = tk.LabelFrame(self.profit_frame)
                
                self.profit_data_dic[var._name] = {
                    "on":tk.BooleanVar(name=var._name+" profit on", value=True),
                    "Expanded":tk.BooleanVar(name=var._name + " profit Expand", value=False),
                    "Fee":tk.StringVar(name=var._name + " profit Fee"),
                    "in.GST":tk.StringVar(name=var._name + " profit in.GST"),
                    "Context":[
                        {
                            "Fee":tk.StringVar(name=var._name+" profit Fee "+str(i)),
                            "in.GST":tk.StringVar(name=var._name+" profit in.GST "+str(i)),
                        } for i in range(3)
                    ]
                }
                
                self.profit_dic[var._name] = {
                    "Fee": tk.Label(self.profits_frame[var._name],
                                    textvariable=self.profit_data_dic[var._name]["Fee"],
                                    width=10,
                                    font=self.conf["font"]),
                    "in.GST": tk.Label(self.profits_frame[var._name],
                                       textvariable=self.profit_data_dic[var._name]["in.GST"],
                                       width=10,
                                       font=self.conf["font"]),
                    "Position": 0,
                    "expand frames": [tk.LabelFrame(self.profit_frame) for _ in range(4)]
                }
                self.profit_dic[var._name]["expand frame"] = [
                    {
                        "Fee": tk.Label(self.profit_dic[var._name]["expand frames"][i],
                                        width=10,
                                        textvariable=self.profit_data_dic[var._name]["Context"][i]["Fee"],
                                        font=self.conf["font"]),
                        "in.GST": tk.Label(self.profit_dic[var._name]["expand frames"][i],
                                           width=10,
                                           textvariable=self.profit_data_dic[var._name]["Context"][i]["in.GST"],
                                           font=self.conf["font"])
                    } for i in range(3)
                ]
                self.profit_dic[var._name]["Total Fee"]=tk.Label(self.profit_dic[var._name]["expand frames"][3], width=10,
                                                              textvariable=self.profit_data_dic[var._name]["Fee"],
                                                              font=self.conf["font"])
                self.profit_dic[var._name]["Total in.GST"]=tk.Label(self.profit_dic[var._name]["expand frames"][3], width=10,
                                                                 textvariable=self.profit_data_dic[var._name][
                                                                     "in.GST"],
                                                                 font=self.conf["font"])
                ###
                #trace function here
                self.app.data["Fee Proposal Page"]["Details"][var._name]["Expanded"].trace("w", lambda a, b, c: self.profit_expand(var._name))
                self.app.data["Fee Proposal Page"]["Details"][var._name]["Fee"].trace("w", self._update_profit)
                self.bill_data_dic[var._name]["Fee"].trace("w", self._update_profit)
                ###
                for i in range(3):
                    self.profit_dic[var._name]["expand frame"][i]["Fee"].grid(row=0, column=0)
                    self.profit_dic[var._name]["expand frame"][i]["in.GST"].grid(row=1, column=0, padx=(0, 2))
                self.profit_dic[var._name]["Total Fee"].grid(row=0, column=0)
                self.profit_dic[var._name]["Total in.GST"].grid(row=1, column=0)

                self.profit_dic[var._name]["Fee"].grid(row=0, column=0)
                self.profit_dic[var._name]["in.GST"].grid(row=1, column=0, padx=(0, 2))

            self.profit_dic[var._name]["Position"] = str(self.profit_frame_number)
            self.profits_frame[var._name].grid(row=self.profit_frame_number, column=0, sticky="ew")
            self.profit_frame_number += 5
        else:
            self.profit_data_dic[var._name]["Fee"].set("")
            if self.profit_data_dic[var._name]["Expanded"].get():
                self.profit_data_dic[var._name]["Expanded"].set(False)
            self.profits_frame[var._name].grid_forget()

    def _expand_invoice(self, service):
        details = self.data["Fee Proposal"]["Fee Details"]["Details"]
        if details[service]["Expand"].get():
            for i in range(self.conf["n_items"]):
                self.invoice_dic[service]["Content"]["Details"][i]["Service"].grid(row=2*i+2, column=0)
                self.invoice_dic[service]["Content"]["Details"][i]["Fee"].grid(row=2*i+2, column=1)
                self.invoice_dic[service]["Content"]["Details"][i]["in.GST"].grid(row=2*i+3, column=1)
                for j in range(6):
                    self.invoice_dic[service]["Content"]["Details"][i]["Invoice"][j].grid(row=2*i+2, column=2+j, rowspan=2, padx=(2, 0))
            self.invoice_dic[service]["Content"]["Service"].grid(row=2*(self.conf["n_items"]+1)+1, column=0)
            self.invoice_dic[service]["Content"]["Fee"].grid(row=2*(self.conf["n_items"]+1)+1, column=1)
            self.invoice_dic[service]["Content"]["in.GST"].grid(row=2*(self.conf["n_items"]+1)+2, column=1)

            self.invoice_dic[service]["Fee"].grid_forget()
            self.invoice_dic[service]["in.GST"].grid_forget()
            for i in range(6):
                self.invoice_dic[service]["Invoice"][i].grid_forget()
        else:
            for i in range(self.conf["n_items"]):
                self.invoice_dic[service]["Content"]["Details"][i]["Service"].grid_forget()
                self.invoice_dic[service]["Content"]["Details"][i]["Fee"].grid_forget()
                self.invoice_dic[service]["Content"]["Details"][i]["in.GST"].grid_forget()
                for j in range(6):
                    self.invoice_dic[service]["Content"]["Details"][i]["Invoice"][j].grid_forget()
            self.invoice_dic[service]["Content"]["Service"].grid_forget()
            self.invoice_dic[service]["Content"]["Fee"].grid_forget()
            self.invoice_dic[service]["Content"]["in.GST"].grid_forget()

            self.invoice_dic[service]["Fee"].grid(row=0, column=1)
            self.invoice_dic[service]["in.GST"].grid(row=1, column=1)
            for i in range(6):
                self.invoice_dic[service]["Invoice"][i].grid(row=0, column=2+i, rowspan=2, padx=(2, 0))
        details[service]["Invoice"].set("None")
        for i in range(3):
            details[service]["Content"][i]["Invoice"].set("None")

    def _expand_bill(self, service):
        details = self.data["Fee Proposal"]["Fee Details"]["Details"]
        bill_fee = self.data["Invoice Page"]["Bill Details"]
        if details[service]["Expand"].get():
            bill_fee[service]["Expand"].set(True)
            for i in range(self.conf["n_items"]):
                self.bill_dic[service]["Content"]["Details"][i]["Service"].grid(row=2*i+2, column=0, pady=(2, 0))
                self.bill_dic[service]["Content"]["Details"][i]["Fee"].grid(row=2*i+2, column=1)
                self.bill_dic[service]["Content"]["Details"][i]["in.GST"].grid(row=2*i+3, column=1)
                for j in range(6):
                    self.bill_dic[service]["Content"]["Details"][i]["Invoice"][j].grid(row=2*i+2, column=2+j, rowspan=2, padx=(2, 0))
            self.bill_dic[service]["Content"]["Service"].grid(row=2*(self.conf["n_items"]+1)+1, column=0)
            self.bill_dic[service]["Content"]["Fee"].grid(row=2*(self.conf["n_items"]+1)+1, column=1)
            self.bill_dic[service]["Content"]["in.GST"].grid(row=2*(self.conf["n_items"]+1)+2, column=1)

            self.bill_dic[service]["Service"].config(state=tk.DISABLED)
            self.bill_dic[service]["Fee"].grid_forget()
            self.bill_dic[service]["in.GST"].grid_forget()
            for i in range(6):
                self.bill_dic[service]["Invoice"][i].grid_forget()
        else:
            bill_fee[service]["Expand"].set(False)
            for i in range(self.conf["n_items"]):
                self.bill_dic[service]["Content"]["Details"][i]["Service"].grid_forget()
                self.bill_dic[service]["Content"]["Details"][i]["Fee"].grid_forget()
                self.bill_dic[service]["Content"]["Details"][i]["in.GST"].grid_forget()
                for j in range(6):
                    self.bill_dic[service]["Content"]["Details"][i]["Invoice"][j].grid_forget()
            self.bill_dic[service]["Content"]["Service"].grid_forget()
            self.bill_dic[service]["Content"]["Fee"].grid_forget()
            self.bill_dic[service]["Content"]["in.GST"].grid_forget()

            self.bill_dic[service]["Service"].config(state=tk.NORMAL)
            self.bill_dic[service]["Fee"].grid(row=0, column=1)
            self.bill_dic[service]["in.GST"].grid(row=1, column=1)
            for i in range(6):
                self.bill_dic[service]["Invoice"][i].grid(row=0, column=2+i, rowspan=2, padx=(2, 0))
        self.data["Invoice Page"]["Bill Fee"][service]["Invoice"].set("None")
        for i in range(3):
            self.data["Invoice Page"]["Bill Fee"][service]["Content"][i]["Invoice"].set("None")

    def profit_expand(self, service):
        if self.app.data["Fee Proposal Page"]["Details"][service]["Expanded"].get():
            self.profit_data_dic[service]["Expanded"].set(True)
            for i in range(4):
                self.profit_dic[service]["expand frames"][i].grid(row=int(self.profit_dic[service]["Position"]) + i + 1,
                                                                  column=0, sticky="ew")
            self.profit_data_dic[service]["Fee"].set("")
            self.profits_frame[service].config(height=28)
            self.profit_dic[service]["Fee"].grid_forget()
            self.profit_dic[service]["in.GST"].grid_forget()
        else:
            self.profit_data_dic[service]["Expanded"].set(False)
            self.profits_frame[service].config(height=56)
            for i in range(4):
                self.profit_dic[service]["expand frames"][i].grid_forget()
            self.profit_dic[service]["Fee"].grid(row=0, column=1)
            self.profit_dic[service]["in.GST"].grid(row=1, column=1, padx=(0, 2))
            for j in range(3):
                self.profit_data_dic[service]["Context"][j]["Fee"].set("")


    def _update_profit(self, *args):
        Total = 0
        inGST_Total = 0
        for service_name, service in self.app.data["Fee Proposal Page"]["Details"].items():
            if not service["on"].get():
                continue
            if service["Expanded"].get():
                for i in range(3):
                    try:
                        invoice = service["Context"][i]["Fee"].get() if service["Context"][i]["Fee"].get() != "" else 0
                        bill = self.bill_data_dic[service_name]["Context"][i]["Fee"].get() if self.bill_data_dic[service_name]["Context"][i]["Fee"].get() != "" else 0
                        self.profit_data_dic[service_name]["Context"][i]["Fee"].set(int(invoice) - int(bill))
                        self.profit_data_dic[service_name]["Context"][i]["in.GST"].set(int(float(invoice)*1.1) - int(float(bill)*1.1))
                    except ValueError:
                        self.profit_data_dic[service_name]["Context"][i]["Fee"].set("Error")
                        self.profit_data_dic[service_name]["Context"][i]["in.GST"].set("Error")
            try:
                invoice = service["Fee"].get() if service["Fee"].get() != "" else 0
                bill = self.bill_data_dic[service_name]["Fee"].get() if self.bill_data_dic[service_name]["Fee"].get() != "" else 0
                self.profit_data_dic[service_name]["Fee"].set(int(invoice) - int(bill))
                self.profit_data_dic[service_name]["in.GST"].set(int(float(invoice)*1.1) - int(float(bill)*1.1))
                Total += int(invoice) - int(bill)
                inGST_Total += int(float(invoice)*1.1) - int(float(bill)*1.1)
            except ValueError:
                self.profit_data_dic[service_name]["Fee"].set("Error")
                self.profit_data_dic[service_name]["in.GST"].set("Error")
        self.total_profit_var.set(str(Total))
        self.total_in_profit_var.set(str(inGST_Total))


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
                    if service["Content"][i]["Invoice"].get() != "None":
                        index = service["Content"][i]["Invoice"].get()
                        try:
                            fee_value[index] += round(float(service["Content"][i]["Fee"].get()), 2)
                        except ValueError:
                            continue
            else:
                if service["Invoice"].get() != "None":
                    index = service["Invoice"].get()
                    try:
                        fee_value[index] += round(float(service["Fee"].get()), 2)
                    except ValueError:
                        continue
        for i in range(6):
            invoice_list[f"INV{str(i+1)}"]["Fee"].set(str(fee_value[f"INV{str(i+1)}"]))
    # def generate_invoice_number(self):
    #     wb = load_workbook(os.getcwd()+"\\PCE INV.xlsx")
    #     cur_number = max([int(num) for num in wb.sheetnames if num.isdigit()])+1
    #     self.invoice_list[0]["Invoice Number"].set(str(cur_number))
    #     self.invoice_number_entry[0].config(state=tk.DISABLED)
    #     self.update_invoice_number = str(cur_number)
    #     # sheet = wb.create_sheet(str(cur_number+1), -1)
    #     # sheet = wb.copy_worksheet(wb[str(cur_number)])
    #     # sheet = wb[str(cur_number)]
    #     # print(sheet.title)
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
