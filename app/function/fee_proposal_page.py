import tkinter as tk
from tkinter import ttk
from datetime import date

class FeeProposalPage(tk.Frame):
    def __init__(self, parent, app):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.app = app
        self.main_frame = tk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=1, side=tk.LEFT)

        self.main_canvas = tk.Canvas(self.main_frame)
        self.main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1, anchor="nw")

        self.scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.main_canvas.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y, expand=1)

        self.main_canvas.config(yscrollcommand=self.scrollbar.set)
        self.main_canvas.bind("<Configure>", lambda e:self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")))
        self.app.bind("<MouseWheel>", self._on_mousewheel, add="+")
        self.main_context_frame = tk.LabelFrame(self.main_canvas, text="Main Context")
        self.main_canvas.create_window((0, 0), window=self.main_context_frame, anchor="center")
        self.main_context_frame.bind("<Configure>", self.reset_scrollregion)
        self.reference_frame()
        self.time_frame()
        self.scope_frame()
        self.fee_frame()

    def reference_frame(self):
        reference_frame = tk.LabelFrame(self.main_context_frame, text="reference")
        reference_frame.pack(fill=tk.BOTH, expand=1, padx=20)
        self.app.data["Fee Proposal Page"] = dict()

        self.app.data["Fee Proposal Page"]["Date"] = tk.StringVar(name="Date", value=date.today().strftime("%d-%b-%Y"))
        tk.Label(reference_frame, width=30, text="Date", font=self.app.font).grid(row=0, column=0, padx=(10, 0))
        tk.Entry(reference_frame, width=70, textvariable=self.app.data["Fee Proposal Page"]["Date"], font=self.app.font).grid(row=0, column=1, padx=(0, 10))

        self.app.data["Fee Proposal Page"]["Revision"] = tk.StringVar(name="Revision", value="1")
        tk.Label(reference_frame, width=30, text="Revision", font=self.app.font).grid(row=1, column=0, padx=(10, 0))
        tk.Entry(reference_frame, width=70, textvariable=self.app.data["Fee Proposal Page"]["Revision"], font=self.app.font).grid(row=1, column=1, padx=(0, 10))

    def time_frame(self):
        time_frame = tk.LabelFrame(self.main_context_frame, text="Time Consuming", font=self.app.font)
        time_frame.pack(fill=tk.BOTH, expand=1, padx=20)

        stages = ["Fee proposal stage", "Pre-design information collection", "Documentation and issue"]
        self.app.data["Fee Proposal Page"]["Time"] = dict()

        for i, stage in enumerate(stages):
            tk.Label(time_frame, width=30,  text=stage, font=self.app.font).grid(row=i, column=0)
            self.app.data["Fee Proposal Page"]["Time"][stage+" Duration"] = (tk.StringVar(value="1"), tk.StringVar(value="2"))
            tk.Entry(time_frame, width=20, font=self.app.font, fg="blue", textvariable=self.app.data["Fee Proposal Page"]["Time"][stage+" Duration"][0]).grid(row=i, column=1)
            tk.Label(time_frame, text="-").grid(row=i, column=2)
            tk.Entry(time_frame, width=20, font=self.app.font, fg="blue", textvariable=self.app.data["Fee Proposal Page"]["Time"][stage+" Duration"][1]).grid(row=i, column=3)
            tk.Label(time_frame, text=" business day", font=self.app.font).grid(row=i, column=4)

    def scope_frame(self):
        self.scope_frame = tk.LabelFrame(self.main_context_frame, text="Scope of Work", font=self.app.font)
        self.scope_frame.pack(fill=tk.BOTH, expand=1, padx=20)
        self.current_frame_number = 0
        self.scope_frames = {}
        self.app.conn.autocommit = True
        self.convert_map = {
            "Hydraulic Service": "Hydraulic Services",
            "Mechanical Service": "Mechanical Service",
            "Electrical Service": "Electrical Services",
            "Fire Service": "Fire Protection Services"
        }
        self.append_context = dict()
        self.app.data["Fee Proposal Page"]["Scope of Work"] = {}

        # self.append_frame = tk.Frame(self.scope_frame)
        # self.append_frame.grid(row=999, column=0)

        # self.append_service = tk.StringVar()
        # service_type = ["Hydraulic Service", "Electrical Service", "Fire Service", "Kitchen Ventilation", "Mechanical Service", "CFD Service", "Mech Review",
        #                 "Installation", "Miscellaneous"]
        # ttk.Combobox(self.append_frame, width=20, textvariable=self.append_service, values=service_type, state="readonly").grid(row=0, column=0)
        #
        # self.append_extra = tk.StringVar()
        # extra_type = ["Extend", "Exclusion", "Deliverables"]
        #
        # ttk.Combobox(self.append_frame, width=20, textvariable=self.append_extra, values=extra_type, state="readonly").grid(row=0, column=1)

    def fee_frame(self):
        self.fee_frame = tk.LabelFrame(self.main_context_frame, text="Fee Proposal Details", font=self.app.font)
        self.fee_frame.pack(fill=tk.BOTH, expand=1, padx=20)


        tk.Label(self.fee_frame, text="Services", font=self.app.font).grid(row=0, column=1)
        tk.Label(self.fee_frame, text="ex.GST", font=self.app.font).grid(row=0, column=2)
        tk.Label(self.fee_frame, text="in.GST", font=self.app.font).grid(row=0, column=3)

        self.app.data["Fee Proposal Page"]["Details"] = dict()
        self.fee_dic = dict()
        self.bool_var_list = []
        self.fee_frame_number = 1

        tk.Label(self.fee_frame, text="Total", font=self.app.font).grid(row=999, column=1)
        self.total_ex_gst_label = tk.Label(self.fee_frame, text="0", font=self.app.font)
        self.total_ex_gst_label.grid(row=999, column=2)
        self.total_in_gst_label = tk.Label(self.fee_frame, text="0", font=self.app.font)
        self.total_in_gst_label.grid(row=999, column=3)

    def update_scope_frame(self, var):
        if var.get() == True:
            if self.scope_frames.get(var._name) == None:
                self.scope_frames[var._name] = dict()
                self.scope_frames[var._name]["main"] = tk.LabelFrame(self.scope_frame, text=var._name, font=self.app.font)

                self.app.data["Fee Proposal Page"]["Scope of Work"][var._name] = dict()
                self.app.data["Fee Proposal Page"]["Scope of Work"][var._name]["on"] = True

                extra_list = ["Extend", "Exclusion", "Deliverables"]
                for i, extra in enumerate(extra_list):

                    self.app.data["Fee Proposal Page"]["Scope of Work"][var._name][extra]=[]

                    self.scope_frames[var._name][extra] = tk.LabelFrame(self.scope_frames[var._name]["main"], text=extra, font=self.app.font)
                    self.scope_frames[var._name][extra].grid(row=i, column=0)
                    self.app.cur.execute(
                        f"""
                            SELECT *
                            FROM service_scope
                            WHERE service_type='{self.convert_map[var._name]}'
                            AND extra ='{extra}'
                        """
                    )
                    service = self.app.cur.fetchall()
                    context_button_list = []
                    context_entry_list = []
                    color_list = ["white", "azure"]
                    for j, context in enumerate(service):
                        self.app.data["Fee Proposal Page"]["Scope of Work"][var._name][extra].append((tk.BooleanVar(value=True), tk.StringVar(value=context[2])))
                        context_button_list.append(tk.Checkbutton(self.scope_frames[var._name][extra], variable=self.app.data["Fee Proposal Page"]["Scope of Work"][var._name][extra][j][0]))
                        context_button_list[j].grid(row=j, column=0)
                        context_entry_list.append(tk.Entry(self.scope_frames[var._name][extra], width=100, textvariable=self.app.data["Fee Proposal Page"]["Scope of Work"][var._name][extra][j][1], font=self.app.font, bg=color_list[j%2]))
                        context_entry_list[j].grid(row=j, column=1)
                self.append_context[var._name] = dict()
                # garbage collection
                self.append_context[var._name]["Extend"] = (tk.StringVar(), tk.BooleanVar())
                append_frame = tk.Frame(self.scope_frames[var._name]["Extend"])
                append_frame.grid(row=999, column=0, columnspan=2)
                tk.Entry(append_frame, width=95, textvariable=self.append_context[var._name]["Extend"][0]).grid(row=0,
                                                                                                                column=0)
                tk.Checkbutton(append_frame, variable=self.append_context[var._name]["Extend"][1],
                               text="Add to Database").grid(row=0, column=1)
                tk.Button(append_frame, text="Submit", command=lambda: self.append_value(var._name, "Extend")).grid(
                    row=0,
                    column=2)
                self.append_context[var._name]["Exclusion"] = (tk.StringVar(), tk.BooleanVar())
                append_frame = tk.Frame(self.scope_frames[var._name]["Exclusion"])
                append_frame.grid(row=999, column=0, columnspan=2)
                tk.Entry(append_frame, width=95, textvariable=self.append_context[var._name]["Exclusion"][0]).grid(
                    row=0,
                    column=0)
                tk.Checkbutton(append_frame, variable=self.append_context[var._name]["Exclusion"][1],
                               text="Add to Database").grid(row=0, column=1)
                tk.Button(append_frame, text="Submit", command=lambda: self.append_value(var._name, "Exclusion")).grid(
                    row=0, column=2)

                self.append_context[var._name]["Deliverables"] = (tk.StringVar(), tk.BooleanVar())
                append_frame = tk.Frame(self.scope_frames[var._name]["Deliverables"])
                append_frame.grid(row=999, column=0, columnspan=2)
                tk.Entry(append_frame, width=95, textvariable=self.append_context[var._name]["Deliverables"][0]).grid(
                    row=0,
                    column=0)
                tk.Checkbutton(append_frame, variable=self.append_context[var._name]["Deliverables"][1],
                               text="Add to Database").grid(row=0, column=1)
                tk.Button(append_frame, text="Submit",
                          command=lambda: self.append_value(var._name, "Deliverables")).grid(
                    row=0, column=2)
            self.scope_frames[var._name]["main"].grid(row=self.current_frame_number, column=0)
            self.current_frame_number += 1
        else:
            self.scope_frames[var._name]["main"].grid_forget()
            self.app.data["Fee Proposal Page"]["Scope of Work"][var._name]["on"] = False



    def update_fee(self, var):
        if var.get() == True:
            if not var._name in self.app.data["Fee Proposal Page"]["Details"]:
                self.app.data["Fee Proposal Page"]["Details"][var._name] = {
                    "on":tk.BooleanVar(name=var._name+" on", value=True),
                    "Expanded":tk.BooleanVar(name=var._name + " Expand", value=False),
                    "Fee":tk.StringVar(name=var._name + " Fee"),
                    "Context":[(tk.StringVar(), tk.StringVar()) for _ in range(3)]
                }

                
                self.fee_dic[var._name] ={
                    "Expand":tk.Checkbutton(self.fee_frame,
                                            text="Expand",
                                            variable=self.app.data["Fee Proposal Page"]["Details"][var._name]["Expanded"],
                                            font=self.app.font),
                    "Service":tk.Label(self.fee_frame,
                                       text=var._name + " design and documentation",
                                       width=50,
                                       font=self.app.font),
                    "ex.GST":tk.Entry(self.fee_frame,
                                      textvariable=self.app.data["Fee Proposal Page"]["Details"][var._name]["Fee"],
                                      width=20,
                                      font=self.app.font,
                                      fg="blue"),
                    "in.GST":tk.Label(self.fee_frame, text=0, width=20, font=self.app.font),
                    "Position":0,
                    "expand frame":{
                        "Service": [tk.Entry(self.fee_frame, width=50, textvariable=self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][i][0]) for i in range(3)],
                        "Fee": [tk.Entry(self.fee_frame, width=20, textvariable=self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][i][1]) for i in range(3)],
                        "in.GST":[tk.Label(self.fee_frame, width=20, text="0", font=self.app.font) for _ in range(3)],
                        "Total":tk.Label(self.fee_frame, width=50,text=var._name + " Total", font=self.app.font)
                    }
                }

                self.app.data["Fee Proposal Page"]["Details"][var._name]["Expanded"].trace("w", lambda a,b,c:self.expand(var._name))
                for i in range(3):
                    sum_fun = lambda a, b, c: self.app._sum_update([item[1] for item in self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"]],
                                                                   self.app.data["Fee Proposal Page"]["Details"][var._name]["Fee"])

                    self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][i][1].trace("w",sum_fun)
                    # self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][i][0].trace("w",lambda a,b,c: self.enable_input(var._name, i))

                #garbage collection value
                ist_fun_0 = lambda a, b, c: self.app._ist_update(
                    self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][0][1],
                    self.fee_dic[var._name]["expand frame"]["in.GST"][0])
                self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][0][1].trace("w", ist_fun_0)
                ist_fun_1 = lambda a, b, c: self.app._ist_update(
                    self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][1][1],
                    self.fee_dic[var._name]["expand frame"]["in.GST"][1])
                self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][1][1].trace("w", ist_fun_1)
                ist_fun_2 = lambda a, b, c: self.app._ist_update(
                    self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][2][1],
                    self.fee_dic[var._name]["expand frame"]["in.GST"][2])
                self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][2][1].trace("w", ist_fun_2)


                self.app.data["Fee Proposal Page"]["Details"][var._name]["Fee"].trace("w", lambda a,b,c:self.app._ist_update(self.app.data["Fee Proposal Page"]["Details"][var._name]["Fee"], self.fee_dic[var._name]["in.GST"]))
                self.app.data["Fee Proposal Page"]["Details"][var._name]["Fee"].trace("w", lambda a,b,c:self.app._sum_update([value["Fee"] for value in self.app.data["Fee Proposal Page"]["Details"].values()], self.total_ex_gst_label))
                self.app.data["Fee Proposal Page"]["Details"][var._name]["Fee"].trace("w", lambda a,b,c: self.app._in_sum_update(
                    [value["Fee"] for value in self.app.data["Fee Proposal Page"]["Details"].values()], self.total_in_gst_label))
                # sample


            self.app.data["Fee Proposal Page"]["Details"][var._name]["on"].set(True)
            self.fee_dic[var._name]["Position"]=str(self.fee_frame_number)

            self.fee_dic[var._name]["Expand"].grid(row=self.fee_dic[var._name]["Position"], column=0)
            self.fee_dic[var._name]["Service"].grid(row=self.fee_dic[var._name]["Position"], column=1)
            self.fee_dic[var._name]["ex.GST"].grid(row=self.fee_dic[var._name]["Position"], column=2)
            self.fee_dic[var._name]["in.GST"].grid(row=self.fee_dic[var._name]["Position"], column=3)
            self.fee_frame_number += 5
        else:
            self.app.data["Fee Proposal Page"]["Details"][var._name]["Fee"].set("")
            if self.app.data["Fee Proposal Page"]["Details"][var._name]["Expanded"].get():
                self.app.data["Fee Proposal Page"]["Details"][var._name]["Expanded"].set(False)
            self.fee_dic[var._name]["Expand"].grid_forget()
            self.fee_dic[var._name]["Service"].grid_forget()
            self.fee_dic[var._name]["ex.GST"].grid_forget()
            self.fee_dic[var._name]["in.GST"].grid_forget()





    def _on_mousewheel(self, event):
        self.main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")


    def append_value(self, service, extra):
        if self.append_context[service][extra][0].get() == "":
            tk.messagebox.showwarning(title="Error", message="You cant need to enter some context")
            return

        self.app.data["Fee Proposal Page"]["Scope of Work"][service][extra].append((tk.BooleanVar(value=True), tk.StringVar(value=self.append_context[service][extra][0].get())))

        tk.Entry(self.scope_frames[service][extra], width=100,
                 textvariable=self.app.data["Fee Proposal Page"]["Scope of Work"][service][extra][-1][1],
                 font=self.app.font).grid(row=len(self.app.data["Fee Proposal Page"]["Scope of Work"][service][extra]), column=1)

        tk.Checkbutton(self.scope_frames[service][extra],
                       variable=self.app.data["Fee Proposal Page"]["Scope of Work"][service][extra][-1][0])\
            .grid(row=len(self.app.data["Fee Proposal Page"]["Scope of Work"][service][extra]), column=0)

        if self.append_context[service][extra][1].get():
            self.app.cur.execute(f"""
                INSERT INTO service_scope(service_type, extra, context)
                VALUES 
                ('{self.convert_map[service]}','{extra}','{self.append_context[service][extra][0].get()}')
                """)
            self.app.conn.commit()




        self.append_context[service][extra][0].set("")
        self.append_context[service][extra][1].set(False)

    def reset_scrollregion(self, event):
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))


    def expand(self, service):
        if self.app.data["Fee Proposal Page"]["Details"][service]["Expanded"].get():
            for i in range(3):
                self.fee_dic[service]["expand frame"]["Service"][i].grid(row=int(self.fee_dic[service]["Position"]) + i + 1, column=1)
                self.fee_dic[service]["expand frame"]["Fee"][i].grid(row=int(self.fee_dic[service]["Position"]) + i + 1, column=2)
                self.fee_dic[service]["expand frame"]["in.GST"][i].grid(row=int(self.fee_dic[service]["Position"]) + i + 1, column=3)
            self.app.data["Fee Proposal Page"]["Details"][service]["Fee"].set("")
            self.fee_dic[service]["expand frame"]["Total"].grid(row=int(self.fee_dic[service]["Position"]) + 4, column=1, pady=(0,20))
            self.fee_dic[service]["ex.GST"].grid(row=int(self.fee_dic[service]["Position"]) + 4, column=2, pady=(0,20))
            self.fee_dic[service]["ex.GST"].config(state="readonly")
            self.fee_dic[service]["in.GST"].grid(row=int(self.fee_dic[service]["Position"]) + 4, column=3, pady=(0,20))
        else:
            for i in range(3):
                self.app.data["Fee Proposal Page"]["Details"][service]["Context"][i][0].set("")
                self.app.data["Fee Proposal Page"]["Details"][service]["Context"][i][1].set("")
                self.fee_dic[service]["expand frame"]["Service"][i].grid_forget()
                self.fee_dic[service]["expand frame"]["Fee"][i].grid_forget()
                self.fee_dic[service]["expand frame"]["in.GST"][i].grid_forget()
            self.app.data["Fee Proposal Page"]["Details"][service]["Fee"].set("")
            self.fee_dic[service]["expand frame"]["Total"].grid_forget()
            self.fee_dic[service]["ex.GST"].grid(row=int(self.fee_dic[service]["Position"]), column=2,pady=0)
            self.fee_dic[service]["ex.GST"].config(state=tk.NORMAL)
            self.fee_dic[service]["in.GST"].grid(row=int(self.fee_dic[service]["Position"]), column=3,pady=0)
        self.app.invoice_page.expand(service)

    def enable_input(self, server, i):
        if len(self.app.data["Fee Proposal Page"]["Details"][server]["Context"][i][0].get())==0:
            self.fee_dic[server]["expand frame"]["Fee"][i].config(state="readonly")
        else:
            self.fee_dic[server]["expand frame"]["Fee"][i].config(state="")