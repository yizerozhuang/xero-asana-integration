import tkinter as tk
from tkinter import ttk
from datetime import date
import os
import json


from scope_list import ScopeList

class FeeProposalPage(tk.Frame):
    def __init__(self, parent, app):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.app = app
        self.app.data["Fee Proposal"] = dict()
        self.data = app.data
        self.conf = app.conf
        self.messagebox = app.messagebox

        self.main_frame = tk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=1, side=tk.LEFT)
        self.main_canvas = tk.Canvas(self.main_frame)
        self.main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1, anchor="nw")
        self.scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.main_canvas.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.main_canvas.config(yscrollcommand=self.scrollbar.set)
        self.main_canvas.bind("<Configure>",
                              lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")))
        self.app.bind("<MouseWheel>", self._on_mousewheel, add="+")
        self.main_context_frame = tk.LabelFrame(self.main_canvas, text="Main Context")
        self.main_canvas.create_window((0, 0), window=self.main_context_frame, anchor="center")
        self.main_context_frame.bind("<Configure>", self.reset_scrollregion)
        # self.pdf_frame = tk.LabelFrame(self.main_frame, text="PDF preview")
        # self.main_canvas.create_window((1500, 0), window=self.pdf_frame, anchor="e")
        # self.pdf_frame.bind("<Configure>", self.reset_scrollregion)

        self.reference_part()
        self.time_part()
        self.stage_part()
        self.scope_part()
        self.fee_part()

    def reference_part(self):
        reference = dict()
        self.data["Fee Proposal"]["Reference"] = reference

        reference_frame = tk.LabelFrame(self.main_context_frame, text="reference")
        reference_frame.grid(row=0, column=0, sticky="ew", padx=20)

        reference["Date"] = tk.StringVar(value=date.today().strftime("%d-%b-%Y"))
        tk.Label(reference_frame, width=30, text="Date", font=self.conf["font"]).grid(row=0, column=0, padx=(10, 0))
        tk.Entry(reference_frame, width=70, textvariable=reference["Date"], fg="blue", font=self.conf["font"]).grid(
            row=0, column=1, padx=(0, 10))
        reference["Revision"] = tk.StringVar(value="1")
        tk.Label(reference_frame, width=30, text="Revision", font=self.conf["font"]).grid(row=1, column=0, padx=(10, 0))
        tk.Entry(reference_frame, width=70, textvariable=reference["Revision"], fg="blue", font=self.conf["font"]).grid(
            row=1, column=1, padx=(0, 10))

    def time_part(self):
        time = dict()
        self.data["Fee Proposal"]["Time"] = time

        self.time_frame = tk.LabelFrame(self.main_context_frame, text="Time Consuming", font=self.conf["font"])
        self.time_frame.grid(row=1, column=0, sticky="ew", padx=20)

        stages = ["Fee Proposal", "Pre-design", "Documentation"]

        for i, stage in enumerate(stages):
            tk.Label(self.time_frame, width=30, text=stage, font=self.conf["font"]).grid(row=i, column=0)
            time[stage] = {
                "Start": tk.StringVar(value="1"),
                "End": tk.StringVar(value="2")
            }
            tk.Entry(self.time_frame, width=20, font=self.conf["font"], fg="blue", textvariable=time[stage]["Start"]).grid(
                row=i, column=1)
            tk.Label(self.time_frame, text="-").grid(row=i, column=2)
            tk.Entry(self.time_frame, width=20, font=self.conf["font"], fg="blue", textvariable=time[stage]["End"]).grid(
                row=i, column=3)
            tk.Label(self.time_frame, text=" business day", font=self.conf["font"]).grid(row=i, column=4)

    def stage_part(self):
        stage_dict = dict()
        self.data["Fee Proposal"]["Stage"] = stage_dict
        self.data["Project Info"]["Project"]["Proposal Type"].trace("w", self._update_stage)
        self.stage_frame = tk.LabelFrame(self.main_context_frame, text="Stage", font=self.conf["font"])

        for i, stage in enumerate(self.conf["major_stage"]):
            stage_dict[stage] = tk.BooleanVar(value=True)
            tk.Checkbutton(self.stage_frame, variable=stage_dict[stage], text=stage, font=self.conf["font"]).grid(row=0, column=i)
    def _update_stage(self, *args):
        if self.data["Project Info"]["Project"]["Proposal Type"].get() == "Minor":
            self.stage_frame.grid_forget()
            self.time_frame.grid(row=1, column=0, sticky="ew", padx=20)
        else:
            self.time_frame.grid_forget()
            self.stage_frame.grid(row=1, column=0, sticky="ew", padx=20)
    def _reset_scope(self):
        service_list = self.conf["service_list"]
        extra_list = self.conf["extra_list"]
        scope = dict()
        scope_dir = os.path.join(self.conf["database_dir"], "scope_of_work.json")
        scope_data = json.load(open(scope_dir))

        for type in ["Minor", "Major"]:
            scope[type] = dict()
            for service in service_list:
                scope[type][service] = dict()
                for extra in extra_list:
                    # scope[service][extra] = ScopeList(self.app, service, extra)
                    scope[type][service][extra] = []
                    items = scope_data[type][service][extra]
                    for context in items:
                        scope[type][service][extra].append(
                            {
                                "Include": tk.BooleanVar(value=True),
                                "Item": tk.StringVar(value=context)
                            }
                        )

        self.data["Fee Proposal"]["Scope"] = scope
    def scope_part(self):
        self._reset_scope()
        self.scope_frame = tk.LabelFrame(self.main_context_frame, text="Scope of Work", font=self.conf["font"])
        self.scope_frame.grid(row=2, column=0, sticky="ew", padx=20)
        self.scope_frames = {}
        self.append_context = dict()
    def update_scope(self, var):
        scope = self.data["Fee Proposal"]["Scope"][self.data["Project Info"]["Project"]["Proposal Type"].get()]
        service = var["Service"].get()
        include = var["Include"].get()
        extra_list = self.conf["extra_list"]
        if include:
            self.scope_frames[service] = dict()
            self.scope_frames[service]["Main Frame"] = tk.LabelFrame(self.scope_frame, text=service, font=self.conf["font"])
            self.scope_frames[service]["Main Frame"].pack()
            self.append_context[service] = dict()
            for i, extra in enumerate(extra_list):
                extra_frame = tk.LabelFrame(self.scope_frames[service]["Main Frame"], text=extra, font=self.conf["font"])
                extra_frame.pack()
                self.scope_frames[service][extra] = tk.Frame(extra_frame)
                self.scope_frames[service][extra].pack()
                color_list = ["white", "azure"]
                for j, context in enumerate(scope[service][extra]):
                    tk.Checkbutton(self.scope_frames[service][extra], variable=scope[service][extra][j]["Include"]).grid(row=j, column=0)
                    tk.Entry(self.scope_frames[service][extra], width=100, textvariable=scope[service][extra][j]["Item"],
                             font=self.conf["font"], bg=color_list[j % 2]).grid(row=j, column=1)

                self.append_context[service][extra] = {
                    "Item": tk.StringVar(),
                    "Add": tk.BooleanVar()
                }
                append_frame = tk.Frame(extra_frame)
                append_frame.pack()
                tk.Entry(append_frame, width=95, textvariable=self.append_context[service][extra]["Item"]).grid(row=0, column=0)
                tk.Checkbutton(append_frame, variable=self.append_context[service][extra]["Add"], text="Add to Database").grid(row=0, column=1)
                func = lambda extra: lambda: self._append_value(service, extra)
                tk.Button(append_frame, text="Submit", command=func(extra)).grid(row=0, column=2)
        else:
            self.scope_frames[service]["Main Frame"].destroy()
    def fee_part(self):
        invoices = {
            "Details": dict(),
            "Fee": tk.StringVar(),
            "in.GST": tk.StringVar()
        }
        for service in self.conf["invoice_list"]:
            invoices["Details"][service] = {
                "Include": tk.BooleanVar(),
                "Service": tk.StringVar(value=service),
                "Fee": tk.StringVar(),
                "in.GST": tk.StringVar(),
                "Expand": tk.BooleanVar(),
                "Number": tk.StringVar(value="None"),
                "Content": [
                    {
                        "Service": tk.StringVar(),
                        "Fee": tk.StringVar(),
                        "in.GST": tk.StringVar(),
                        "Number": tk.StringVar(value="None")
                    } for _ in range(self.conf["n_items"])
                ]
            }
            expand_fun = lambda service : lambda a, b, c: self._expand(service)
            invoices["Details"][service]["Expand"].trace("w", expand_fun(service))
            invoices["Details"][service]["Expand"].trace("w", self.update_sum)


            ist_update_fun = lambda service : lambda a, b, c: self.app._ist_update(invoices["Details"][service]["Fee"], invoices["Details"][service]["in.GST"])
            invoices["Details"][service]["Fee"].trace("w", ist_update_fun(service))
            # invoices["Details"][service]["Fee"].trace("w", self.update_sum)
            invoices["Details"][service]["in.GST"].trace("w", self.update_sum)

            func = lambda service, i: lambda a, b, c: self.app._ist_update(
                invoices["Details"][service]["Content"][i]["Fee"],
                invoices["Details"][service]["Content"][i]["in.GST"])
            sum_fun = lambda service: lambda a, b, c: self.app._sum_update(
                [item["Fee"] for item in invoices["Details"][service]["Content"]], invoices["Details"][service]["Fee"])
            for i in range(self.conf["n_items"]):

                invoices["Details"][service]["Content"][i]["Fee"].trace("w", func(service, i))

                invoices["Details"][service]["Content"][i]["Fee"].trace("w", sum_fun(service))
            invoices["Details"][service]["Expand"].trace("w", sum_fun(service))

        self.data["Invoices"] = invoices

        self.fee_frames = dict()
        self.fee_dic = dict()

        self.fee_frame = tk.LabelFrame(self.main_context_frame, text="Fee Proposal Details", font=self.conf["font"])
        self.fee_frame.grid(row=3, column=0, sticky="ew", padx=20)

        top_frame = tk.LabelFrame(self.fee_frame)
        top_frame.pack(side=tk.TOP)
        tk.Label(top_frame, text="", width=6, font=self.conf["font"]).grid(row=0, column=0)
        tk.Label(top_frame, width=50, text="Services", font=self.conf["font"]).grid(row=0, column=1)
        tk.Label(top_frame, width=20, text="ex.GST", font=self.conf["font"]).grid(row=0, column=2)
        tk.Label(top_frame, width=20, text="in.GST", font=self.conf["font"]).grid(row=0, column=3)

        bottom_frame = tk.LabelFrame(self.fee_frame)
        bottom_frame.pack(side=tk.BOTTOM)
        tk.Label(bottom_frame, width=6, text="", font=self.conf["font"]).grid(row=0, column=0)
        tk.Label(bottom_frame, width=50, text="Total", font=self.conf["font"]).grid(row=0, column=1)

        tk.Label(bottom_frame, width=20, textvariable=invoices["Fee"], font=self.conf["font"]).grid(row=0, column=2)
        tk.Label(bottom_frame, width=20, textvariable=invoices["in.GST"], font=self.conf["font"]).grid(row=0, column=3)
    def update_fee(self, var):
        invoices_details = self.data["Invoices"]["Details"]
        # archive = self.data["Archive"]["Invoice"]
        service = var["Service"].get()
        include = var["Include"].get()

        if include:
            invoices_details[service]["Include"].set(True)
            if not service in self.fee_dic:
                self.fee_frames[service] = tk.LabelFrame(self.fee_frame)
                self.fee_frames[service].pack()
                self.fee_dic[service] = {
                    "Service": tk.Label(self.fee_frames[service],
                                        text=service + " design and documentation",
                                        width=50,
                                        font=self.conf["font"]),
                    "Fee": tk.Entry(self.fee_frames[service],
                                    textvariable=invoices_details[service]["Fee"],
                                    width=20,
                                    font=self.conf["font"],
                                    fg="blue"),
                    "in.GST": tk.Label(self.fee_frames[service],
                                       textvariable=invoices_details[service]["in.GST"],
                                       width=20,
                                       font=self.conf["font"]),
                    "Expand": tk.Checkbutton(self.fee_frames[service],
                                             text="Expand",
                                             variable=invoices_details[service]["Expand"],
                                             font=self.conf["font"]),
                    "Content": {
                        "Details": [
                            {
                                "Service": tk.Entry(self.fee_frames[service],
                                                    width=36,
                                                    font=self.conf["font"],
                                                    fg="blue",
                                                    textvariable=invoices_details[service]["Content"][i]["Service"]),
                                "Fee": tk.Entry(self.fee_frames[service],
                                                width=20,
                                                font=self.conf["font"],
                                                fg="blue",
                                                textvariable=invoices_details[service]["Content"][i]["Fee"]),
                                "in.GST": tk.Label(self.fee_frames[service],
                                                   width=17,
                                                   textvariable=invoices_details[service]["Content"][i]["in.GST"],
                                                   font=self.conf["font"])
                            } for i in range(self.conf["n_items"])
                        ],
                        "Service": tk.Label(self.fee_frames[service], width=36, text=service + " Total",
                                            font=self.conf["font"]),
                        "Fee": tk.Label(self.fee_frames[service],
                                        width=20,
                                        textvariable=invoices_details[service]["Fee"],
                                        font=self.conf["font"]),
                        "in.GST": tk.Label(self.fee_frames[service],
                                           width=17,
                                           textvariable=invoices_details[service]["in.GST"],
                                           font=self.conf["font"])
                    }
                }

                self.fee_dic[service]["Expand"].grid(row=0, column=0)
                self.fee_dic[service]["Service"].grid(row=0, column=1)
                self.fee_dic[service]["Fee"].grid(row=0, column=2)
                self.fee_dic[service]["in.GST"].grid(row=0, column=3)

            self.fee_frames[service].pack()
            self.update_sum()
        else:
            # archive[service] = invoices_details.pop(service)
            invoices_details[service]["Include"].set(False)
            # invoices_details[service]["Expand"].set(False)
            self.fee_frames[service].pack_forget()
            self.update_sum()
    def _append_value(self, service, extra):
        scope_dir = os.path.join(self.conf["database_dir"], "scope_of_work.json")
        scope_type = self.data["Project Info"]["Project"]["Proposal Type"].get()
        scope = self.app.data["Fee Proposal"]["Scope"][scope_type]
        item = self.append_context[service][extra]["Item"].get()
        if len(item.strip()) == 0:
            self.messagebox.show_error(title="Error", message="You need to enter some context")
            return
        scope[service][extra].append(
            {
                "Include": tk.BooleanVar(value=True),
                "Item": tk.StringVar(value=item)
            }
        )
        tk.Entry(self.scope_frames[service][extra], width=100,
                 textvariable=scope[service][extra][-1]["Item"],
                 font=self.conf["font"]).grid(row=len(scope[service][extra]), column=1)

        tk.Checkbutton(self.scope_frames[service][extra],
                       variable=scope[service][extra][-1]["Include"]).grid(row=len(scope[service][extra]), column=0)

        if self.append_context[service][extra]["Add"].get():
            scope_data = json.load(open(scope_dir))
            scope_data[scope_type][service][extra].append(item)
            with open(scope_dir, "w") as f:
                json_object = json.dumps(scope_data, indent=4)
                f.write(json_object)
        self.append_context[service][extra]["Item"].set("")
        self.append_context[service][extra]["Add"].set(False)

    def reset_scrollregion(self, event):
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))

    def _expand(self, service):
        details = self.data["Invoices"]["Details"]
        if not service in self.fee_dic.keys():
            return

        if details[service]["Expand"].get():
            for i in range(self.conf["n_items"]):
                self.fee_dic[service]["Content"]["Details"][i]["Service"].grid(row=i + 1, column=1)
                self.fee_dic[service]["Content"]["Details"][i]["Fee"].grid(row=i + 1, column=2)
                self.fee_dic[service]["Content"]["Details"][i]["in.GST"].grid(row=i + 1, column=3)


            self.fee_dic[service]["Content"]["Service"].grid(row=self.conf["n_items"] + 1,
                                                             column=1)
            self.fee_dic[service]["Content"]["Fee"].grid(row=self.conf["n_items"] + 1,
                                                         column=2)
            self.fee_dic[service]["Content"]["in.GST"].grid(row=self.conf["n_items"] + 1,
                                                            column=3)
            self.fee_dic[service]["Fee"].grid_forget()
            self.fee_dic[service]["in.GST"].grid_forget()
        else:
            for i in range(self.conf["n_items"]):
                self.fee_dic[service]["Content"]["Details"][i]["Service"].grid_forget()
                self.fee_dic[service]["Content"]["Details"][i]["Fee"].grid_forget()
                self.fee_dic[service]["Content"]["Details"][i]["in.GST"].grid_forget()

            self.fee_dic[service]["Content"]["Service"].grid_forget()
            self.fee_dic[service]["Content"]["Fee"].grid_forget()
            self.fee_dic[service]["Content"]["in.GST"].grid_forget()

            self.fee_dic[service]["Fee"].grid(row=0, column=2)
            self.fee_dic[service]["in.GST"].grid(row=0, column=3)

    def _on_mousewheel(self, event):
        self.main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def update_sum(self, *args):
        details = self.data["Invoices"]["Details"]
        total = self.data["Invoices"]["Fee"]
        total_ist = self.data["Invoices"]["in.GST"]
        # variation = self.data["Variation"]
        sum = 0
        ist_sum = 0
        for fee in details.values():
            if len(fee["Fee"].get().strip()) != 0 and fee["Include"].get():
                try:
                    sum += float(fee["Fee"].get())
                    # ist_sum += float(fee["Fee"].get()) * self.conf["tax rates"]
                    ist_sum += float(fee["in.GST"].get())
                except ValueError as e:
                    total.set("Error")
                    total_ist.set("Error")
                    print(e)
                    return
        total.set(str(round(sum, 2)))
        total_ist.set(str(round(ist_sum, 2)))
