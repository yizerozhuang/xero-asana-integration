import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date
import os
import json


# from scope_list import ScopeList

class FeeProposalPage(tk.Frame):
    def __init__(self, parent, app):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.app = app
        self.app.data["Fee Proposal"] = dict()
        self.data = app.data
        self.conf = app.conf

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
        self.scope_part()
        self.fee_part()

    def reference_part(self):
        reference = dict()
        self.data["Fee Proposal"]["Reference"] = reference

        reference_frame = tk.LabelFrame(self.main_context_frame, text="reference")
        reference_frame.pack(fill=tk.BOTH, expand=1, padx=20)

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

        time_frame = tk.LabelFrame(self.main_context_frame, text="Time Consuming", font=self.conf["font"])
        time_frame.pack(fill=tk.BOTH, expand=1, padx=20)

        stages = ["Fee Proposal", "Pre-design", "Documentation"]

        for i, stage in enumerate(stages):
            tk.Label(time_frame, width=30, text=stage, font=self.conf["font"]).grid(row=i, column=0)
            time[stage] = {
                "Start": tk.StringVar(value="1"),
                "End": tk.StringVar(value="2")
            }
            tk.Entry(time_frame, width=20, font=self.conf["font"], fg="blue", textvariable=time[stage]["Start"]).grid(
                row=i, column=1)
            tk.Label(time_frame, text="-").grid(row=i, column=2)
            tk.Entry(time_frame, width=20, font=self.conf["font"], fg="blue", textvariable=time[stage]["End"]).grid(
                row=i, column=3)
            tk.Label(time_frame, text=" business day", font=self.conf["font"]).grid(row=i, column=4)

    def scope_part(self):
        scope = dict()
        self.data["Fee Proposal"]["Scope"] = scope

        self.scope_frame = tk.LabelFrame(self.main_context_frame, text="Scope of Work", font=self.conf["font"])
        self.scope_frame.pack(fill=tk.BOTH, expand=1, padx=20)
        self.scope_frames = {}
        self.append_context = dict()

    def fee_part(self):
        invoices = {
            "Details": dict(),
            "Fee": tk.StringVar(),
            "in.GST": tk.StringVar()
        }
        bills = {
            "Details": dict(),
            "Fee": tk.StringVar(),
            "in.GST": tk.StringVar()
        }
        variation = [
            {
                "Service": tk.StringVar(),
                "Fee": tk.StringVar(),
                "in.GST": tk.StringVar(),
                "Number": tk.StringVar(value="None")
            } for _ in range(self.conf["n_variation"])
        ]
        ist_update_fuc = lambda i: lambda a,b,c: self.app._ist_update(variation[i]["Fee"], variation[i]["in.GST"])
        for i in range(self.conf["n_variation"]):
            variation[i]["Fee"].trace("w", ist_update_fuc(i))

        self.data["Invoices"] = invoices
        self.data["Bills"] = bills
        self.data["Variation"] = variation

        self.fee_frames = dict()
        self.fee_dic = dict()

        self.fee_frame = tk.LabelFrame(self.main_context_frame, text="Fee Proposal Details", font=self.conf["font"])
        self.fee_frame.pack(fill=tk.BOTH, expand=1, padx=20)

        top_frame = tk.LabelFrame(self.fee_frame)
        top_frame.pack()
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

        for i in range(self.conf["n_variation"]):
            variation_frame = tk.LabelFrame(self.fee_frame)
            variation_frame.pack(side=tk.BOTTOM, fill=tk.X)
            tk.Label(variation_frame, width=10, text="", font=self.conf["font"]).grid(row=0, column=0)
            tk.Entry(variation_frame, width=50, textvariable=variation[i]["Service"], font=self.conf["font"], fg="blue").grid(row=0, column=1)
            tk.Entry(variation_frame, width=20, textvariable=variation[i]["Fee"], font=self.conf["font"], fg="blue").grid(row=0, column=2, padx=(40,0))
            tk.Label(variation_frame, width=20, textvariable=variation[i]["in.GST"], font=self.conf["font"]).grid(row=0, column=3)

    def update_scope(self, var):
        scope = self.data["Fee Proposal"]["Scope"]
        service = var["Service"].get()
        include = var["Include"].get()
        scope_dir = os.path.join(self.conf["database_dir"], "scope_of_work.json")
        scope_data = json.load(open(scope_dir))
        if include:
            self.scope_frames[service] = tk.LabelFrame(self.scope_frame, text=service, font=self.conf["font"])
            self.scope_frames[service].pack()
            scope[service] = dict()
            self.append_context[service] = dict()

            extra_list = ["Extend", "Exclusion", "Deliverables"]
            for i, extra in enumerate(extra_list):
                extra_frame = tk.LabelFrame(self.scope_frames[service], text=extra, font=self.conf["font"])
                extra_frame.pack()
                # self.app.cur.execute(
                #     f"""
                #         SELECT *
                #         FROM service_scope
                #         WHERE service_type='{service}'
                #         AND extra ='{extra}'
                #     """
                # )
                # items = self.app.cur.fetchall()
                items = scope_data[service][extra]
                color_list = ["white", "azure"]
                scope[service][extra] = []
                for j, context in enumerate(items):
                    scope[service][extra].append(
                        {
                            "Include": tk.BooleanVar(value=True),
                            "Item": tk.StringVar(value=context)
                        }
                    )
                    tk.Checkbutton(extra_frame, variable=scope[service][extra][j]["Include"]).grid(row=j, column=0)
                    tk.Entry(extra_frame, width=100, textvariable=scope[service][extra][j]["Item"],
                             font=self.conf["font"], bg=color_list[j % 2]).grid(row=j, column=1)

                self.append_context[service][extra] = {
                    "Item": tk.StringVar(),
                    "Add": tk.BooleanVar()
                }
                append_frame = tk.Frame(extra_frame)
                append_frame.grid(row=999, column=0, columnspan=2)
                tk.Entry(append_frame, width=95, textvariable=self.append_context[service][extra]["Item"]).grid(row=0,
                                                                                                                column=0)
                tk.Checkbutton(append_frame, variable=self.append_context[service][extra]["Add"],
                               text="Add to Database").grid(row=0, column=1)
                func = lambda extra: lambda: self._append_value(service, extra)
                tk.Button(append_frame, text="Submit", command=func(extra)).grid(row=0, column=2)
        else:
            self.scope_frames[service].destroy()
            scope.pop(service)

    def update_fee(self, var):
        invoices_details = self.data["Invoices"]["Details"]
        service = var["Service"].get()
        include = var["Include"].get()
        if include:
            invoices_details[service] = {
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

            invoices_details[service]["Expand"].trace("w", lambda a, b, c: self._expand(service))

            invoices_details[service]["Fee"].trace("w", lambda a, b, c: self.app._ist_update(
                invoices_details[service]["Fee"], invoices_details[service]["in.GST"]))
            invoices_details[service]["Fee"].trace("w", lambda a, b, c: self.app._sum_update(
                [value["Fee"] for value in invoices_details.values()], self.data["Invoices"]["Fee"]))
            invoices_details[service]["in.GST"].trace("w", lambda a, b, c: self.app._sum_update(
                [value["in.GST"] for value in invoices_details.values()], self.data["Invoices"]["in.GST"]))

            for i in range(self.conf["n_items"]):
                func = lambda i: lambda a, b, c: self.app._ist_update(invoices_details[service]["Content"][i]["Fee"],
                                                                      invoices_details[service]["Content"][i]["in.GST"])
                invoices_details[service]["Content"][i]["Fee"].trace("w", func(i))

                sum_fun = lambda a, b, c: self.app._sum_update(
                    [item["Fee"] for item in invoices_details[service]["Content"]],
                    invoices_details[service]["Fee"])
                invoices_details[service]["Content"][i]["Fee"].trace("w", sum_fun)
        else:
            self.fee_frames[service].destroy()
            invoices_details[service]["Fee"].set("")
            invoices_details.pop(service)

    def _append_value(self, service, extra):
        scope_dir = os.path.join(self.conf["database_dir"], "scope_of_work.json")
        scope = self.app.data["Fee Proposal"]["Scope"]
        item = self.append_context[service][extra]["Item"].get()
        if item == "":
            messagebox.showwarning(title="Error", message="You cant need to enter some context")
            return
        scope[service][extra].append(
            {
                "Include": tk.BooleanVar(value=True),
                "Item": tk.StringVar(value=item)
            }
        )
        if extra == "Extend":
            index = 0
        elif extra == "Exclusion":
            index = 1
        else:
            index = 2
        tk.Entry(self.scope_frames[service].winfo_children()[index], width=100,
                 textvariable=scope[service][extra][-1]["Item"],
                 font=self.conf["font"]).grid(row=len(scope[service][extra]), column=1)

        tk.Checkbutton(self.scope_frames[service].winfo_children()[index],
                       variable=scope[service][extra][-1]["Include"]).grid(row=len(scope[service][extra]), column=0)

        if self.append_context[service][extra]["Add"].get():
            scope_data = json.load(open(scope_dir))
            scope_data[service][extra].append(item)
            with open(scope_dir, "w") as f:
                json_object = json.dumps(scope_data, indent=4)
                f.write(json_object)
            # self.app.cur.execute(f"""
            #     INSERT INTO service_scope(service_type, extra, context)
            #     VALUES
            #     ('{service}','{extra}','{item}')
            #     """)
            # self.app.conn.commit()
        self.append_context[service][extra]["Item"].set("")
        self.append_context[service][extra]["Add"].set(False)

    def reset_scrollregion(self, event):
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))

    def _expand(self, service):
        details = self.data["Invoices"]["Details"]
        if details[service]["Expand"].get():
            for i in range(self.conf["n_items"]):
                self.fee_dic[service]["Content"]["Details"][i]["Service"].grid(row=i + 1, column=1)
                self.fee_dic[service]["Content"]["Details"][i]["Fee"].grid(row=i + 1, column=2)
                self.fee_dic[service]["Content"]["Details"][i]["in.GST"].grid(row=i + 1, column=3)

            details[service]["Fee"].set("")

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
                details[service]["Content"][i]["Service"].set("")
                details[service]["Content"][i]["Fee"].set("")
                self.fee_dic[service]["Content"]["Details"][i]["Service"].grid_forget()
                self.fee_dic[service]["Content"]["Details"][i]["Fee"].grid_forget()
                self.fee_dic[service]["Content"]["Details"][i]["in.GST"].grid_forget()
            details[service]["Fee"].set("")

            self.fee_dic[service]["Content"]["Service"].grid_forget()
            self.fee_dic[service]["Content"]["Fee"].grid_forget()
            self.fee_dic[service]["Content"]["in.GST"].grid_forget()

            self.fee_dic[service]["Fee"].grid(row=0, column=2)
            self.fee_dic[service]["in.GST"].grid(row=0, column=3)

    def _on_mousewheel(self, event):
        self.main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
