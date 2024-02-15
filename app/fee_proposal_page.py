import tkinter as tk
from tkinter import ttk
from datetime import date
import os
import json
from text_extension import TextExtension

from utility import preview_installation_fee_proposal, email_installation_proposal, isfloat

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


        self.email_part()
        self.calculation_part()

        self.function_part()
        self.reference_part()
        self.time_part()
        self.stage_part()
        self.scope_part()
        self.fee_part()

    def email_part(self):
        pass

    def calculation_part(self):
        max_row = 12

        calculate_frame = tk.LabelFrame(self.main_context_frame)
        calculate_frame.grid(row=0, column=1, rowspan=6, sticky="n")
        calculate = {
            "Car Park":[
                {
                    "Project": tk.StringVar(),
                    "Car park Level": tk.StringVar(),
                    "Number of Carports": tk.StringVar(),
                    "Level Factor": tk.StringVar(value="0"),
                    "Carport Factor": tk.StringVar(value="0"),
                    "Complex Factor": tk.StringVar(value="0"),
                    "CFD Cost": tk.StringVar(value="0")
                } for _ in range(self.conf["car_park_row"])
            ],
            "Apt": tk.StringVar(),
            "Custom Apt": tk.StringVar(),
            "Area": tk.StringVar(),
            "Custom Area": tk.StringVar()
        }
        self.data["Fee Proposal"]["Calculation Part"] = calculate

        self.data["Project Info"]["Building Features"]["Minor"]["Total Area"].trace("w", lambda a, b, c: self.set_variable(calculate["Area"], self.data["Project Info"]["Building Features"]["Minor"]["Total Area"]))
        self.data["Project Info"]["Building Features"]["Major"]["Total Apt"].trace("w", lambda a, b, c: self.set_variable(calculate["Apt"], self.data["Project Info"]["Building Features"]["Major"]["Total Apt"]))

        area_frame = tk.Frame(calculate_frame)
        area_frame.pack()

        tk.Label(area_frame, text="Total Areas: ").grid(row=0, column=0)
        tk.Entry(area_frame, textvariable=calculate["Area"], font=self.conf["font"], fg="blue").grid(row=0, column=1)
        tk.Label(area_frame, text="Custom Area: ").grid(row=0, column=2)
        tk.Entry(area_frame, textvariable=calculate["Custom Area"], font=self.conf["font"], fg="blue").grid(row=0,
                                                                                                            column=3)
        tk.Label(area_frame, text="Price: ").grid(row=0, column=4)
        self.cus_area_entry = tk.Entry(area_frame, font=self.conf["font"])
        self.cus_area_entry.grid(row=0, column=5)
        calculate["Area"].trace("w", lambda a, b, c: self.calculate_apt_price(calculate["Custom Area"].get(),
                                                                              self.cus_area_entry, "Area"))
        calculate["Custom Area"].trace("w", lambda a, b, c: self.calculate_apt_price(calculate["Custom Area"].get(),
                                                                                     self.cus_area_entry, "Area"))
        calculation_function = lambda i, entry, type: lambda a,b,c: self.calculate_apt_price(i, entry, type)
        self.area_entry = []
        num = 0
        for i in list(range(100, 300, 50)) + [300, 325] + list(range(375, 1500, 25)) + list(range(1500, 1950, 50)):
            # i = float("{0:.2f}".format(float(i)/100))
            i = float(i) / 100
            tk.Label(area_frame, text=f"${i}").grid(row=num % max_row + 1, column=(num // max_row) * 2)
            self.area_entry.append(tk.Entry(area_frame, font=self.conf["font"]))
            self.area_entry[num].grid(row=num % max_row + 1, column=(num // max_row) * 2 + 1)
            calculate["Area"].trace("w", calculation_function(i, self.area_entry[num], "Area"))
            num += 1

        apt_frame = tk.Frame(calculate_frame)
        apt_frame.pack()
        tk.Label(apt_frame, text="Total Apts: ").grid(row=0, column=0)
        tk.Entry(apt_frame, textvariable=calculate["Apt"], font=self.conf["font"], fg="blue").grid(row=0, column=1)
        tk.Label(apt_frame, text="Custom Apt: ").grid(row=0, column=2)
        tk.Entry(apt_frame, textvariable=calculate["Custom Apt"], font=self.conf["font"], fg="blue").grid(row=0, column=3)
        tk.Label(apt_frame, text="Price: ").grid(row=0, column=4)
        self.cus_apt_entry = tk.Entry(apt_frame, font=self.conf["font"])
        self.cus_apt_entry.grid(row=0, column=5)
        calculate["Apt"].trace("w", lambda a, b, c: self.calculate_apt_price(calculate["Custom Apt"].get(), self.cus_apt_entry, "Apt"))
        calculate["Custom Apt"].trace("w", lambda a, b, c: self.calculate_apt_price(calculate["Custom Apt"].get(), self.cus_apt_entry, "Apt"))
        self.apt_entry = []
        num=0
        for i in list(range(80, 130, 5)) + list(range(130, 630, 10)):
            tk.Label(apt_frame, text=f"${i}").grid(row=num%max_row+1, column=(num//max_row)*2)
            self.apt_entry.append(tk.Entry(apt_frame, font=self.conf["font"]))
            self.apt_entry[num].grid(row=num%max_row+1, column=(num//max_row)*2+1)
            calculate["Apt"].trace("w", calculation_function(i, self.apt_entry[num], "Apt"))
            num+=1


        car_park_frame = tk.Frame(calculate_frame)
        car_park_frame.pack()
        tk.Label(car_park_frame, text="Project").grid(row=0, column=0)
        tk.Label(car_park_frame, text="CFD Calc Level").grid(row=0, column=1)
        tk.Label(car_park_frame, text="No of Carports").grid(row=0, column=2)
        tk.Label(car_park_frame, text="Level Factor").grid(row=0, column=3)
        tk.Label(car_park_frame, text="Carport Factor").grid(row=0, column=4)
        tk.Label(car_park_frame, text="Complex Factor").grid(row=0, column=5)
        tk.Label(car_park_frame, text="CFD Cost").grid(row=0, column=6)

        level_factor_function = lambda i: lambda a,b,c: self.level_factor_calculation(calculate["Car Park"][i]["Car park Level"], calculate["Car Park"][i]["Level Factor"])
        carport_factor_function = lambda i: lambda a,b,c: self.carport_factor_calculation(calculate["Car Park"][i]["Number of Carports"], calculate["Car Park"][i]["Carport Factor"])
        complex_factor_function = lambda i:lambda a,b,c: self.complex_factor_calculation(calculate["Car Park"][i]["Level Factor"], calculate["Car Park"][i]["Carport Factor"], calculate["Car Park"][i]["Complex Factor"])
        cfd_cost_function = lambda i:lambda a,b,c: self.cfd_cost_calculation(calculate["Car Park"][i]["Complex Factor"], calculate["Car Park"][i]["CFD Cost"])
        for i in range(self.conf["car_park_row"]):
            calculate["Car Park"][i]["Car park Level"].trace("w", level_factor_function(i))
            calculate["Car Park"][i]["Number of Carports"].trace("w", carport_factor_function(i))
            calculate["Car Park"][i]["Level Factor"].trace("w", complex_factor_function(i))
            calculate["Car Park"][i]["Carport Factor"].trace("w", complex_factor_function(i))
            calculate["Car Park"][i]["Complex Factor"].trace("w", cfd_cost_function(i))
            tk.Entry(car_park_frame, textvariable=calculate["Car Park"][i]["Project"], font=self.conf["font"], fg="blue").grid(row=1+i, column=0)
            tk.Entry(car_park_frame, textvariable=calculate["Car Park"][i]["Car park Level"], font=self.conf["font"], fg="blue").grid(row=1+i, column=1)
            tk.Entry(car_park_frame, textvariable=calculate["Car Park"][i]["Number of Carports"], font=self.conf["font"], fg="blue").grid(row=1+i, column=2)
            tk.Label(car_park_frame, textvariable=calculate["Car Park"][i]["Level Factor"], font=self.conf["font"]).grid(row=1 + i, column=3)
            tk.Label(car_park_frame, textvariable=calculate["Car Park"][i]["Carport Factor"], font=self.conf["font"]).grid(row=1 + i, column=4)
            tk.Label(car_park_frame, textvariable=calculate["Car Park"][i]["Complex Factor"], font=self.conf["font"]).grid(row=1 + i, column=5)
            tk.Label(car_park_frame, textvariable=calculate["Car Park"][i]["CFD Cost"], font=self.conf["font"]).grid(row=1 + i, column=6)

    def function_part(self):
        reference = {
            "Date": tk.StringVar(value=date.today().strftime("%d-%b-%Y")),
            "Revision": tk.StringVar(value="1"),
            "Program": tk.StringVar()
        }
        self.data["Fee Proposal"]["Installation Reference"] = reference
        self.installation_frame = tk.LabelFrame(self.main_context_frame, text="Installation Function")
        function_frame = tk.Frame(self.installation_frame)

        self.installation_dic = {
            "Date": tk.Entry(self.installation_frame, width=44, textvariable=reference["Date"], fg="blue", font=self.conf["font"]),
            "Today": tk.Button(self.installation_frame, width=20, text="Today", font=self.conf["font"], bg="Brown",
                               fg="white", command=self.installation_today),
            "Revision": tk.Entry(self.installation_frame, width=70, textvariable=reference["Revision"], fg="blue", font=self.conf["font"]),
            "Program": TextExtension(self.installation_frame, textvariable=reference["Program"], font=self.conf["font"], height=4, fg="blue")
        }

        function_frame.grid(row=0, column=0, columnspan=2)
        self.preview_installation_proposal_button = tk.Button(function_frame, text="Preview Installation Proposal", command=lambda: preview_installation_fee_proposal(self.app),
                                                              bg="Brown", fg="white", font=self.conf["font"])
        self.preview_installation_proposal_button.grid(row=0, column=1)
        self.email_installation_proposal_buttion = tk.Button(function_frame, text="Email Installation Proposal", command=lambda: email_installation_proposal(self.app),
                                                             bg="Brown", fg="white", font=self.conf["font"])
        self.email_installation_proposal_buttion.grid(row=0, column=2)
        reference["Date"] = tk.StringVar(value=date.today().strftime("%d-%b-%Y"))

        tk.Label(self.installation_frame, width=30, text="Date", font=self.conf["font"]).grid(row=1, column=0, padx=(10, 0))
        self.installation_dic["Date"].grid(row=1, column=1)
        self.installation_dic["Today"].grid(row=1, column=2)
        reference["Revision"] = tk.StringVar(value="1")
        tk.Label(self.installation_frame, width=30, text="Revision", font=self.conf["font"]).grid(row=2, column=0, padx=(10, 0))
        self.installation_dic["Revision"].grid(row=2, column=1, padx=(0, 10), columnspan=2)
        tk.Label(self.installation_frame, text="Program", font=self.conf["font"]).grid(row=3, column=0)
        self.installation_dic["Program"].grid(row=3, column=1, padx=(0, 10), columnspan=2)
        reference["Program"].set(
            """
                Week 1: site induction, site inspection and site measure, coordination.
Week 2-3: Order ductwork, VCDs, material and arrange labour.
Week 4-5: Installation commissioning and testing to critical area.
Week 6-8: Based on site condition, finalize all installation, provide installation certificate.
            """
        )

        self.data["Project Info"]["Project"]["Service Type"]["Installation"]["Include"].trace("w", self._update_frame)


    def reference_part(self):
        reference = {
            "Date": tk.StringVar(value=date.today().strftime("%d-%b-%Y")),
            "Revision": tk.StringVar(value="1")
        }
        self.data["Fee Proposal"]["Reference"] = reference

        reference_frame = tk.LabelFrame(self.main_context_frame, text="Reference")
        reference_frame.grid(row=1, column=0, sticky="ew", padx=20)

        self.reference_dic = {
            "Date":tk.Entry(reference_frame, width=44, textvariable=reference["Date"], fg="blue", font=self.conf["font"]),
            "Today":tk.Button(reference_frame, width=20, text="Today", font=self.conf["font"], bg="Brown", fg="white", command=self.today),
            "Revision":tk.Entry(reference_frame, width=70, textvariable=reference["Revision"], fg="blue", font=self.conf["font"])
        }


        tk.Label(reference_frame, width=30, text="Date", font=self.conf["font"]).grid(row=0, column=0, padx=(10, 0))
        self.reference_dic["Date"].grid(row=0, column=1)
        self.reference_dic["Today"].grid(row=0, column=2)
        tk.Label(reference_frame, width=30, text="Revision", font=self.conf["font"]).grid(row=1, column=0, padx=(10, 0))
        self.reference_dic["Revision"].grid(row=1, column=1, padx=(0, 10), columnspan=2)

    def time_part(self):
        time = dict()
        self.data["Fee Proposal"]["Time"] = time

        self.time_dic = {}

        self.time_frame = tk.LabelFrame(self.main_context_frame, text="Time Frame", font=self.conf["font"])
        self.time_frame.grid(row=2, column=0, sticky="ew", padx=20)

        stages = ["Fee Proposal", "Pre-design", "Documentation"]

        for i, stage in enumerate(stages):
            tk.Label(self.time_frame, width=30, text=stage, font=self.conf["font"]).grid(row=i, column=0)
            time[stage] = tk.StringVar(value="1-2")
            self.time_dic[stage] = tk.Entry(self.time_frame, width=20, font=self.conf["font"], fg="blue", textvariable=time[stage])
            self.time_dic[stage].grid(row=i, column=1)
            tk.Label(self.time_frame, text=" business days", font=self.conf["font"]).grid(row=i, column=4)

    def stage_part(self):
        stage_dict = dict()
        self.data["Fee Proposal"]["Stage"] = stage_dict
        self.data["Project Info"]["Project"]["Proposal Type"].trace("w", self._update_stage)
        self.stage_frame = tk.LabelFrame(self.main_context_frame, text="Stage", font=self.conf["font"])
        self.stage_frames = dict()
        self.append_stage = dict()
        self.stage_dic = dict()

        stage_dir = os.path.join(self.conf["database_dir"], "general_scope_of_staging.json")
        stage_data = json.load(open(stage_dir))
        for i, stage in enumerate(self.conf["major_stage"]):
            stage_dict[f"Stage{i+1}"]={
                "Service": tk.StringVar(value=stage),
                "Include": tk.BooleanVar(value=True),
                "Items": []
            }

            include_frame = tk.Frame(self.stage_frame)
            include_frame.pack(anchor="w")

            self.stage_dic[f"Stage{i+1}"]={
                "Service": tk.Entry(include_frame, textvariable=stage_dict[f"Stage{i+1}"]["Service"], font=self.conf["font"], fg="blue", width=30),
                "Include": tk.Checkbutton(include_frame, variable=stage_dict[f"Stage{i+1}"]["Include"], font=self.conf["font"]),
                "Items": []
            }
            self.stage_dic[f"Stage{i + 1}"]["Service"].grid(row=0, column=1)
            self.stage_dic[f"Stage{i + 1}"]["Include"].grid(row=0, column=0)

            extra_frame = tk.LabelFrame(self.stage_frame, font=self.conf["font"])
            extra_frame.pack()
            self.stage_frames[f"Stage{i+1}"] = tk.Frame(extra_frame)
            self.stage_frames[f"Stage{i+1}"].pack()
            color_list = ["white", "azure"]


            for j, item in enumerate(stage_data[f"Stage{i+1}"]):
                content = {
                    "Include": tk.BooleanVar(value=True),
                    "Item": tk.StringVar(value=item)
                }
                stage_dict[f"Stage{i+1}"]["Items"].append(content)

                self.stage_dic[f"Stage{i+1}"]["Items"].append(
                    {
                        "Include": tk.Checkbutton(self.stage_frames[f"Stage{i+1}"], variable=content["Include"]),
                        "Item": tk.Entry(self.stage_frames[f"Stage{i+1}"], width=100, textvariable=content["Item"], font=self.conf["font"], bg=color_list[j % 2])
                    }
                )
                self.stage_dic[f"Stage{i+1}"]["Items"][j]["Include"].grid(row=j, column=0)
                self.stage_dic[f"Stage{i+1}"]["Items"][j]["Item"].grid(row=j, column=1)

            self.append_stage[f"Stage{i+1}"] = {
                "Item": tk.StringVar(),
                "Add": tk.BooleanVar()
            }
            append_frame = tk.Frame(extra_frame)
            append_frame.pack()

            tk.Entry(append_frame, width=95, textvariable=self.append_stage[f"Stage{i+1}"]["Item"]).grid(row=0, column=0)
            tk.Checkbutton(append_frame, variable=self.append_stage[f"Stage{i+1}"]["Add"], text="Add to Database").grid(row=0, column=1)
            func = lambda i: lambda: self._append_stage(i)
            tk.Button(append_frame, text="Submit", command=func(i)).grid(row=0, column=2)

    # def _update_stage(self):

    def _update_stage(self, *args):
        if self.data["Project Info"]["Project"]["Proposal Type"].get() == "Minor":
            self.stage_frame.grid_forget()
            self.time_frame.grid(row=2, column=0, sticky="ew", padx=20)
        else:
            self.time_frame.grid_forget()
            self.stage_frame.grid(row=2, column=0, sticky="ew", padx=20)
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
        self.scope_frame.grid(row=3, column=0, sticky="ew", padx=20)
        self.scope_frames = {}
        self.scope_dic = {}
        self.append_context = dict()

    def update_scope(self, var):
        scope = self.data["Fee Proposal"]["Scope"][self.data["Project Info"]["Project"]["Proposal Type"].get()]
        service = var["Service"].get()
        include = var["Include"].get()
        extra_list = self.conf["extra_list"]
        if include:
            self.scope_frames[service] = dict()
            self.scope_dic[service] = dict()
            self.scope_frames[service]["Main Frame"] = tk.LabelFrame(self.scope_frame, text=service, font=self.conf["font"])
            self.scope_frames[service]["Main Frame"].pack()
            self.append_context[service] = dict()
            for i, extra in enumerate(extra_list):
                # if len(scope[service][extra]) == 0:
                #     continue
                extra_frame = tk.LabelFrame(self.scope_frames[service]["Main Frame"], text=extra, font=self.conf["font"])
                extra_frame.pack()
                self.scope_frames[service][extra] = tk.Frame(extra_frame)
                self.scope_dic[service][extra] = []
                self.scope_frames[service][extra].pack()
                color_list = ["white", "azure"]
                for j, context in enumerate(scope[service][extra]):
                    self.scope_dic[service][extra].append(
                        {
                            "Checkbutton": tk.Checkbutton(self.scope_frames[service][extra], variable=scope[service][extra][j]["Include"])
                        }
                    )
                    self.scope_dic[service][extra][j]["Checkbutton"].grid(row=j, column=0)
                    self.scope_dic[service][extra][j]["Entry"] = tk.Entry(self.scope_frames[service][extra], width=100, textvariable=scope[service][extra][j]["Item"], font=self.conf["font"], bg=color_list[j%2])
                    self.scope_dic[service][extra][j]["Entry"].grid(row=j, column=1)

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

        self.fee_frame = tk.LabelFrame(self.main_context_frame, text="Fee Proposal Details", font=self.conf["font"])
        self.fee_frame.grid(row=4, column=0, sticky="ew", padx=20)

        invoices = {
            "Details": dict(),
            "Fee": tk.StringVar(),
            "in.GST": tk.StringVar(),
            "Paid Fee":tk.StringVar()
        }
        for service in self.conf["invoice_list"]:
            invoices["Details"][service] = {
                "Include": tk.BooleanVar(),
                "Service": tk.StringVar(value=service),
                "Expand": tk.BooleanVar(),
                "Number": tk.StringVar(value="None"),
                "Content": [
                    {
                        "Service": tk.StringVar(),
                        "Fee": tk.StringVar(),
                        "in.GST": tk.StringVar(),
                        "Number": tk.StringVar(value="None")
                    } for _ in range(self.conf["n_items"])
                ],
                "Fee": tk.StringVar(),
                "in.GST": tk.StringVar()
            }
            if service != "Variation":
                invoices["Details"][service]["Content"][0]["Service"].set(service+" Kickoff")
                invoices["Details"][service]["Content"][1]["Service"].set(service+" Final Documentation")
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

        top_frame = tk.LabelFrame(self.fee_frame)
        top_frame.pack(side=tk.TOP)
        # tk.Label(top_frame, text="", width=6, font=self.conf["font"]).grid(row=0, column=0)
        self.proposal_lock=tk.Button(top_frame, width=20, text="Unlock", bg="Brown", fg="white", command=self.unlock, font=self.conf["font"])
        self.proposal_lock.grid(row=0, column=0)
        self.data["Lock"]["Proposal"].trace("w", self.config_lock_button)
        self.data["Lock"]["Proposal"].trace("w", self._config_entry)
        self.data["Lock"]["Proposal"].set(False)

        tk.Label(top_frame, width=35, text="Services", font=self.conf["font"]).grid(row=0, column=1)
        tk.Label(top_frame, width=20, text="ex.GST", font=self.conf["font"]).grid(row=0, column=2)
        tk.Label(top_frame, width=20, text="in.GST", font=self.conf["font"]).grid(row=0, column=3)

        bottom_frame = tk.LabelFrame(self.fee_frame)
        bottom_frame.pack(side=tk.BOTTOM)
        tk.Label(bottom_frame, width=6, text="", font=self.conf["font"]).grid(row=0, column=0)
        tk.Label(bottom_frame, width=50, text="Total", font=self.conf["font"]).grid(row=0, column=1)

        tk.Label(bottom_frame, width=20, textvariable=invoices["Fee"], font=self.conf["font"]).grid(row=0, column=2)
        tk.Label(bottom_frame, width=20, textvariable=invoices["in.GST"], font=self.conf["font"]).grid(row=0, column=3)

        self.data["Project Info"]["Project"]["Quotation Number"].trace("w", self._config_entry)
        self.data["State"]["Email to Client"].trace("w", self.auto_lock)

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

            self._config_entry()
            self.fee_frames[service].pack()
            self.update_sum()
        else:
            # archive[service] = invoices_details.pop(service)
            invoices_details[service]["Include"].set(False)
            # invoices_details[service]["Expand"].set(False)
            self.fee_frames[service].pack_forget()
            self.update_sum()

    def today(self):
        self.data["Fee Proposal"]["Reference"]["Date"].set(date.today().strftime("%d-%b-%Y"))
    def installation_today(self):
        self.data["Fee Proposal"]["Installation Reference"]["Date"].set(date.today().strftime("%d-%b-%Y"))

    def _append_value(self, service, extra):
        scope_dir = os.path.join(self.conf["database_dir"], "scope_of_work.json")
        scope_type = self.data["Project Info"]["Project"]["Proposal Type"].get()
        scope = self.data["Fee Proposal"]["Scope"][scope_type]
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

    def _append_stage(self, i):
        stage_dir = os.path.join(self.conf["database_dir"], "general_scope_of_staging.json")
        app_stage = self.data["Fee Proposal"]["Stage"]
        item = self.append_stage[f"Stage{i+1}"]["Item"].get()
        if len(item.strip()) == 0:
            self.messagebox.show_error(title="Error", message="You need to enter some context")
            return
        app_stage[f"Stage{i+1}"]["Items"].append(
            {
                "Include": tk.BooleanVar(value=True),
                "Item": tk.StringVar(value=item)
            }
        )
        tk.Entry(self.stage_frames[f"Stage{i+1}"], width=100,
                 textvariable=app_stage[f"Stage{i+1}"]["Items"][-1]["Item"],
                 font=self.conf["font"]).grid(row=len(app_stage[f"Stage{i+1}"]["Items"]), column=1)

        tk.Checkbutton(self.stage_frames[f"Stage{i+1}"],
                       variable=app_stage[f"Stage{i+1}"]["Items"][-1]["Include"]).grid(row=len(app_stage[f"Stage{i+1}"]["Items"]), column=0)

        if self.append_stage[f"Stage{i+1}"]["Add"].get():
            stage_data = json.load(open(stage_dir))
            stage_data[f"Stage{i+1}"].append(item)
            with open(stage_dir, "w") as f:
                json_object = json.dumps(stage_data, indent=4)
                f.write(json_object)
        self.append_stage[f"Stage{i+1}"]["Item"].set("")
        self.append_stage[f"Stage{i+1}"]["Add"].set(False)

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

    def _update_frame(self, *args):
        if self.data["Project Info"]["Project"]["Service Type"]["Installation"]["Include"].get():
            self.installation_frame.grid(row=0, column=0, sticky="ew", padx=20)
        else:
            self.installation_frame.grid_forget()

    def calculate_apt_price(self, i, entry, type, *args):
        entry.delete(0, tk.END)
        if len(self.data["Fee Proposal"]["Calculation Part"][type].get()) == 0 or len(str(i)) == 0:
            entry.insert(0, "")
        elif not isfloat(self.data["Fee Proposal"]["Calculation Part"][type].get()) or not isfloat(str(i)):
            entry.insert(0, "Error")
        else:
            i = float(i)
            entry.insert(0, round(i*float(self.data["Fee Proposal"]["Calculation Part"][type].get()), 2))

    def level_factor_calculation(self, car_park_level, level_factor, *args):
        car_park_level = car_park_level.get()
        if len(car_park_level) == 0:
            level_factor.set("0")
        elif not car_park_level.isdigit():
            level_factor.set("Error")
        else:
            level_factor.set(1+(float(car_park_level)-1)*0.25 if float(car_park_level)>0 else 0)
    def carport_factor_calculation(self, number_of_carports, carport_factor, *args):
        number_of_carports = number_of_carports.get()
        if len(number_of_carports) == 0:
            carport_factor.set("0")
        elif not number_of_carports.isdigit():
            carport_factor.set("Error")
        else:
            carport_factor.set(round(float(number_of_carports)/90, 2))

    def complex_factor_calculation(self, level_factor, carport_factor, complex_factor, *args):
        level_factor = level_factor.get()
        carport_factor = carport_factor.get()
        if level_factor == "Error" or carport_factor == "Error":
            complex_factor.set("Error")
        else:
            complex_factor.set(round((float(level_factor)*float(carport_factor)), 2))

    def cfd_cost_calculation(self, complex_factor, cfd_cost,*args):
        complex_factor = complex_factor.get()
        if complex_factor == "Error":
            cfd_cost.set("Error")
        else:
            complex_factor = float(complex_factor)
            if complex_factor<=0:
                cfd_cost.set("0")
            elif complex_factor>0 and complex_factor<=0.5:
                cfd_cost.set("1000")
            elif complex_factor>0.5 and complex_factor<=1:
                cfd_cost.set("2000")
            elif complex_factor>1 and complex_factor<=5:
                cfd_cost.set("3000")
            elif complex_factor>5 and complex_factor<=10:
                cfd_cost.set("4000")
            elif complex_factor>10 and complex_factor<=15:
                cfd_cost.set("5000")
            else:
                cfd_cost.set("6000")
    def _config_entry(self, *args):
        if self.data["Lock"]["Proposal"].get():
            self.reference_dic["Date"].config(state=tk.DISABLED)
            self.reference_dic["Today"].config(state=tk.DISABLED)
            self.reference_dic["Revision"].config(state=tk.DISABLED)
            self.installation_dic["Date"].config(state=tk.DISABLED)
            self.installation_dic["Today"].config(state=tk.DISABLED)
            self.installation_dic["Revision"].config(state=tk.DISABLED)
            # self.installation_dic["Program"].config(state=tk.DISABLED)
            self.time_dic["Fee Proposal"].config(state=tk.DISABLED)
            self.time_dic["Pre-design"].config(state=tk.DISABLED)
            self.time_dic["Documentation"].config(state=tk.DISABLED)
            for stage in self.stage_dic.keys():
                self.stage_dic[stage]["Include"].config(state=tk.DISABLED)
                self.stage_dic[stage]["Service"].config(state=tk.DISABLED)
                for item in self.stage_dic[stage]["Items"]:
                    item["Include"].config(state=tk.DISABLED)
                    item["Item"].config(state=tk.DISABLED)
            for service in self.fee_dic.values():
                service["Service"].config(state=tk.DISABLED)
                service["Fee"].config(state=tk.DISABLED)
                service["Expand"].config(state=tk.DISABLED)
                for content in service["Content"]["Details"]:
                    content["Service"].config(state=tk.DISABLED)
                    content["Fee"].config(state=tk.DISABLED)
            # for service in self.scope_dic.keys():
            #     for extra in self.scope_dic[service].keys():
            #         for item in self.scope_dic[service][extra]:
            #             item["Checkbutton"].config(state=tk.DISABLED)
            #             item["Entry"].config(state=tk.DISABLED)
        else:
            self.reference_dic["Date"].config(state=tk.NORMAL)
            self.reference_dic["Today"].config(state=tk.NORMAL)
            self.reference_dic["Revision"].config(state=tk.NORMAL)
            self.installation_dic["Date"].config(state=tk.NORMAL)
            self.installation_dic["Today"].config(state=tk.NORMAL)
            self.installation_dic["Revision"].config(state=tk.NORMAL)
            # self.installation_dic["Program"].config(state=tk.NORMAL)
            self.time_dic["Fee Proposal"].config(state=tk.NORMAL)
            self.time_dic["Pre-design"].config(state=tk.NORMAL)
            self.time_dic["Documentation"].config(state=tk.NORMAL)
            for stage in self.stage_dic.keys():
                self.stage_dic[stage]["Include"].config(state=tk.NORMAL)
                self.stage_dic[stage]["Service"].config(state=tk.NORMAL)
                for item in self.stage_dic[stage]["Items"]:
                    item["Include"].config(state=tk.NORMAL)
                    item["Item"].config(state=tk.NORMAL)
            for service in self.fee_dic.values():
                service["Service"].config(state=tk.NORMAL)
                service["Fee"].config(state=tk.NORMAL)
                service["Expand"].config(state=tk.NORMAL)
                for content in service["Content"]["Details"]:
                    content["Service"].config(state=tk.NORMAL)
                    content["Fee"].config(state=tk.NORMAL)
            # for service in self.scope_dic.keys():
            #     for extra in self.scope_dic[service].keys():
            #         for item in self.scope_dic[service][extra]:
            #             item["Entry"].config(state=tk.NORMAL)
            #             item["Checkbutton"].config(state=tk.NORMAL)
    def unlock(self):
        self.data["Lock"]["Proposal"].set(not self.data["Lock"]["Proposal"].get())

    def config_lock_button(self, *args):
        if self.data["Lock"]["Proposal"].get():
            self.proposal_lock.config(text="Unlock")
        else:
            self.proposal_lock.config(text="Lock")
    def auto_lock(self, *args):
        self.data["Lock"]["Proposal"].set(self.data["State"]["Email to Client"].get())

    def set_variable(self,var1, var2, *args):
        var1.set(var2.get())
        # for service in self.fee_dic.values():
        #     service["Fee"].config(state=tk.NORMAL)
        #     for content in service["Content"]["Details"]:
        #         content["Fee"].config(state=tk.NORMAL)