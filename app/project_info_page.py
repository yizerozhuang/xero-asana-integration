import tkinter as tk
from tkinter import ttk, messagebox

from utility import load, get_quotation_number, save, reset, finish_setup, delete_project, config_state, config_log

import os
import webbrowser


class ProjectInfoPage(tk.Frame):
    def __init__(self, parent, app):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.app = app
        self.app.data["Project Info"] = dict()
        self.data = app.data
        self.conf = app.conf

        self.main_frame = tk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=1)
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

        self.main_log_frame = tk.LabelFrame(self.main_canvas, text="Log")
        self.main_canvas.create_window((800, -600), window=self.main_log_frame, anchor="nw")

        self.log_part()
        self.project_part()
        self.client_part()
        self.main_contact_part()
        self.building_features_part()
        self.drawing_number_part()
        self.finish_part()

    def log_part(self):
        log_frame = tk.Frame(self.main_log_frame)
        log_frame.pack()
        tk.Label(log_frame, textvariable=self.app.log_text, font=self.conf["font"], justify=tk.LEFT).pack(anchor="w")

    def project_part(self):
        # User_information frame
        project = dict()
        self.data["Project Info"]["Project"] = project

        project_frame = tk.LabelFrame(self.main_context_frame, text="Project Information", font=self.conf["font"])
        project_frame.pack(padx=20)

        tk.Label(project_frame, width=30, text="Project Name", font=self.conf["font"]).grid(row=0, column=0,
                                                                                            padx=(10, 0))
        project["Project Name"] = tk.StringVar()
        tk.Entry(project_frame, width=70, font=self.conf["font"], fg="blue", textvariable=project["Project Name"]).grid(
            row=0,
            column=1,
            columnspan=3,
            padx=(0, 10))
        # self.project_list = ttk.Combobox(project_frame,font=self.conf["font"], width=85, textvariable=project["Project Name"])
        # self.project_list.grid(row=0, column=1, columnspan=3, padx=(0, 10))
        # project["Project Name"].trace("w", self._search_projects)

        tk.Label(project_frame, width=30, text="Quotation Number", font=self.conf["font"]).grid(row=1, column=0,
                                                                                                padx=(10, 0))
        project["Quotation Number"] = tk.StringVar()
        tk.Entry(project_frame, width=34, font=self.conf["font"], fg="blue",
                 textvariable=project["Quotation Number"]).grid(row=1,
                                                                column=1,
                                                                columnspan=2,
                                                                padx=(0, 10))
        save_load_frame = tk.Frame(project_frame)
        save_load_frame.grid(row=1, column=3, rowspan=2)

        tk.Button(save_load_frame, width=10, height=2, text="Reset", command=lambda: reset(self.app), bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=0)
        tk.Button(save_load_frame, width=8, height=2, text="Load", command=self.load_data, bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=1)
        tk.Button(save_load_frame, width=10, height=2, text="Open Folder", command=self.open_folder, bg="brown",
                  fg="white",
                  font=self.conf["font"]).grid(row=0, column=2)


        tk.Label(project_frame, width=30, text="Project Number", font=self.conf["font"]).grid(row=2, column=0,
                                                                                              padx=(10, 0))
        project["Project Number"] = tk.StringVar()
        tk.Entry(project_frame, width=34, font=self.conf["font"], fg="blue",
                 textvariable=project["Project Number"]).grid(row=2, column=1, columnspan=2, padx=(0, 10))

        project["Project Name"].trace("w", self._update_quotation_number)

        tk.Label(project_frame, width=30, text="Shop Name", font=self.conf["font"]).grid(row=3, column=0, padx=(10, 0),
                                                                                         pady=(0, 10))
        project["Shop Name"] = tk.StringVar()
        tk.Entry(project_frame, width=70, font=self.conf["font"], fg="blue", textvariable=project["Shop Name"]).grid(
            row=3,
            column=1,
            columnspan=3,
            padx=(0, 10), pady=(0, 10))

        tk.Label(project_frame, text="Project Type", font=self.conf["font"]).grid(row=4, column=0)
        project_types = ["Restaurant", "Office", "Commercial", "Group house", "Apartment", "Mixed-use complex",
                         "School", "Others"]

        project["Project Type"] = tk.StringVar(value="Restaurant")

        for i, types in enumerate(project_types):
            button = tk.Radiobutton(project_frame, text=types, variable=project["Project Type"], value=types,
                                    font=self.conf["font"])
            if i < 3:
                button.grid(row=4, column=i + 1, sticky="W")
            elif i < 6:
                button.grid(row=5, column=i - 2, sticky="W")
            else:
                button.grid(row=6, column=i - 5, sticky="W")

        service_types = ["Mechanical Service", "CFD Service", "Electrical Service", "Hydraulic Service", "Fire Service",
                         "Kitchen Ventilation", "Mech Review",
                         "Installation", "Miscellaneous"]

        tk.Label(project_frame, width=30, text="Service Type", font=self.conf["font"]).grid(row=7, column=0)

        project["Service Type"] = []
        for i, service in enumerate(service_types):
            project["Service Type"].append(
                {
                    "Service": tk.StringVar(value=service),
                    "Include": tk.BooleanVar()
                }
            )
            button = tk.Checkbutton(project_frame, text=service, variable=project["Service Type"][i]["Include"],
                                    font=self.conf["font"])
            if i < 3:
                button.grid(row=7, column=i + 1, sticky="W")
            elif i < 6:
                button.grid(row=8, column=i - 2, sticky="W")
            else:
                button.grid(row=9, column=i - 5, sticky="W")

        project["Service Type"][0]["Include"].trace("w",
                                                    lambda a, b, c: self._update_service(project["Service Type"][0]))
        project["Service Type"][2]["Include"].trace("w",
                                                    lambda a, b, c: self._update_service(project["Service Type"][2]))
        project["Service Type"][3]["Include"].trace("w",
                                                    lambda a, b, c: self._update_service(project["Service Type"][3]))
        project["Service Type"][4]["Include"].trace("w",
                                                    lambda a, b, c: self._update_service(project["Service Type"][4]))


    def client_part(self):
        client = dict()
        self.data["Project Info"]["Client"] = client

        # client frame
        client_frame = tk.LabelFrame(self.main_context_frame, text="Client", font=self.conf["font"])
        client_frame.pack(padx=20)

        tk.Label(client_frame, width=30, text="Client Full Name", font=self.conf["font"]).grid(row=0, column=0,
                                                                                               padx=(10, 0))
        client["Client Full Name"] = tk.StringVar()
        tk.Entry(client_frame, width=70, font=self.conf["font"], fg="blue",
                 textvariable=client["Client Full Name"]).grid(row=0, column=1, columnspan=3, padx=(0, 10))

        tk.Label(client_frame, width=30, text="Client Contact Type", font=self.conf["font"]).grid(row=1, column=0)

        contact_type = ["Architect", "Builder", "Certifier", "Contractor", "Developer", "Engineer", "Government", "RDM",
                        "Strata", "Supplier", "Owner/Tenant", "Others"]

        client["Client Contact Type"] = tk.StringVar(value="Architect")
        for i, types in enumerate(contact_type):
            button = tk.Radiobutton(client_frame, text=types, variable=client["Client Contact Type"], value=types,
                                    font=self.conf["font"])
            if i < 3:
                button.grid(row=1, column=i + 1, sticky="W")
            elif i < 6:
                button.grid(row=2, column=i - 2, sticky="W")
            elif i < 9:
                button.grid(row=3, column=i - 5, sticky="W")
            else:
                button.grid(row=4, column=i - 8, sticky="W")
        client_info = ["Client Company", "Client Address", "Client ABN", "Contact Number",
                       "Contact Email"]
        for i, info in enumerate(client_info):
            tk.Label(client_frame, width=30, text=info, font=self.conf["font"]).grid(row=i + 5, column=0)
            client[info] = tk.StringVar()
            tk.Entry(client_frame, width=70, font=self.conf["font"], fg="blue", textvariable=client[info]).grid(
                row=i + 5, column=1, columnspan=3)

    def main_contact_part(self):
        main_contact = dict()
        self.data["Project Info"]["Main Contact"] = main_contact

        contact_frame = tk.LabelFrame(self.main_context_frame, text="Main Contact", font=self.conf["font"])
        contact_frame.pack(padx=20)

        tk.Label(contact_frame, width=30, text="Main Contact Full Name", font=self.conf["font"]).grid(row=0, column=0,
                                                                                                      padx=(10, 0))

        main_contact["Main Contact Full Name"] = tk.StringVar()
        tk.Entry(contact_frame, width=70, font=self.conf["font"], fg="blue",
                 textvariable=main_contact["Main Contact Full Name"]).grid(row=0, column=1, columnspan=3, padx=(0, 10))

        tk.Label(contact_frame, text="Main Contact Contact Type", font=self.conf["font"]).grid(row=1, column=0)

        contact_type = ["Architect", "Builder", "Certifier", "Contractor", "Developer", "Engineer", "Government", "RDM",
                        "Strata", "Supplier", "Owner/Tenant", "Others"]

        main_contact["Main Contact Contact Type"] = tk.StringVar(value="Architect")
        for i, types in enumerate(contact_type):
            button = tk.Radiobutton(contact_frame, text=types, variable=main_contact["Main Contact Contact Type"],
                                    value=types, font=self.conf["font"])
            if i < 3:
                button.grid(row=1, column=i + 1, sticky="W")
            elif i < 6:
                button.grid(row=2, column=i - 2, sticky="W")
            elif i < 9:
                button.grid(row=3, column=i - 5, sticky="W")
            else:
                button.grid(row=4, column=i - 8, sticky="W")

        contact_info = ["Main Contact Company", "Main Contact Address", "Main Contact ABN",
                        "Main Contact Number", "Main Contact Email"]
        for i, info in enumerate(contact_info):
            tk.Label(contact_frame, width=30, text=info, font=self.conf["font"]).grid(row=i + 5, column=0)
            main_contact[info] = tk.StringVar(name=info)
            tk.Entry(contact_frame, width=70, font=self.conf["font"], fg="blue", textvariable=main_contact[info]).grid(
                row=i + 5, column=1, columnspan=3)

    def building_features_part(self):
        n_building = self.conf["n_building"]
        building_features = {
            "Details": [
                {
                    "Levels": tk.StringVar(),
                    "Space/room Description": tk.StringVar(),
                    "Area/m^2": tk.StringVar()
                } for _ in range(n_building)
            ],
            "Feature/Notes": tk.StringVar(),
            "Total Area": tk.StringVar(value="0")
        }
        self.data["Project Info"]["Building Features"] = building_features

        build_feature_frame = tk.LabelFrame(self.main_context_frame, text="Building Features", font=self.conf["font"])
        build_feature_frame.pack(padx=20)

        tk.Label(build_feature_frame, text="Levels", font=self.conf["font"]).grid(row=0, column=0)
        tk.Label(build_feature_frame, text="Space/room Description", font=self.conf["font"]).grid(row=0, column=1)
        tk.Label(build_feature_frame, text="Area/m^2", font=self.conf["font"]).grid(row=0, column=2)

        tk.Label(build_feature_frame, text="Feature/Notes", width=30, font=self.conf["font"]).grid(row=n_building + 1,
                                                                                                   column=0)
        tk.Entry(build_feature_frame, width=72, font=self.conf["font"], fg="blue",
                 textvariable=building_features["Feature/Notes"]).grid(row=n_building + 1, column=1, columnspan=2)
        tk.Label(build_feature_frame, text="Total", font=self.conf["font"]).grid(row=n_building + 2, column=0)
        tk.Label(build_feature_frame, textvariable=building_features["Total Area"], font=self.conf["font"]).grid(
            row=n_building + 2, column=2)

        for i in range(n_building):
            tk.Entry(build_feature_frame, width=34, font=self.conf["font"], fg="blue",
                     textvariable=building_features["Details"][i]["Levels"]).grid(row=i + 1, column=0, padx=(10, 0))
            tk.Entry(build_feature_frame, width=50, font=self.conf["font"], fg="blue",
                     textvariable=building_features["Details"][i]["Space/room Description"]).grid(row=i + 1, column=1)
            tk.Entry(build_feature_frame, width=20, font=self.conf["font"], fg="blue",
                     textvariable=building_features["Details"][i]["Area/m^2"]).grid(row=i + 1, column=2, padx=(0, 10))

            building_features["Details"][i]["Area/m^2"].trace("w", self._update_area)

    def drawing_number_part(self):
        n_drawing = self.conf["n_drawing"]
        drawing = [
            {
                "Drawing Number": tk.StringVar(),
                "Drawing Name": tk.StringVar(),
                "Revision": tk.StringVar()
            } for _ in range(n_drawing)
        ]
        self.app.data["Project Info"]["Drawing"] = drawing

        drawing_number_frame = tk.LabelFrame(self.main_context_frame, text="Drawing Number", font=self.conf["font"])
        drawing_number_frame.pack(padx=20)

        tk.Label(drawing_number_frame, text="Drawing Number", font=self.conf["font"]).grid(row=0, column=0)
        tk.Label(drawing_number_frame, text="Drawing Name", font=self.conf["font"]).grid(row=0, column=1)
        tk.Label(drawing_number_frame, text="Revision", font=self.conf["font"]).grid(row=0, column=2)

        for i in range(n_drawing):
            tk.Entry(drawing_number_frame, width=34, font=self.conf["font"], fg="blue",
                     textvariable=drawing[i]["Drawing Number"]).grid(row=i + 1, column=0, padx=(10, 0))
            tk.Entry(drawing_number_frame, width=50, font=self.conf["font"], fg="blue",
                     textvariable=drawing[i]["Drawing Name"]).grid(row=i + 1, column=1)
            tk.Entry(drawing_number_frame, width=20, font=self.conf["font"], fg="blue",
                     textvariable=drawing[i]["Revision"]).grid(row=i + 1, column=2, padx=(0, 10))

        drawing[0]["Drawing Number"].set("M-000")
        drawing[0]["Drawing Name"].set("Cover Sheet")
        drawing[0]["Revision"].set("A")
        drawing[1]["Drawing Number"].set("M-001")
        drawing[1]["Drawing Name"].set("Tenancy Level Layout")
        drawing[1]["Revision"].set("A")
        drawing[2]["Drawing Number"].set("M-002")
        drawing[2]["Drawing Name"].set("Roof Layout")
        drawing[2]["Revision"].set("A")

    def finish_part(self):
        finish_frame = tk.LabelFrame(self.main_context_frame)
        finish_frame.pack(padx=20, side=tk.RIGHT)

        tk.Button(finish_frame, text="Delete Project", command=self._delete_project,
                  bg="brown", fg="white", font=self.conf["font"]).pack(side=tk.RIGHT)

        self.quote_text = tk.StringVar(value="Quote Unsuccessful")
        tk.Button(finish_frame, textvariable=self.quote_text, command=self.quote_unsuccessful,
                  bg="brown", fg="white", font=self.conf["font"]).pack(side=tk.RIGHT)

        self.data["State"]["Quote Unsuccessful"].trace("w", self._update_quote_button_text)

        tk.Button(finish_frame, text="Finish Set Up", command=self._finish_setup,
                  bg="brown", fg="white", font=self.conf["font"]).pack(side=tk.RIGHT)

    def _update_quote_button_text(self, *args):
        if self.data["State"]["Quote Unsuccessful"].get():
            self.quote_text.set("Restore")
        else:
            self.quote_text.set("Quote Unsuccessful")

    def _delete_project(self):
        if len(self.data["Project Info"]["Project"]["Quotation Number"].get())==0:
            messagebox.showerror("Error", "You cant delete an empty project")
            return
        delete = messagebox.askretrycancel('Warming', "Are you sure you want to delete the Project? only the admin can restore after you delete the project")
        if delete:
            delete_project(self.app)
            messagebox.showinfo("Deleted", f"Project {self.data['Project Info']['Project']['Quotation Number'].get()} deleted")

    def load_data(self):
        quotation_number = self.data["Project Info"]["Project"]["Quotation Number"].get().upper()
        database_dir = os.path.join(self.conf["database_dir"], quotation_number)
        if len(quotation_number) == 0:
            messagebox.showerror("Error", "Please enter a Quotation Number before you load")
        elif not os.path.exists(database_dir):
            messagebox.showerror("Error", f"The quotation {quotation_number} number doesn't exist")
        else:
            load(self.app)

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

    def _finish_setup(self):
        if len(self.data["Project Info"]["Project"]["Quotation Number"].get()) == 0:
            messagebox.showerror("Error", "Please Create an quotation Number first")
            return

        finish_setup(self.app)

    def quote_unsuccessful(self):
        if self.data["State"]["Quote Unsuccessful"].get():
            restore = messagebox.askyesno("Warming", "Are you sure you want to restore this project")
            if restore:
                self.data["State"]["Quote Unsuccessful"].set(False)
                save(self.app)
                config_state(self.app)
                self.app.log.log_restore(self.app)
                config_log(self.app)
        else:
            quote = messagebox.askyesno("Warming", "Are you sure you want to put this project in Quote Unsuccessful")
            if quote:
                self.data["State"]["Quote Unsuccessful"].set(True)
                save(self.app)
                config_state(self.app)
                self.app.log.log_unsuccessful(self.app)
                config_log(self.app)

    def _update_quotation_number(self, *args):
        if self.data["Project Info"]["Project"]["Quotation Number"].get() != "":
            return
        self.data["Project Info"]["Project"]["Quotation Number"].set(get_quotation_number())

    # def _search_projects(self):
    #     os.listdir(self.conf["working_dir"])

    def _update_area(self, *args):
        area_sum = 0
        for drawing in self.data["Project Info"]["Building Features"]["Details"]:
            if len(drawing["Area/m^2"].get()) == 0:
                continue
            try:
                area_sum += float(drawing["Area/m^2"].get())
            except ValueError:
                self.data["Project Info"]["Building Features"]["Total Area"].set("Error")
                return
        self.data["Project Info"]["Building Features"]["Total Area"].set(str(round(area_sum, 2)))

    def _update_service(self, var, *args):
        self.app.fee_proposal_page.update_scope(var)
        self.app.fee_proposal_page.update_fee(var)
        self.app.financial_panel_page.update_invoice(var)
        # self.app.invoice_page.update_bill(var)
        # self.app.invoice_page.update_profit(var)

    def _on_mousewheel(self, event):
        self.main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
