import tkinter as tk
from tkinter import ttk

from utility import load_data, get_quotation_number, save, reset, delete_project, config_state, config_log
from text_extension import TextExtension
from asana_function import update_asana

import os

class ProjectInfoPage(tk.Frame):
    def __init__(self, parent, app):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.app = app
        self.app.data["Project Info"] = dict()
        self.data = app.data
        self.conf = app.conf
        self.messagebox = self.app.messagebox

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
        self.main_canvas.create_window((0, 0), window=self.main_context_frame, anchor="nw")
        self.main_context_frame.bind("<Configure>", self.reset_scrollregion)

        # self.main_log_frame = tk.LabelFrame(self.main_canvas, text="Log")
        # self.main_canvas.create_window((800, -600), window=self.main_log_frame, anchor="nw")

        # self.search_part()
        self.log_part()
        self.project_part()
        self.client_part()
        self.main_contact_part()
        self.building_features_part()
        self.drawing_number_part()
        self.finish_part()

    def log_part(self):
        # log_frame = tk.Frame(self.main_context_frame)
        # log_frame.grid(row=0, column=1, column_span=6)
        # tk.Entry(log_frame, textvariable=self.data["Email_Content"], font=self.conf["font"], justify=tk.LEFT).pack(anchor="w")

        # self.app.email_text = tk.Text(self.main_context_frame, font=self.conf["font"], height=68)
        # self.app.email_text.grid(row=0, column=1, rowspan=6, sticky="n")

        TextExtension(self.main_context_frame, textvariable=self.data["Email_Content"], font=self.conf["font"], fg="blue", height=85).grid(row=0, column=1, rowspan=6, sticky="n")

        tk.Label(self.main_context_frame, textvariable=self.app.log_text, font=self.conf["font"], justify=tk.LEFT).grid(row=0, column=2, rowspan=6, sticky="n")

    # def search_part(self):
    #     search_frame = tk.LabelFrame(self.main_context_frame)
    #     search_frame.pack(padx=10)
    #     tk.Label(search_frame, text="Search Bar", font=self.conf["font"], width=20).grid(row=0, column=0)
    #     tk.Entry(search_frame, width=100).grid(row=0, column=1)

    def project_part(self):
        # User_information frame
        project = dict()
        self.data["Project Info"]["Project"] = project

        project_frame = tk.LabelFrame(self.main_context_frame, text="Project Information", font=self.conf["font"])
        project_frame.grid(row=0, column=0)

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
        tk.Entry(project_frame, width=45, font=self.conf["font"], fg="blue",
                 textvariable=project["Quotation Number"]).grid(row=1,
                                                                column=1,
                                                                columnspan=2,
                                                                padx=(0, 10))

        tk.Label(project_frame, width=30, text="Project Number", font=self.conf["font"]).grid(row=2, column=0,
                                                                                              padx=(10, 0))
        project["Project Number"] = tk.StringVar()
        tk.Entry(project_frame, width=45, font=self.conf["font"], fg="blue",
                 textvariable=project["Project Number"]).grid(row=2, column=1, columnspan=2, padx=(0, 10))


        save_load_frame = tk.Frame(project_frame)
        save_load_frame.grid(row=1, column=3, rowspan=2)


        tk.Button(save_load_frame, width=10, height=2, text="Clear Up", command=lambda: reset(self.app), bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=0)
        tk.Button(save_load_frame, width=8, height=2, text="Load", command=self.load_data, bg="brown", fg="white",
                  font=self.conf["font"]).grid(row=0, column=1)

        project["Project Name"].trace("w", self._update_quotation_number)

        tk.Label(project_frame, width=30, text="Shop Name", font=self.conf["font"]).grid(row=3, column=0, padx=(10, 0),
                                                                                         pady=(0, 10))
        project["Shop Name"] = tk.StringVar()
        tk.Entry(project_frame, width=70, font=self.conf["font"], fg="blue", textvariable=project["Shop Name"]).grid(
            row=3,
            column=1,
            columnspan=3,
            padx=(0, 10), pady=(0, 10))

        project["Proposal Type"] = tk.StringVar(value="Minor")

        tk.Label(project_frame, text="Proposal Type", font=self.conf["font"]).grid(row=4, column=0)
        tk.Radiobutton(project_frame, text="Minor Project", variable=project["Proposal Type"], value="Minor", font=self.conf["font"]).grid(row=4, column=1, sticky="W")
        tk.Radiobutton(project_frame, text="Major Project", variable=project["Proposal Type"], value="Major", font=self.conf["font"]).grid(row=4, column=2, sticky="W")

        project["Proposal Type"].trace("w", self._update_proposal)

        tk.Label(project_frame, text="Project Type", font=self.conf["font"]).grid(row=5, column=0)
        project_types = ["Restaurant", "Office", "Commercial", "Group House", "Apartment", "Mixed-use complex",
                         "School", "Others"]

        project["Project Type"] = tk.StringVar(value="Restaurant")

        for i, types in enumerate(project_types):
            button = tk.Radiobutton(project_frame, text=types, variable=project["Project Type"], value=types,
                                    font=self.conf["font"])
            button.grid(row=i//3+5, column=i%3+1, sticky="W")

        service_types = ["Mechanical Service", "CFD Service", "Electrical Service", "Hydraulic Service",
                         "Fire Service", "Kitchen Ventilation", "Mechanical Review", "Miscellaneous", "Installation"]

        tk.Label(project_frame, width=30, text="Service Type", font=self.conf["font"]).grid(row=8, column=0)

        project["Service Type"] = {}
        for i, service in enumerate(service_types):
            project["Service Type"][service] = {
                "Service": tk.StringVar(value=service),
                "Include": tk.BooleanVar()
            }
            button = tk.Checkbutton(project_frame, text=service, variable=project["Service Type"][service]["Include"],
                                    font=self.conf["font"])
            button.grid(row=i//3+8, column=i%3+1, sticky="W")

        update_service_fuc = lambda s : lambda a, b, c: self._update_service(project["Service Type"][s])
        for service in self.conf["service_list"]:
            project["Service Type"][service]["Include"].trace("w", update_service_fuc(service))
        # project["Service Type"]["Variation"] = {
        #         "Service": tk.StringVar(value="Variation"),
        #         "Include": tk.BooleanVar(value=True)
        # }

        # project["Service Type"]["Mechanical Service"]["Include"].trace("w", lambda a, b, c: self._update_service(project["Service Type"]["Mechanical Service"]))
        # project["Service Type"]["Electrical Service"]["Include"].trace("w", lambda a, b, c: self._update_service(project["Service Type"]["Electrical Service"]))
        # project["Service Type"]["Hydraulic Service"]["Include"].trace("w", lambda a, b, c: self._update_service(project["Service Type"]["Hydraulic Service"]))
        # project["Service Type"]["Fire Service"]["Include"].trace("w", lambda a, b, c: self._update_service(project["Service Type"]["Fire Service"]))

    def client_part(self):
        client = dict()
        self.data["Project Info"]["Client"] = client

        # client frame
        client_frame = tk.LabelFrame(self.main_context_frame, text="Client", font=self.conf["font"])
        client_frame.grid(row=1, column=0)

        tk.Radiobutton(client_frame, value="Client", variable=self.data["Address_to"], width=7).grid(row=0, column=0)

        tk.Label(client_frame, width=20, text="Client Full Name", font=self.conf["font"]).grid(row=0, column=1)
        client["Full Name"] = tk.StringVar()
        tk.Entry(client_frame, width=70, font=self.conf["font"], fg="blue",
                 textvariable=client["Full Name"]).grid(row=0, column=2, columnspan=3, padx=(0, 10))

        tk.Label(client_frame, width=20, text="Client Contact Type", font=self.conf["font"]).grid(row=1, column=1)

        contact_type = ["Architect", "Builder", "Certifier", "Contractor", "Developer", "Engineer", "Government", "RDM",
                        "Strata", "Supplier", "Owner/Tenant", "Others"]

        client["Contact Type"] = tk.StringVar(value="Architect")
        for i, types in enumerate(contact_type):
            button = tk.Radiobutton(client_frame, text=types, variable=client["Contact Type"], value=types,
                                    font=self.conf["font"])
            button.grid(row=i//3 + 1, column=i%3 + 2, sticky="W")

        client_info = ["Company", "Address", "ABN", "Contact Number", "Contact Email"]
        for i, info in enumerate(client_info):
            tk.Label(client_frame, width=20, text=info, font=self.conf["font"]).grid(row=i + 5, column=1)
            client[info] = tk.StringVar()
            tk.Entry(client_frame, width=70, font=self.conf["font"], fg="blue", textvariable=client[info]).grid(
                row=i + 5, column=2, columnspan=3)

    def main_contact_part(self):
        main_contact = dict()
        self.data["Project Info"]["Main Contact"] = main_contact

        contact_frame = tk.LabelFrame(self.main_context_frame, text="Main Contact", font=self.conf["font"])
        contact_frame.grid(row=2, column=0)

        tk.Radiobutton(contact_frame, value="Main Contact", variable=self.data["Address_to"], width=7).grid(row=0, column=0)

        tk.Label(contact_frame, width=20, text="Main Contact Full Name", font=self.conf["font"]).grid(row=0, column=1)

        main_contact["Full Name"] = tk.StringVar()
        tk.Entry(contact_frame, width=70, font=self.conf["font"], fg="blue",
                 textvariable=main_contact["Full Name"]).grid(row=0, column=2, columnspan=3, padx=(0, 10))

        tk.Label(contact_frame, text="Main Contact Contact Type", font=self.conf["font"]).grid(row=1, column=1)

        contact_type = ["Architect", "Builder", "Certifier", "Contractor", "Developer", "Engineer", "Government", "RDM",
                        "Strata", "Supplier", "Owner/Tenant", "Others"]

        main_contact["Contact Type"] = tk.StringVar(value="Architect")
        for i, types in enumerate(contact_type):
            button = tk.Radiobutton(contact_frame, text=types, variable=main_contact["Contact Type"],
                                    value=types, font=self.conf["font"])
            button.grid(row=i//3+1, column=i%3+2, sticky="W")

        contact_info = ["Company", "Address", "ABN",
                        "Contact Number", "Contact Email"]
        for i, info in enumerate(contact_info):
            tk.Label(contact_frame, width=20, text=info, font=self.conf["font"]).grid(row=i + 5, column=1)
            main_contact[info] = tk.StringVar(name=info)
            tk.Entry(contact_frame, width=70, font=self.conf["font"], fg="blue", textvariable=main_contact[info]).grid(
                row=i + 5, column=2, columnspan=3)

    def building_features_part(self):
        n_building = self.conf["n_building"]
        building_features = {
            "Minor": {
                "Details": [
                    {
                        "Levels": tk.StringVar(),
                        "Space": tk.StringVar(),
                        "Area": tk.StringVar()
                    } for _ in range(n_building)
                ],
                "Total Area": tk.StringVar(value="0"),
            },
            "Major": {
                "Car Park": [[tk.StringVar() for _ in range(8)] for _ in range(self.conf["n_major_building"])],
                "Block": [[tk.StringVar() for _ in range(8)] for _ in range(self.conf["n_major_building"])]
            },
            "Feature": tk.StringVar(),
            "Notes": tk.StringVar()
        }
        self.data["Project Info"]["Building Features"] = building_features

        build_feature_frame = tk.LabelFrame(self.main_context_frame, text="Building Features", font=self.conf["font"])
        build_feature_frame.grid(row=3, column=0)

        # tk.Radiobutton(build_feature_frame, text="Minor", variable=building_features["Type"], value="Minor", font=self.conf["font"]).grid(row=0, column=0, sticky="w")
        # tk.Radiobutton(build_feature_frame, text="Major", variable=building_features["Type"], value="Major", font=self.conf["font"]).grid(row=2, column=0, sticky="w")
        self.minor_frame = tk.Frame(build_feature_frame)
        tk.Label(self.minor_frame, text="Levels", font=self.conf["font"]).grid(row=0, column=0)
        tk.Label(self.minor_frame, text="Space/room Description", font=self.conf["font"]).grid(row=0, column=1)
        tk.Label(self.minor_frame, text="Area/m^2", font=self.conf["font"]).grid(row=0, column=2)
        self.minor_frame.grid(row=0, column=0)

        for i in range(n_building):
            tk.Entry(self.minor_frame, width=34, font=self.conf["font"], fg="blue",
                     textvariable=building_features["Minor"]["Details"][i]["Levels"]).grid(row=i + 1, column=0, padx=(10, 0))
            tk.Entry(self.minor_frame, width=50, font=self.conf["font"], fg="blue",
                     textvariable=building_features["Minor"]["Details"][i]["Space"]).grid(row=i + 1, column=1)
            tk.Entry(self.minor_frame, width=20, font=self.conf["font"], fg="blue",
                     textvariable=building_features["Minor"]["Details"][i]["Area"]).grid(row=i + 1, column=2, padx=(0, 10))
            building_features["Minor"]["Details"][i]["Area"].trace("w", self._update_area)

        tk.Label(self.minor_frame, text="Total", font=self.conf["font"]).grid(row=n_building + 1, column=0)
        tk.Label(self.minor_frame, textvariable=building_features["Minor"]["Total Area"], font=self.conf["font"]).grid(row=n_building+1, column=2)


        self.major_frame = tk.Frame(build_feature_frame)

        car_park_frame = tk.Frame(self.major_frame)
        car_park_frame.pack()
        tk.Label(car_park_frame, text="Level", font=self.conf["font"]).grid(row=0, column=0, rowspan=2)
        tk.Label(car_park_frame, text="Other type of room except \n car spots and storage cages",font=self.conf["font"]).grid(row=0, column=1, rowspan=2)
        tk.Label(car_park_frame, text="Car Park A",font=self.conf["font"]).grid(row=0, column=2, columnspan=2)
        tk.Label(car_park_frame, text="Car Spots", font=self.conf["font"]).grid(row=1, column=2)
        tk.Label(car_park_frame, text="Other", font=self.conf["font"]).grid(row=1, column=3)
        tk.Label(car_park_frame, text="Car Park B",font=self.conf["font"]).grid(row=0, column=4, columnspan=2)
        tk.Label(car_park_frame, text="Car Spots", font=self.conf["font"]).grid(row=1, column=4)
        tk.Label(car_park_frame, text="Other", font=self.conf["font"]).grid(row=1, column=5)
        tk.Label(car_park_frame, text="Car Park C",font=self.conf["font"]).grid(row=0, column=6, columnspan=2)
        tk.Label(car_park_frame, text="Car Spots", font=self.conf["font"]).grid(row=1, column=6)
        tk.Label(car_park_frame, text="Other", font=self.conf["font"]).grid(row=1, column=7)
        tk.Label(car_park_frame, text="Total", font=self.conf["font"]).grid(row=0, column=8)

        for i in range(8):
            for j in range(self.conf["n_major_building"]):
                entry = tk.Entry(car_park_frame, width=9, textvariable=building_features["Major"]["Car Park"][j][i], fg="blue", font=self.conf["font"])
                entry.grid(row=j+2, column=i)
                building_features["Major"]["Car Park"][j][i].trace("w", self._update_car_total)
                if i == 0:
                    entry.config(width=10)
                elif i ==1:
                    entry.config(width=30)
        tk.Label(car_park_frame, text="Total Car spot", font=self.conf["font"]).grid(row=self.conf["n_major_building"]+2, column=0, columnspan=2)
        tk.Label(car_park_frame, text="Total Others", font=self.conf["font"]).grid(row=self.conf["n_major_building"]+3, column=0, columnspan=2)

        self.car_part_row_total = [tk.StringVar(value="0") for _ in range(6)]
        self.car_part_column_total = [tk.StringVar(value="0") for _ in range(self.conf["n_major_building"])]
        self.car_part_spot_total = tk.StringVar(value="0")
        self.car_part_other_total = tk.StringVar(value="0")

        for i in range(6):
            tk.Label(car_park_frame, textvariable=self.car_part_row_total[i], font=self.conf["font"]).grid(row=self.conf["n_major_building"]+2+i%2, column=i+2)
        for i in range(self.conf["n_major_building"]):
            tk.Label(car_park_frame, textvariable=self.car_part_column_total[i], font=self.conf["font"]).grid(row=i+2 , column=8)
        tk.Label(car_park_frame, textvariable=self.car_part_spot_total, font=self.conf["font"]).grid(row=self.conf["n_major_building"]+2, column=8)
        tk.Label(car_park_frame, textvariable=self.car_part_other_total, font=self.conf["font"]).grid(row=self.conf["n_major_building"]+3, column=8)


        block_frame = tk.Frame(self.major_frame)
        block_frame.pack()
        tk.Label(block_frame, text="Level", font=self.conf["font"]).grid(row=0, column=0, rowspan=2)
        tk.Label(block_frame, text="Other type of room \n except apartment", font=self.conf["font"]).grid(row=0, column=1, rowspan=2)
        tk.Label(block_frame, text="Block A", font=self.conf["font"]).grid(row=0, column=2, columnspan=2)
        tk.Label(block_frame, text="Apt", font=self.conf["font"]).grid(row=1, column=2)
        tk.Label(block_frame, text="Comm", font=self.conf["font"]).grid(row=1, column=3)
        tk.Label(block_frame, text="Block B", font=self.conf["font"]).grid(row=0, column=4, columnspan=2)
        tk.Label(block_frame, text="Apt", font=self.conf["font"]).grid(row=1, column=4)
        tk.Label(block_frame, text="Comm", font=self.conf["font"]).grid(row=1, column=5)
        tk.Label(block_frame, text="Block C", font=self.conf["font"]).grid(row=0, column=6, columnspan=2)
        tk.Label(block_frame, text="Apt", font=self.conf["font"]).grid(row=1, column=6)
        tk.Label(block_frame, text="Comm", font=self.conf["font"]).grid(row=1, column=7)
        tk.Label(block_frame, text="Total", font=self.conf["font"]).grid(row=0, column=8)

        for i in range(8):
            for j in range(self.conf["n_major_building"]):
                entry = tk.Entry(block_frame, width=9, textvariable=building_features["Major"]["Block"][j][i], fg="blue",font=self.conf["font"])
                entry.grid(row=j+2, column=i)
                building_features["Major"]["Block"][j][i].trace("w", self._update_block_total)
                if i == 0:
                    entry.config(width=10)
                elif i ==1:
                    entry.config(width=30)
        tk.Label(block_frame, text="Total Apt", font=self.conf["font"]).grid(row=self.conf["n_major_building"]+2, column=0, columnspan=2)
        tk.Label(block_frame, text="Total Commercial", font=self.conf["font"]).grid(row=self.conf["n_major_building"]+3, column=0, columnspan=2)

        self.block_row_total = [tk.StringVar(value="0") for _ in range(6)]
        self.block_column_total = [tk.StringVar(value="0") for _ in range(self.conf["n_major_building"])]
        self.block_apt_total = tk.StringVar(value="0")
        self.block_comm_total = tk.StringVar(value="0")

        for i in range(6):
            tk.Label(block_frame, textvariable=self.block_row_total[i], font=self.conf["font"]).grid(row=self.conf["n_major_building"]+2+i%2, column=i+2)
        for i in range(self.conf["n_major_building"]):
            tk.Label(block_frame, textvariable=self.block_column_total[i], font=self.conf["font"]).grid(row=i+2 , column=8)
        tk.Label(block_frame, textvariable=self.block_apt_total, font=self.conf["font"]).grid(row=self.conf["n_major_building"]+2, column=8)
        tk.Label(block_frame, textvariable=self.block_comm_total, font=self.conf["font"]).grid(row=self.conf["n_major_building"]+3, column=8)

        self.data["Project Info"]["Project"]["Proposal Type"].trace("w", self._update_building_frame)

        tk.Label(build_feature_frame, text="Asana Notes: ", font=self.conf["font"]).grid(row=1, column=0, sticky="w", padx=(90, 0))
        tk.Entry(build_feature_frame, width=80, font=self.conf["font"], fg="blue", textvariable=building_features["Feature"]).grid(row=2, column=0)

        tk.Label(build_feature_frame, text="Project Notes: ", font=self.conf["font"]).grid(row=3, column=0, sticky="w", padx=(90, 0))
        TextExtension(build_feature_frame, textvariable=building_features["Notes"], font=self.conf["font"], height=10, fg="blue").grid(row=4, column=0)
        building_features["Notes"].set(
            """
                Email sent from ??? on ???, with architecture drawings.
The project is a residential develop and consists of:
-      ??? levels of Basement car park.
-      Ground level with ??? business premises.
-      Level ???-??? with Residential Apartments with approximately ??? Apartments.
            """
        )


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
        drawing_number_frame.grid(row=4, column=0)

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
        finish_frame.grid(row=5, column=0, sticky="e")

        tk.Button(finish_frame, text="Delete Project", command=self._delete_project,
                  bg="brown", fg="white", font=self.conf["font"]).pack(side=tk.RIGHT)

        self.quote_text = tk.StringVar(value="Quote Unsuccessful")
        tk.Button(finish_frame, textvariable=self.quote_text, command=self.quote_unsuccessful,
                  bg="brown", fg="white", font=self.conf["font"]).pack(side=tk.RIGHT)

        self.data["State"]["Quote Unsuccessful"].trace("w", self._update_quote_button_text)

        # tk.Button(finish_frame, text="Finish Set Up", command=self._finish_setup,
        #           bg="brown", fg="white", font=self.conf["font"]).pack(side=tk.RIGHT)

    def _update_quote_button_text(self, *args):
        if self.data["State"]["Quote Unsuccessful"].get():
            self.quote_text.set("Restore")
        else:
            self.quote_text.set("Quote Unsuccessful")

    def _delete_project(self):
        if len(self.data["Project Info"]["Project"]["Quotation Number"].get())==0:
            self.messagebox.show_error("You can't delete an empty project")
            return
        delete = self.messagebox.ask_yes_no("Are you sure you want to delete the Project?")
        if delete:
            delete_project(self.app)
            self.messagebox.show_info(f"Project {self.data['Project Info']['Project']['Quotation Number'].get()} deleted")

    def load_data(self):
        quotation_number = self.data["Project Info"]["Project"]["Quotation Number"].get().upper()
        # project_number = self.data["Project Info"]["Project"]["Project Number"].get()

        # project_quotation_dir = os.path.join(self.conf["database_dir"], "project_quotation_number_map.json")
        # project_quotation_json = json.load(open(project_quotation_dir))

        if len(quotation_number) == 0:
            self.messagebox.show_error("Please Enter a Quotation Number before You Load")
            # if len(project_number) == 0:
            #     self.messagebox.show_error("Please enter a Quotation Number or Project Number before you load")
            #     return
            # elif not project_number in project_quotation_json.keys():
            #     self.messagebox.show_error("Cannot found the Project Number")
            #     return
            # else:
            #     quotation_number = project_quotation_json[project_number]

        database_dir = os.path.join(self.conf["database_dir"], quotation_number)

        if not os.path.exists(database_dir):
            self.messagebox.show_error(f"The quotation {quotation_number} number doesn't exist")
        else:
            load_data(self.app)


    def quote_unsuccessful(self):
        if len(self.data["Project Info"]["Project"]["Quotation Number"].get())==0:
            self.messagebox.show_error("You need to load a project first")
            return

        if self.data["State"]["Quote Unsuccessful"].get():
            restore = self.messagebox.ask_yes_no("Are you sure you want to restore this project")
            if restore:
                self.data["State"]["Quote Unsuccessful"].set(False)
                save(self.app)
                config_state(self.app)
                self.app.log.log_restore(self.app)
                config_log(self.app)
                if len(self.data["Asana_id"].get()) != 0:
                    update_asana(self.app)
                    self.messagebox.show_info("Quote Unsuccessful and Asana Updated")
                else:
                    self.messagebox.show_info("Quote Unsuccessful")
        else:
            quote = self.messagebox.ask_yes_no("Are you sure you want to put this project into Quote Unsuccessful")
            if quote:
                self.data["State"]["Quote Unsuccessful"].set(True)
                save(self.app)
                config_state(self.app)
                self.app.log.log_unsuccessful(self.app)
                config_log(self.app)
                if len(self.data["Asana_id"].get()) != 0:
                    update_asana(self.app)
                    self.messagebox.show_info("Project Restore and Asana Updated")
                else:
                    self.messagebox.show_info("Project Restore")

    def _update_quotation_number(self, *args):
        if self.data["Project Info"]["Project"]["Quotation Number"].get() != "":
            return
        self.data["Project Info"]["Project"]["Quotation Number"].set(get_quotation_number())

    # def _search_projects(self):
    #     os.listdir(self.conf["working_dir"])

    def _update_area(self, *args):
        area_sum = 0
        for drawing in self.data["Project Info"]["Building Features"]["Minor"]["Details"]:
            if len(drawing["Area"].get()) == 0:
                continue
            try:
                area_sum += float(drawing["Area"].get())
            except ValueError:
                self.data["Project Info"]["Building Features"]["Minor"]["Total Area"].set("Error")
                return
        self.data["Project Info"]["Building Features"]["Minor"]["Total Area"].set(str(round(area_sum, 2)))

    def _update_car_total(self, *args):
        car_park = self.data["Project Info"]["Building Features"]["Major"]["Car Park"]
        car_part_row_total = [0]*6
        car_part_column_total = [0] * self.conf["n_major_building"]
        car_part_spot_total = 0
        car_part_other_total = 0
        for i in range(6):
            for j in range(self.conf["n_major_building"]):
                try:
                    car_part_row_total[i] += int(car_park[j][i+2].get())
                    car_part_column_total[j] += int(car_park[j][i+2].get())
                except ValueError:
                    continue
        for i in range(6):
            self.car_part_row_total[i].set(str(car_part_row_total[i]))
            if i%2 == 0:
                car_part_spot_total += car_part_row_total[i]
            else:
                car_part_other_total += car_part_row_total[i]
        for i in range(self.conf["n_major_building"]):
            self.car_part_column_total[i].set(str(car_part_column_total[i]))
        self.car_part_spot_total.set(str(car_part_spot_total))
        self.car_part_other_total.set(str(car_part_other_total))

    def _update_block_total(self, *args):
        block = self.data["Project Info"]["Building Features"]["Major"]["Block"]
        block_row_total = [0]*6
        block_column_total = [0] * self.conf["n_major_building"]
        block_apt_total = 0
        block_comm_total = 0
        for i in range(6):
            for j in range(self.conf["n_major_building"]):
                try:
                    block_row_total[i] += int(block[j][i+2].get())
                    block_column_total[j] += int(block[j][i+2].get())
                except ValueError:
                    continue
        for i in range(6):
            self.block_row_total[i].set(str(block_row_total[i]))
            if i%2 == 0:
                block_apt_total += block_row_total[i]
            else:
                block_comm_total += block_row_total[i]
        for i in range(self.conf["n_major_building"]):
            self.block_column_total[i].set(str(block_column_total[i]))
        self.block_apt_total.set(str(block_apt_total))
        self.block_comm_total.set(str(block_comm_total))

    def _update_building_frame(self, *args):
        type = self.data["Project Info"]["Project"]["Proposal Type"].get()
        if type == "Minor":
            self.major_frame.grid_forget()
            self.minor_frame.grid(row=0, column=0)
        else:
            self.minor_frame.grid_forget()
            self.major_frame.grid(row=0, column=0)

    def _update_service(self, var, *args):
        self.app.fee_proposal_page.update_scope(var)
        self.app.fee_proposal_page.update_fee(var)
        self.app.financial_panel_page.update_invoice(var)
        self.app.financial_panel_page.update_bill(var)
        self.app.financial_panel_page.update_profit(var)

    def _update_proposal(self, *args):
        for service in self.data["Project Info"]["Project"]["Service Type"].values():
            if service["Include"].get():
                service["Include"].set(False)
                service["Include"].set(True)

    def _on_mousewheel(self, event):
        self.main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def reset_scrollregion(self, event):
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
