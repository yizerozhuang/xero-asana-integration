import tkinter as tk
from tkinter import ttk


class ProjectInfoPage(tk.Frame):
    def __init__(self, parent, app):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.app = app
        self.main_frame = tk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=1)

        self.main_canvas = tk.Canvas(self.main_frame)
        self.main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

        self.scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.main_canvas.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y, expand=1)

        self.main_canvas.config(yscrollcommand=self.scrollbar.set)
        self.main_canvas.bind("<Configure>", lambda e:self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")))
        self.main_canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.main_context_frame = tk.LabelFrame(self.main_canvas, text="Main Context")
        self.main_canvas.create_window((0, 0), window=self.main_context_frame, anchor="center")

        self.project_info_frame()
        self.client_frame()
        self.main_contact_frame()
        self.building_features_frame()
        self.drawing_number_frame()

    def project_info_frame(self):
        # User_information frame



        user_info_frame = tk.LabelFrame(self.main_context_frame, text="Project Information", font=self.app.font)
        user_info_frame.grid(row=1, column=0, padx=20)

        user_info = ["Project Name", "Quotation Number", "Project Number"]

        self.app.data["Project Information"] = dict()

        for i, info in enumerate(user_info):
            tk.Label(user_info_frame,  width=30, text=info, font=self.app.font).grid(row=i, column=0, padx=(10, 0))
            self.app.data["Project Information"][info] = tk.StringVar(name=info)
            tk.Entry(user_info_frame, width=70, font=self.app.font, fg="blue", textvariable=self.app.data["Project Information"][info]).grid(row=i, column=1, columnspan=3, padx=(0, 10))

        tk.Label(user_info_frame, text="Project Type", font=self.app.font).grid(row=3, column=0)
        project_types = ["Restaurant", "Office", "Commercial", "Group house", "Apartment", "Mixed-use complex",
                         "School", "Others"]

        self.app.data["Project Information"]["Project Type"] = tk.StringVar(name="Project Type", value="Restaurant")

        for i, p_type in enumerate(project_types):
            button = tk.Radiobutton(user_info_frame, text=p_type, variable=self.app.data["Project Information"]["Project Type"],value=p_type, font=self.app.font)
            if i < 3:
                button.grid(row=3, column=i + 1, sticky="W")
            elif i < 6:
                button.grid(row=4, column=i - 2, sticky="W")
            else:
                button.grid(row=5, column=i - 5, sticky="W")

        service_type = ["Hydraulic Service", "Electrical Service", "Fire Service", "Kitchen Ventilation", "Mechanical Service", "CFD Service", "Mech Review",
                        "Installation", "Miscellaneous"]

        tk.Label(user_info_frame, width=30, text="Service Type", font=self.app.font).grid(row=6, column=0)

        self.app.data["Project Information"]["Service Type"] = []
        for i, service in enumerate(service_type):
            self.app.data["Project Information"]["Service Type"].append(tk.BooleanVar(name=service))
            button = tk.Checkbutton(user_info_frame, text=service, variable=self.app.data["Project Information"]["Service Type"][i], font=self.app.font)
            if i < 3:
                button.grid(row=6, column=i + 1, sticky="W")
            elif i < 6:
                button.grid(row=7, column=i - 2, sticky="W")
            else:
                button.grid(row=8, column=i - 5, sticky="W")

        self.app.data["Project Information"]["Service Type"][0].trace("w", lambda a, b, c: self.app.update_service_type(
            self.app.data["Project Information"]["Service Type"][0]))
        self.app.data["Project Information"]["Service Type"][1].trace("w", lambda a, b, c: self.app.update_service_type(
            self.app.data["Project Information"]["Service Type"][1]))
        self.app.data["Project Information"]["Service Type"][2].trace("w", lambda a, b, c: self.app.update_service_type(
            self.app.data["Project Information"]["Service Type"][2]))
        self.app.data["Project Information"]["Service Type"][4].trace("w", lambda a, b, c: self.app.update_service_type(
            self.app.data["Project Information"]["Service Type"][4]))

        # #error for loop does not work here for some reason
        #
        # service_list[0].trace("w", lambda a, b, c: self.app.update_scope(service_list[0]))
        # service_list[1].trace("w", lambda a, b, c: self.app.update_scope(service_list[1]))
        # service_list[2].trace("w", lambda a, b, c: self.app.update_scope(service_list[2]))
        # # service_list[3].trace("w", lambda a, b, c:self.app.update_scope(service_list[3]))
        # service_list[4].trace("w", lambda a, b, c: self.app.update_scope(service_list[4]))
        # # service_list[5].trace("w", lambda a, b, c: self.app.master.update_scope(service_list[5]))
        # # service_list[6].trace("w", lambda a, b, c: self.app.master.update_scope(service_list[6]))
        # # service_list[7].trace("w", lambda a, b, c: self.app.master.update_scope(service_list[6]))
        # # service_list[8].trace("w", lambda a, b, c: self.app.master.update_scope(service_list[6]))
        #
        # # for i in range(6):
        # #     service_list[i].trace("w", lambda a, b, c: self.parent.master.update_scope(service_list[i]))
        # # i = 0
        # # while i <len(service_list):
        # #     service_list[i].trace("w", lambda a, b, c: self.parent.master.update_scope(service_list[i]))
        # # for s in service_list:
        # #     s.trace("w", lambda a, b, c:self.parent.master.update_scope(s))
        # # [s.trace("w", lambda a, b, c:self.parent.master.update_scope(s)) for s in service_list]


    def client_frame(self):

        #client frame
        client_frame = tk.LabelFrame(self.main_context_frame, text="Client", font=self.app.font)
        client_frame.grid(row=2, column=0)
        self.app.data["Client"] = dict()

        tk.Label(client_frame, width=30, text="Client Full Name", font=self.app.font).grid(row=0, column=0, padx=(10, 0))
        self.app.data["Client"]["Client Full Name"] = tk.StringVar(name="Client Full Name")
        tk.Entry(client_frame, width=70, font=self.app.font, fg="blue", textvariable=self.app.data["Client"]["Client Full Name"]).grid(row=0, column=1, columnspan=3, padx=(0,10))
        self.app.data["Client"]["Client Full Name"].trace("w", self.app.update_client)

        tk.Label(client_frame, width=30, text="Client Contact Type", font=self.app.font).grid(row=1, column=0)

        contact_type = ["Architect", "Builder", "Certifier", "Contractor", "Developer", "Engineer", "Government", "RDM", "Strata", "Supplier", "Owner/Tenant", "Others"]

        self.app.data["Client"]["Client Contact Type"] = tk.StringVar(value="Architect", name="Client Contact Type")
        for i, p_type in enumerate(contact_type):
            button = tk.Radiobutton(client_frame, text=p_type, variable=self.app.data["Client"]["Client Contact Type"], value=p_type, font=self.app.font)
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
            tk.Label(client_frame, width=30, text=info, font=self.app.font).grid(row=i+5, column=0)
            self.app.data["Client"][info] = tk.StringVar(name=info)
            tk.Entry(client_frame, width=70, font=self.app.font, fg="blue", textvariable=self.app.data["Client"][info]).grid(row=i+5, column=1, columnspan=3)

    
    def main_contact_frame(self):
        contact_frame = tk.LabelFrame(self.main_context_frame, text="Main Contact", font=self.app.font)
        contact_frame.grid(row=3, column=0, padx=20)
        
        self.app.data["Main Contact"] = dict()
        tk.Label(contact_frame, width=30, text="Main Contact Full Name", font=self.app.font).grid(row=0, column=0, padx=(10, 0))

        self.app.data["Main Contact"]["Main Contact Full Name"]= tk.StringVar(name="Main Contact Full Name")
        tk.Entry(contact_frame, width=70, font=self.app.font, fg="blue", textvariable=self.app.data["Main Contact"]["Main Contact Full Name"]).grid(row=0, column=1, columnspan=3, padx=(0, 10))

        tk.Label(contact_frame, text="Main Contact Contact Type", font=self.app.font).grid(row=1, column=0)

        contact_type = ["Architect", "Builder", "Certifier", "Contractor", "Developer", "Engineer", "Government", "RDM",
                        "Strata", "Supplier", "Owner/Tenant", "Others"]

        self.app.data["Main Contact"]["Main Contact Contact Type"] = tk.StringVar(name="Main Contact Contact Type", value="Architect")
        for i, p_type in enumerate(contact_type):
            button = tk.Radiobutton(contact_frame, text=p_type, variable=self.app.data["Main Contact"]["Main Contact Contact Type"],
                                                               value=p_type, font=self.app.font)
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
            tk.Label(contact_frame,width=30, text=info, font=self.app.font).grid(row=i+5, column=0)
            self.app.data["Main Contact"][info] = tk.StringVar(name=info)
            tk.Entry(contact_frame, width=70, font=self.app.font, fg="blue", textvariable=self.app.data["Main Contact"][info]).grid(row=i+5, column=1, columnspan=3)


    def building_features_frame(self):

        build_feature_frame = tk.LabelFrame(self.main_context_frame, text="Building Features", font=self.app.font)
        build_feature_frame.grid(row=4, column=0)

        tk.Label(build_feature_frame, text="Levels", font=self.app.font).grid(row=0, column=0)
        tk.Label(build_feature_frame, text="Space/room Description", font=self.app.font).grid(row=0, column=1)
        tk.Label(build_feature_frame, text="Area/m^2", font=self.app.font).grid(row=0, column=2)

        n_building = 5
        self.app.data["Building Feature"] = {
            "Levels": [],
            "Space/room Description": [],
            "Area/m^2": [],
            "Total": None
        }

        def area_update(*args):
            sum = 0
            for area in self.app.data["Building Feature"]["Area/m^2"]:
                if area.get() == "":
                    continue
                try:
                    sum += float(area.get())
                except:
                    sum = "Error"
                    total_area_label.config(text=str(sum), bg="red")
                    self.app.data["Building Feature"]["Total"] = sum
                    return
            total_area_label.config(text=str(sum), bg=self.cget("bg"))
            self.app.data["Building Feature"]["Total"] = sum

        tk.Label(build_feature_frame, text="Total", font=self.app.font).grid(row=n_building + 1, column=0)
        total_area_label = tk.Label(build_feature_frame, text="0", font=self.app.font)
        total_area_label.grid(row=n_building + 1, column=2)

        for i in range(n_building):
            self.app.data["Building Feature"]["Levels"].append(tk.StringVar(name=f"Levels {str(i)}"))
            self.app.data["Building Feature"]["Space/room Description"].append(tk.StringVar(name=f"Space/room Description {str(i)} "))
            self.app.data["Building Feature"]["Area/m^2"].append(tk.StringVar(name=f"Area/m^2 {str(i)} "))

            tk.Entry(build_feature_frame, width=34, font=self.app.font, fg="blue", textvariable=self.app.data["Building Feature"]["Levels"][i]).grid(row=i + 1, column=0, padx=(10, 0))
            tk.Entry(build_feature_frame, width=50, font=self.app.font, fg="blue", textvariable=self.app.data["Building Feature"]["Space/room Description"][i]).grid(row=i + 1, column=1)
            tk.Entry(build_feature_frame, width=20, font=self.app.font, fg="blue", textvariable=self.app.data["Building Feature"]["Area/m^2"][i]).grid(row=i + 1, column=2, padx=(0, 10))

            self.app.data["Building Feature"]["Area/m^2"][i].trace("w", area_update)


    def drawing_number_frame(self):
        drawing_number_frame = tk.LabelFrame(self.main_context_frame, text="Drawing Number", font=self.app.font)
        drawing_number_frame.grid(row=5, column=0, padx=20)

        tk.Label(drawing_number_frame, text="Drawing Number", font=self.app.font).grid(row=0, column=0)
        tk.Label(drawing_number_frame, text="Drawing Name", font=self.app.font).grid(row=0, column=1)
        tk.Label(drawing_number_frame, text="Revision", font=self.app.font).grid(row=0, column=2)

        n_drawing = 5
        drawing_dic = dict()

        self.app.data["Building Feature"] = {
            "Drawing Number": [],
            "Drawing Name": [],
            "Revision": []
        }

        for i in range(n_drawing):
            self.app.data["Building Feature"]["Drawing Number"].append(tk.StringVar(name=f"Drawing Number {str(i)}"))
            self.app.data["Building Feature"]["Drawing Name"].append(tk.StringVar(name=f"Drawing Name {str(i)}"))
            self.app.data["Building Feature"]["Revision"].append(tk.StringVar(name=f"Drawing Revision {str(i)}"))


            tk.Entry(drawing_number_frame, width=34, font=self.app.font, fg="blue", textvariable=self.app.data["Building Feature"]["Drawing Number"][i]).grid(row=i + 1, column=0, padx=(10, 0))
            tk.Entry(drawing_number_frame, width=50, font=self.app.font, fg="blue", textvariable=self.app.data["Building Feature"]["Drawing Name"][i]).grid(row=i + 1, column=1)
            tk.Entry(drawing_number_frame, width=20, font=self.app.font, fg="blue", textvariable=self.app.data["Building Feature"]["Revision"][i]).grid(row=i + 1, column=2, padx=(0, 10))

        self.app.data["Building Feature"]["Drawing Number"][0].set("M-000")
        self.app.data["Building Feature"]["Drawing Name"][0].set("Cover Sheet")
        self.app.data["Building Feature"]["Revision"][0].set("A")
        self.app.data["Building Feature"]["Drawing Number"][1].set("M-001")
        self.app.data["Building Feature"]["Drawing Name"][1].set("Tenancy Level Layout")
        self.app.data["Building Feature"]["Revision"][1].set("A")
        self.app.data["Building Feature"]["Drawing Number"][2].set("M-002")
        self.app.data["Building Feature"]["Drawing Name"][2].set("Roof Layout")
        self.app.data["Building Feature"]["Revision"][2].set("A")

    def _on_mousewheel(self, event):
        self.main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

