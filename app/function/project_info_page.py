import tkinter as tk



class ProjectInfoPage(tk.Frame):
    def __init__(self, parent, app):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.app = app
        self.user_information_fame()
        self.client_frame()
        self.main_contact_frame()
        self.building_features_frame()
        self.drawing_number_frame()
    def user_information_fame(self):
        # User_information frame
        user_info_frame = tk.LabelFrame(self, text="User Information", font=self.app.font)
        user_info_frame.grid(row=1, column=0, padx=20)

        user_info = ["Project Name", "Quotation Number", "Project Number"]

        self.app.data["User Information"] = dict()

        for i, info in enumerate(user_info):
            tk.Label(user_info_frame,  width=30, text=info, font=self.app.font).grid(row=i, column=0, padx=(10, 0))
            self.app.data["User Information"][info] = tk.StringVar(name=info)
            tk.Entry(user_info_frame, width=70, font=self.app.font, fg="blue", textvariable=self.app.data["User Information"][info]).grid(row=i, column=1, columnspan=3, padx=(0, 10))

        tk.Label(user_info_frame, text="Project Type(Single)", font=self.app.font).grid(row=3, column=0)
        project_types = ["Restaurant", "Office", "Commercial", "Group house", "Apartment", "Mixed-use complex",
                         "School", "Others"]

        self.app.data["User Information"]["Project Type"] = tk.StringVar(name="Project Type", value="None")

        for i, p_type in enumerate(project_types):
            button = tk.Radiobutton(user_info_frame, text=p_type, variable=self.app.data["User Information"]["Project Type"],value=p_type, font=self.app.font)
            if i < 3:
                button.grid(row=3, column=i + 1)
            elif i < 6:
                button.grid(row=4, column=i - 2)
            else:
                button.grid(row=5, column=i - 5)

        service_type = ["Hydraulic Service", "Electrical Service", "Fire Service", "Kitchen Ventilation", "Mechanical Service", "CFD Service", "Mech Review",
                        "Installation", "Miscellaneous"]

        tk.Label(user_info_frame, width=30, text="Service Type(Multiple)", font=self.app.font).grid(row=6, column=0)

        self.app.data["User Information"]["Service Type"] = []
        for i, service in enumerate(service_type):
            self.app.data["User Information"]["Service Type"].append(tk.BooleanVar(name=service))
            button = tk.Checkbutton(user_info_frame, text=service, variable=self.app.data["User Information"]["Service Type"][i], font=self.app.font)
            if i < 3:
                button.grid(row=6, column=i + 1)
            elif i < 6:
                button.grid(row=7, column=i - 2)
            else:
                button.grid(row=8, column=i - 5)

        self.app.data["User Information"]["Service Type"][0].trace("w", lambda a, b, c: self.app.update_scope(
            self.app.data["User Information"]["Service Type"][0]))
        self.app.data["User Information"]["Service Type"][1].trace("w", lambda a, b, c: self.app.update_scope(
            self.app.data["User Information"]["Service Type"][1]))
        self.app.data["User Information"]["Service Type"][2].trace("w", lambda a, b, c: self.app.update_scope(
            self.app.data["User Information"]["Service Type"][2]))
        self.app.data["User Information"]["Service Type"][4].trace("w", lambda a, b, c: self.app.update_scope(
            self.app.data["User Information"]["Service Type"][4]))
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
        client_dic = dict()
        #client frame
        client_frame = tk.LabelFrame(self, text="Client", font=self.app.font)
        client_frame.grid(row=2, column=0)
        client_dic["Client Full Name Label"] = tk.Label(client_frame, width=30, text="Client Full Name", font=self.app.font)
        client_dic["Client Full Name Label"].grid(row=0, column=0, padx=(10, 0))
        client_dic["Client Full Name Entry"] = tk.Entry(client_frame, width=70, font=self.app.font, fg="blue")
        client_dic["Client Full Name Entry"].grid(row=0, column=1, columnspan=3, padx=(0,10))

        client_dic["Client Full Contact Type"] = tk.Label(client_frame, width=30, text="Client Contact Type",
                                                            font=self.app.font)
        client_dic["Client Full Contact Type"].grid(row=1, column=0)

        contact_type = ["Architect", "Builder", "Certifier", "Contractor", "Developer", "Engineer", "Government", "RDM", "Strata", "Supplier", "Owner/Tenant", "Others"]
        client_contact_type_dic = dict()
        client_type = tk.StringVar(value="None")
        for i, p_type in enumerate(contact_type):
            client_contact_type_dic["type" + str(i)] = tk.Radiobutton(client_frame, text=p_type, variable=client_type,
                                                               value=p_type, font=self.app.font)
            if i < 3:
                client_contact_type_dic["type" + str(i)].grid(row=1, column=i + 1)
            elif i < 6:
                client_contact_type_dic["type" + str(i)].grid(row=2, column=i - 2)
            elif i < 9:
                client_contact_type_dic["type" + str(i)].grid(row=3, column=i - 5)
            else:
                client_contact_type_dic["type" + str(i)].grid(row=4, column=i - 8)
        client_info = ["Client Company", "Client Address", "Client ABN", "Contact Number",
                       "Contact Email"]

        for i, info in enumerate(client_info):
            client_dic[info + " Label"] = tk.Label(client_frame, width=30, text=info, font=self.app.font)
            client_dic[info + " Label"].grid(row=i+5, column=0)
            client_dic[info + " Entry"] = tk.Entry(client_frame, width=70, font=self.app.font, fg="blue")
            client_dic[info + " Entry"].grid(row=i+5, column=1, columnspan=3)

    
    def main_contact_frame(self):
        contact_frame = tk.LabelFrame(self, text="Main Contact", font=self.app.font)
        contact_frame.grid(row=3, column=0, padx=20)
        
        main_contact_dic = dict()
        main_contact_dic["Main Contact Full Name Label"] = tk.Label(contact_frame, width=30, text="Main Contact Full Name", font=self.app.font)
        main_contact_dic["Main Contact Full Name Label"].grid(row=0, column=0, padx=(10, 0))
        main_contact_dic["Main Contact Full Name Entry"] = tk.Entry(contact_frame, width=70, font=self.app.font, fg="blue")
        main_contact_dic["Main Contact Full Name Entry"].grid(row=0, column=1, columnspan=3, padx=(0, 10))

        main_contact_dic["Main Contact Contact Type"] = tk.Label(contact_frame, text="Main Contact Contact Type",
                                                            font=self.app.font)
        main_contact_dic["Main Contact Contact Type"].grid(row=1, column=0)

        contact_type = ["Architect", "Builder", "Certifier", "Contractor", "Developer", "Engineer", "Government", "RDM",
                        "Strata", "Supplier", "Owner/Tenant", "Others"]
        main_contact_type_dic = dict()
        main_contact_type = tk.StringVar(value="None")
        for i, p_type in enumerate(contact_type):
            main_contact_type_dic["type" + str(i)] = tk.Radiobutton(contact_frame, text=p_type, variable=main_contact_type,
                                                               value=p_type, font=self.app.font)
            if i < 3:
                main_contact_type_dic["type" + str(i)].grid(row=1, column=i + 1)
            elif i < 6:
                main_contact_type_dic["type" + str(i)].grid(row=2, column=i - 2)
            elif i < 9:
                main_contact_type_dic["type" + str(i)].grid(row=3, column=i - 5)
            else:
                main_contact_type_dic["type" + str(i)].grid(row=4, column=i - 8)

        contact_info = ["Main Contact Company", "Main Contact Address", "Main Contact ABN",
                        "Main Contact Number", "Main Contact Email"]
        for i, info in enumerate(contact_info):
            main_contact_dic[info + " Label"] = tk.Label(contact_frame,width=30, text=info, font=self.app.font)
            main_contact_dic[info + " Label"].grid(row=i+5, column=0)
            main_contact_dic[info + " Entry"] = tk.Entry(contact_frame, width=70, font=self.app.font, fg="blue")
            main_contact_dic[info + " Entry"].grid(row=i+5, column=1, columnspan=3)


    def building_features_frame(self):

        build_feature_frame = tk.LabelFrame(self, text="Building Features", font=self.app.font)
        build_feature_frame.grid(row=4, column=0)
        levels_label = tk.Label(build_feature_frame, text="Levels", font=self.app.font)
        levels_label.grid(row=0, column=0)
        space_label = tk.Label(build_feature_frame, text="Space/room Description", font=self.app.font)
        space_label.grid(row=0, column=1)
        area_label = tk.Label(build_feature_frame, text="Area/m^2", font=self.app.font)
        area_label.grid(row=0, column=2)

        n_building = 5
        building_dic = dict()
        area_list = []
        for i in range(n_building):
            building_dic["Level " + str(i + 1)] = tk.Entry(build_feature_frame, width=34, font=self.app.font, fg="blue")
            building_dic["Level " + str(i + 1)].grid(row=i + 1, column=0, padx=(10, 0))
            building_dic["Space " + str(i + 1)] = tk.Entry(build_feature_frame, width=50, font=self.app.font, fg="blue")
            building_dic["Space " + str(i + 1)].grid(row=i + 1, column=1)
            area_list.append(tk.StringVar())
            building_dic["Area " + str(i + 1)] = tk.Entry(build_feature_frame, width=20, textvariable=area_list[i], font=self.app.font, fg="blue")
            building_dic["Area " + str(i + 1)].grid(row=i + 1, column=2, padx=(0, 10))

        total_label = tk.Label(build_feature_frame, text="Total", font=self.app.font)
        total_label.grid(row=n_building + 1, column=0)
        total_area_label = tk.Label(build_feature_frame, text="0", font=self.app.font)
        total_area_label.grid(row=n_building + 1, column=2)

        def area_update(*args):
            sum = 0
            for area in area_list:
                if area.get() == "":
                    continue
                try:
                    sum += float(area.get())
                except:
                    sum = "Error"
                    total_area_label.config(text=str(sum), bg="red")
                    return
            total_area_label.config(text=str(sum), bg=self.cget("bg"))

        [s.trace("w", area_update) for s in area_list]

    def drawing_number_frame(self):
        drawing_number_frame = tk.LabelFrame(self, text="Drawing Number", font=self.app.font)
        drawing_number_frame.grid(row=5, column=0, padx=20)

        levels_label = tk.Label(drawing_number_frame, text="Drawing Number", font=self.app.font)
        levels_label.grid(row=0, column=0)
        space_label = tk.Label(drawing_number_frame, text="Drawing Name", font=self.app.font)
        space_label.grid(row=0, column=1)
        area_label = tk.Label(drawing_number_frame, text="Revision", font=self.app.font)
        area_label.grid(row=0, column=2)

        n_drawing = 5
        drawing_dic = dict()
        for i in range(n_drawing):
            drawing_dic["Number " + str(i + 1)] = tk.Entry(drawing_number_frame, width=34, font=self.app.font, fg="blue")
            drawing_dic["Number " + str(i + 1)].grid(row=i + 1, column=0, padx=(10, 0))
            drawing_dic["Name " + str(i + 1)] = tk.Entry(drawing_number_frame, width=50, font=self.app.font, fg="blue")
            drawing_dic["Name " + str(i + 1)].grid(row=i + 1, column=1)
            drawing_dic["Revision " + str(i + 1)] = tk.Entry(drawing_number_frame, width=20, font=self.app.font, fg="blue")
            drawing_dic["Revision " + str(i + 1)].grid(row=i + 1, column=2, padx=(0, 10))
