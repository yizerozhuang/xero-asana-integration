import tkinter as tk



class ProjectInfoPage(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        # User_information
        user_info_frame = tk.LabelFrame(self, text="User Information")
        user_info_frame.grid(row=1, column=0, padx=20)

        user_info = ["Project Name", "Quotation Number", "Project Number"]
        user_info_dict = dict()

        for i, info in enumerate(user_info):
            user_info_dict[info + " Label"] = tk.Label(user_info_frame, text=info)
            user_info_dict[info + " Label"].grid(row=i, column=0)
            user_info_dict[info + " Entry"] = tk.Entry(user_info_frame, width=100)
            user_info_dict[info + " Entry"].grid(row=i, column=1, columnspan=8)

        project_type_label = tk.Label(user_info_frame, text="Project Type(Single)")
        project_type_label.grid(row=3, column=0)
        project_types = ["Restaurant", "Office", "Commercial", "Group house", "Apartment", "Mixed-use complex",
                         "School", "Others"]
        project_type_dic = dict()
        project_type = tk.StringVar()

        for i, p_type in enumerate(project_types):
            project_type_dic["type" + str(i)] = tk.Radiobutton(user_info_frame, text=p_type, variable=project_type,
                                                               value=p_type)
            project_type_dic["type" + str(i)].grid(row=3, column=i + 1)

        project_type_dic["None"] = tk.Radiobutton(user_info_frame, text="None", variable=project_type, value="None")
        project_type.set("None")

        service_type = ["Hydraulic Service", "Electrical Service", "Fire Service", "Kitchen Ventilation", "Mechanical Service", "CFD Service", "Mech Review",
                        "Installation", "Miscellaneous"]

        service_label = tk.Label(user_info_frame, text="Service Type(Multiple)")
        service_label.grid(row=4, column=0)

        services = dict()
        service_list = []
        for i, service in enumerate(service_type):
            service_list.append(tk.BooleanVar(name=service))
            services[service] = tk.Checkbutton(user_info_frame, text=service, variable=service_list[i])
            services[service].grid(row=4, column=i + 1)

        #error for loop does not work here for some reason

        service_list[0].trace("w", lambda a, b, c: self.parent.master.update_scope(service_list[0]))
        service_list[1].trace("w", lambda a, b, c: self.parent.master.update_scope(service_list[1]))
        service_list[2].trace("w", lambda a, b, c: self.parent.master.update_scope(service_list[2]))
        # service_list[3].trace("w", lambda a, b, c: self.parent.master.update_scope(service_list[3]))
        service_list[4].trace("w", lambda a, b, c: self.parent.master.update_scope(service_list[4]))
        # service_list[5].trace("w", lambda a, b, c: self.parent.master.update_scope(service_list[5]))
        # service_list[6].trace("w", lambda a, b, c: self.parent.master.update_scope(service_list[6]))
        # service_list[7].trace("w", lambda a, b, c: self.parent.master.update_scope(service_list[6]))
        # service_list[8].trace("w", lambda a, b, c: self.parent.master.update_scope(service_list[6]))

        # for i in range(6):
        #     service_list[i].trace("w", lambda a, b, c: self.parent.master.update_scope(service_list[i]))
        # i = 0
        # while i <len(service_list):
        #     service_list[i].trace("w", lambda a, b, c: self.parent.master.update_scope(service_list[i]))
        # for s in service_list:
        #     s.trace("w", lambda a, b, c:self.parent.master.update_scope(s))
        # [s.trace("w", lambda a, b, c:self.parent.master.update_scope(s)) for s in service_list]


        client_frame = tk.LabelFrame(self, text="Client")
        client_frame.grid(row=2, column=0, padx=20)

        client_info = ["Client Full Name", "Client Company", "Client Address", "Client ABN", "Contact Number",
                       "Contact Email"]
        for i, info in enumerate(client_info):
            user_info_dict[info + " Label"] = tk.Label(client_frame, text=info)
            user_info_dict[info + " Label"].grid(row=i, column=0)
            user_info_dict[info + " Entry"] = tk.Entry(client_frame, width=100)
            user_info_dict[info + " Entry"].grid(row=i, column=1)

        contact_frame = tk.LabelFrame(self, text="Main Contact")
        contact_frame.grid(row=3, column=0, padx=20)

        contact_info = ["Main Contact Full Name", "Main Contact Company", "Main Contact Address", "Main Contact ABN",
                        "Main Contact Number", "Main Contact Email"]
        for i, info in enumerate(contact_info):
            user_info_dict[info + " Label"] = tk.Label(contact_frame, text=info)
            user_info_dict[info + " Label"].grid(row=i, column=0)
            user_info_dict[info + " Entry"] = tk.Entry(contact_frame, width=100)
            user_info_dict[info + " Entry"].grid(row=i, column=1)

        build_feature_frame = tk.LabelFrame(self, text="Building Features")
        build_feature_frame.grid(row=4, column=0, padx=20)

        levels_label = tk.Label(build_feature_frame, text="Levels")
        levels_label.grid(row=0, column=0)
        space_label = tk.Label(build_feature_frame, text="Space/room Description")
        space_label.grid(row=0, column=1)
        area_label = tk.Label(build_feature_frame, text="Area/m^2")
        area_label.grid(row=0, column=2)

        n_building = 5
        building_dic = dict()
        area_list = []
        for i in range(n_building):
            building_dic["Level " + str(i + 1)] = tk.Entry(build_feature_frame, width=30)
            building_dic["Level " + str(i + 1)].grid(row=i + 1, column=0)
            building_dic["Space " + str(i + 1)] = tk.Entry(build_feature_frame, width=50)
            building_dic["Space " + str(i + 1)].grid(row=i + 1, column=1)
            area_list.append(tk.StringVar())
            building_dic["Area " + str(i + 1)] = tk.Entry(build_feature_frame, textvariable=area_list[i], width=20)
            building_dic["Area " + str(i + 1)].grid(row=i + 1, column=2)

        total_label = tk.Label(build_feature_frame, text="Total")
        total_label.grid(row=n_building + 1, column=0)
        total_area_label = tk.Label(build_feature_frame, text="0")
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
                    break
            total_area_label.config(text=str(sum))

        [s.trace("w", area_update) for s in area_list]

        drawing_number_frame = tk.LabelFrame(self, text="Drawing Number")
        drawing_number_frame.grid(row=5, column=0, padx=20)

        levels_label = tk.Label(drawing_number_frame, text="Drawing Number")
        levels_label.grid(row=0, column=0)
        space_label = tk.Label(drawing_number_frame, text="Drawing Name")
        space_label.grid(row=0, column=1)
        area_label = tk.Label(drawing_number_frame, text="Revision")
        area_label.grid(row=0, column=2)

        n_drawing = 5
        drawing_dic = dict()
        for i in range(n_drawing):
            drawing_dic["Number " + str(i + 1)] = tk.Entry(drawing_number_frame, width=30)
            drawing_dic["Number " + str(i + 1)].grid(row=i + 1, column=0)
            drawing_dic["Name " + str(i + 1)] = tk.Entry(drawing_number_frame, width=50)
            drawing_dic["Name " + str(i + 1)].grid(row=i + 1, column=1)
            drawing_dic["Revision " + str(i + 1)] = tk.Entry(drawing_number_frame, width=20)
            drawing_dic["Revision " + str(i + 1)].grid(row=i + 1, column=2)
