import tkinter as tk
import psycopg2


class FeeProposalPage(tk.Frame):
    def __init__(self, parent, app):
        tk.Frame.__init__(self, parent)
        self.parent=parent
        self.app=app
        user_info_frame = tk.LabelFrame(self, text="User Information", font=self.app.font)
        user_info_frame.grid(row=0, column=0, padx=20)

        user_info = ["Reference", "Date", "Revision"]
        user_info_dict = dict()

        for i, info in enumerate(user_info):
            user_info_dict[info + " Label"] = tk.Label(user_info_frame, text=info, font=self.app.font)
            user_info_dict[info + " Label"].grid(row=i, column=0)
            user_info_dict[info + " Entry"] = tk.Entry(user_info_frame, width=100, font=self.app.font, fg="blue")
            user_info_dict[info + " Entry"].grid(row=i, column=1, columnspan=7)

        time_frame = tk.LabelFrame(self, text="Time Consuming", font=self.app.font)
        time_frame.grid(row=1, column=0, padx=20)

        stages = ["Fee proposal stage", "Pre-design information collection", "Documentation and issue"]
        stages_dict = dict()
        working_day_list = {}
        for i, stage in enumerate(stages):
            stages_dict[stage + " Label"] = tk.Label(time_frame, text=stage, font=self.app.font)
            stages_dict[stage + " Label"].grid(row=i, column=0)
            working_day_list[stage+" Duration"] = (tk.StringVar(value="1"), tk.StringVar(value="2"))
            stages_dict[stage + " Start Day"] = tk.Entry(time_frame, font=self.app.font, fg="blue", textvariable=working_day_list[stage+" Duration"][0])
            stages_dict[stage + " Start Day"].grid(row=i, column=1)
            tk.Label(time_frame, text="-").grid(row=i, column=2)
            stages_dict[stage + " End Day"] = tk.Entry(time_frame, font=self.app.font, fg="blue", textvariable=working_day_list[stage+" Duration"][1])
            stages_dict[stage + " End Day"].grid(row=i, column=3)
            tk.Label(time_frame, text=" business day", font=self.app.font).grid(row=i, column=4)

        self.scope_frame = tk.LabelFrame(self, text="Scope of Work", font=self.app.font)
        self.scope_frame.grid(row=2, column=0, padx=20)
        self.current_frame_number = 0
        self.scope_frames = {}
        conn = psycopg2.connect(
            host="127.0.0.1",
            database="Service_Scope",
            user="postgres",
            password="Zero0929"
        )
        self.cur = conn.cursor()
        self.convert_map = {
            "Hydraulic Service": "Hydraulic Services",
            "Mechanical Service": "Mechanical Service",
            "Electrical Service": "Electrical Services",
            "Fire Service": "Fire Protection Services"
        }
        self.fee_frame = tk.LabelFrame(self, text="Fee Proposal Details", font=self.app.font)
        self.fee_frame.grid(row=3, column=0)
        service_label = tk.Label(self.fee_frame, text="Services", font=self.app.font)
        service_label.grid(row=0, column=0)
        self.total_ex_gst_label = tk.Label(self.fee_frame, text="Total ex.GST", font=self.app.font)
        self.total_ex_gst_label.grid(row=0, column=1)
        self.total_in_gst_label = tk.Label(self.fee_frame, text="Total in.GST", font=self.app.font)
        self.total_in_gst_label.grid(row=0, column=2)
        self.fee_dic = dict()
        self.total_label = tk.Label(self.fee_frame, text="Total", font=self.app.font)
        self.total_label.grid(row=1, column=0)
        self.total_ex_gst_label = tk.Label(self.fee_frame, text="0", font=self.app.font)
        self.total_ex_gst_label.grid(row=1, column=1)
        self.total_in_gst_label = tk.Label(self.fee_frame, text="0", font=self.app.font)
        self.total_in_gst_label.grid(row=1, column=2)

    def update_scope_frame(self, var):
        if var.get() == True:
            if self.scope_frames.get(var._name) == None:
                #scope part
                self.scope_frames[var._name] = tk.LabelFrame(self.scope_frame, text=var._name, font=self.app.font)
                extra_list = ["Extend", "Exclusion", "Deliverables"]
                extra_frame_list = []
                for i, extra in enumerate(extra_list):
                    extra_frame_list.append(tk.LabelFrame(self.scope_frames[var._name], text=extra, font=self.app.font))
                    extra_frame_list[i].grid(row=i, column=0)
                    self.cur.execute(
                        f"""
                            SELECT *
                            FROM service_scope
                            WHERE service_type='{self.convert_map[var._name]}'
                            AND extra ='{extra}'
                        """
                    )
                    service = self.cur.fetchall()
                    context_button_list = []
                    context_entry_list = []
                    str_var_list = []
                    bool_var_list = []
                    color_list = ["white", "powder blue"]
                    for j, context in enumerate(service):
                        str_var_list.append(tk.StringVar())
                        context_entry_list.append(tk.Entry(extra_frame_list[i], width=50, textvariable=str_var_list[j], font=self.app.font, bg=color_list[j%2]))
                        context_entry_list[j].insert(0, context[2])
                        context_entry_list[j].grid(row=j, column=1)
                        #default value need to set True
                        bool_var_list.append(tk.BooleanVar())
                        context_button_list.append(tk.Checkbutton(extra_frame_list[i], variable=bool_var_list[j], onvalue=True, offvalue=False))
                        bool_var_list[j].set(True)
                        context_button_list[j].grid(row=j, column=0)
                # fee part
                self.fee_dic["Service " + var._name] = tk.Label(self.fee_frame, text=var._name+" design and documentation", width=60, font=self.app.font)
                self.fee_dic["Service " + var._name].grid(row=self.current_frame_number+1, column=0)
                self.fee_dic["Fee " + var._name] = tk.StringVar()
                self.fee_dic["Total ex.GST " + var._name] = tk.Entry(self.fee_frame, textvariable=self.fee_dic["Fee " + var._name], font=self.app.font, fg="blue")
                self.fee_dic["Total ex.GST " + var._name].grid(row=self.current_frame_number+1, column=1)
                self.fee_dic["Total in.GST " + var._name] = tk.Label(self.fee_frame, text=0, font=self.app.font)
                self.fee_dic["Total in.GST " + var._name].grid(row=self.current_frame_number+1, column=2)

                def in_gst_update(*args):
                    num = 0
                    try:
                        num+=int(float(self.fee_dic["Fee " + var._name].get() if self.fee_dic["Fee " + var._name].get()!="" else 0)*1.1)
                    except:
                        num = "Error"
                        self.fee_dic["Total in.GST " + var._name].config(text=str(num), bg="red")
                        return
                    self.fee_dic["Total in.GST " + var._name].config(text=str(num), bg=self.cget("bg"))

                def sum_update(*args):
                    sum = 0
                    in_sum = 0
                    for key, value in self.fee_dic.items():
                        if key.startswith("Fee"):
                            try:
                                sum += float(self.fee_dic[key].get() if self.fee_dic[key].get()!="" else 0)
                                in_sum += int(float(self.fee_dic[key].get() if self.fee_dic[key].get() != "" else 0) * 1.1)
                            except:
                                sum = "Error"
                                in_sum = "Error"
                                self.total_ex_gst_label.config(text=str(sum), bg="red")
                                self.total_in_gst_label.config(text=str(in_sum), bg="red")
                                return
                        self.total_ex_gst_label.config(text=str(sum), bg=self.cget("bg"))
                        self.total_in_gst_label.config(text=str(in_sum), bg=self.cget("bg"))

                self.fee_dic["Fee " + var._name].trace("w", in_gst_update)
                self.fee_dic["Fee " + var._name].trace("w", sum_update)
            # def in_gst_update(*args):
            #     text = ""
            #     try:
            #         text = float(self.fee_dic[var._name+" Fee"])
            #     except:
            #         text = "Error"
            #     self.fee_dic["Total in.GST " + var._name].config(text=str(text))
            #
            # self.fee_dic[var._name + " Fee"].trace("w", in_gst_update)

            self.scope_frames[var._name].grid(row=0, column=self.current_frame_number)
            self.fee_dic["Service " + var._name].grid(row=self.current_frame_number+1, column=0)
            self.fee_dic["Total ex.GST " + var._name].grid(row=self.current_frame_number+1, column=1)
            self.fee_dic["Total in.GST " + var._name].grid(row=self.current_frame_number+1, column=2)
            self.current_frame_number += 1
            self.total_label.grid(row=self.current_frame_number+1, column=0)
            self.total_ex_gst_label.grid(row=self.current_frame_number+1, column=1)
            self.total_in_gst_label.grid(row=self.current_frame_number+1, column=2)
        else:
            self.scope_frames[var._name].grid_forget()
            self.fee_dic["Service " + var._name].grid_forget()
            self.fee_dic["Fee "+ var._name].set("")
            self.fee_dic["Total ex.GST " + var._name].grid_forget()
            self.fee_dic["Total in.GST " + var._name].grid_forget()