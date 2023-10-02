import tkinter as tk
from tkinter import ttk
from datetime import date
import psycopg2


class FeeProposalPage(tk.Frame):
    def __init__(self, parent, app):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.app = app
        self.main_frame = tk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=1, side=tk.LEFT)

        self.main_canvas = tk.Canvas(self.main_frame)
        self.main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

        self.scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.main_canvas.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y, expand=1)

        self.main_canvas.config(yscrollcommand=self.scrollbar.set)
        self.main_canvas.bind("<Configure>", lambda e:self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")))
        self.main_canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.main_context_frame = tk.LabelFrame(self.main_canvas, text="Main Context")
        self.main_canvas.create_window((0, 0), window=self.main_context_frame, anchor="center")
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
        self.conn = psycopg2.connect(
            host="127.0.0.1",
            database="postgres",
            user="postgres",
            password="Zero0929"
        )
        self.cur = self.conn.cursor()
        self.conn.autocommit = True
        self.convert_map = {
            "Hydraulic Service": "Hydraulic Services",
            "Mechanical Service": "Mechanical Service",
            "Electrical Service": "Electrical Services",
            "Fire Service": "Fire Protection Services"
        }
        self.app.data["Fee Proposal Page"]["Scope of Work"] = {}

        self.append_frame = tk.Frame(self.scope_frame)
        self.append_frame.grid(row=999, column=0)

        self.append_service = tk.StringVar()
        service_type = ["Hydraulic Service", "Electrical Service", "Fire Service", "Kitchen Ventilation", "Mechanical Service", "CFD Service", "Mech Review",
                        "Installation", "Miscellaneous"]
        ttk.Combobox(self.append_frame, width=20, textvariable=self.append_service, values=service_type, state="readonly").grid(row=0, column=0)

        self.append_extra = tk.StringVar()
        extra_type = ["Extend", "Exclusion", "Deliverables"]

        ttk.Combobox(self.append_frame, width=20, textvariable=self.append_extra, values=extra_type, state="readonly").grid(row=0, column=1)

        self.append_context = tk.StringVar()
        tk.Entry(self.append_frame, width=60, textvariable=self.append_context).grid(row=0, column=2)

        tk.Button(self.append_frame, text="Submit", command=self.append_value).grid(row=1, column=0)

        self.append_database = tk.BooleanVar()
        tk.Checkbutton(self.append_frame, variable=self.append_database, text="Add to Database").grid(row=1, column=1)



    def fee_frame(self):
        self.fee_frame = tk.LabelFrame(self.main_context_frame, text="Fee Proposal Details", font=self.app.font)
        self.fee_frame.pack(fill=tk.BOTH, expand=1, padx=20)


        tk.Label(self.fee_frame, text="Services", font=self.app.font).grid(row=0, column=0)
        tk.Label(self.fee_frame, text="Total ex.GST", font=self.app.font).grid(row=0, column=1)
        tk.Label(self.fee_frame, text="Total in.GST", font=self.app.font).grid(row=0, column=2)

        self.app.data["Fee Proposal Page"]["Details"] = dict()
        self.fee_dic = dict()
        self.bool_var_list = []

        tk.Label(self.fee_frame, text="Total", font=self.app.font).grid(row=999, column=0)
        self.total_ex_gst_label = tk.Label(self.fee_frame, text="0", font=self.app.font)
        self.total_ex_gst_label.grid(row=999, column=1)
        self.total_in_gst_label = tk.Label(self.fee_frame, text="0", font=self.app.font)
        self.total_in_gst_label.grid(row=999, column=2)

    def update_scope_frame(self, var):
        if var.get() == True:
            if self.scope_frames.get(var._name) == None:
                #scope part
                self.scope_frames[var._name] = dict()
                self.scope_frames[var._name]["main"] = tk.LabelFrame(self.scope_frame, text=var._name, font=self.app.font)

                self.app.data["Fee Proposal Page"]["Scope of Work"][var._name] = dict()
                self.app.data["Fee Proposal Page"]["Scope of Work"][var._name]["on"] = True

                extra_list = ["Extend", "Exclusion", "Deliverables"]
                for i, extra in enumerate(extra_list):

                    self.app.data["Fee Proposal Page"]["Scope of Work"][var._name][extra]=[]

                    self.scope_frames[var._name][extra] = tk.LabelFrame(self.scope_frames[var._name]["main"], text=extra, font=self.app.font)
                    self.scope_frames[var._name][extra].grid(row=i, column=0)
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
                    color_list = ["white", "azure"]
                    for j, context in enumerate(service):
                        self.app.data["Fee Proposal Page"]["Scope of Work"][var._name][extra].append((tk.BooleanVar(value=True), tk.StringVar(value=context[2])))
                        context_entry_list.append(tk.Entry(self.scope_frames[var._name][extra], width=100, textvariable=self.app.data["Fee Proposal Page"]["Scope of Work"][var._name][extra][j][1], font=self.app.font, bg=color_list[j%2]))
                        context_entry_list[j].grid(row=j, column=1)
                        context_button_list.append(tk.Checkbutton(self.scope_frames[var._name][extra], variable=self.app.data["Fee Proposal Page"]["Scope of Work"][var._name][extra][j][0]))
                        context_button_list[j].grid(row=j, column=0)



                # fee part

                self.app.data["Fee Proposal Page"]["Details"][var._name] = tk.StringVar(name=var._name+" Fee")
                self.fee_dic["Service " + var._name] = tk.Label(self.fee_frame, text=var._name+" design and documentation", width=60, font=self.app.font)
                self.fee_dic["Total ex.GST " + var._name] = tk.Entry(self.fee_frame, textvariable=self.app.data["Fee Proposal Page"]["Details"][var._name], width=20, font=self.app.font, fg="blue")
                self.fee_dic["Total in.GST " + var._name] = tk.Label(self.fee_frame, text=0, width=20, font=self.app.font)

                def in_gst_update(*args):
                    num = 0
                    try:
                        num+=int(float(self.app.data["Fee Proposal Page"]["Details"][var._name].get() if self.app.data["Fee Proposal Page"]["Details"][var._name].get()!="" else 0)*1.1)
                    except:
                        num = "Error"
                        self.fee_dic["Total in.GST " + var._name].config(text=str(num), bg="red")
                        return
                    self.fee_dic["Total in.GST " + var._name].config(text=str(num), bg=self.cget("bg"))

                def sum_update(*args):
                    sum = 0
                    in_sum = 0
                    for key, value in self.app.data["Fee Proposal Page"]["Details"].items():
                        try:
                            sum += float(self.app.data["Fee Proposal Page"]["Details"][key].get() if self.app.data["Fee Proposal Page"]["Details"][key].get()!="" else 0)
                            in_sum += int(float(self.app.data["Fee Proposal Page"]["Details"][key].get() if self.app.data["Fee Proposal Page"]["Details"][key].get() != "" else 0) * 1.1)
                        except:
                            sum = "Error"
                            in_sum = "Error"
                            self.total_ex_gst_label.config(text=str(sum), bg="red")
                            self.total_in_gst_label.config(text=str(in_sum), bg="red")
                            return
                        self.total_ex_gst_label.config(text=str(sum), bg=self.cget("bg"))
                        self.total_in_gst_label.config(text=str(in_sum), bg=self.cget("bg"))

                self.app.data["Fee Proposal Page"]["Details"][var._name].trace("w", in_gst_update)
                self.app.data["Fee Proposal Page"]["Details"][var._name].trace("w", sum_update)

            self.scope_frames[var._name]["main"].grid(row=self.current_frame_number, column=0)

            self.fee_dic["Service " + var._name].grid(row=self.current_frame_number+1, column=0)
            self.fee_dic["Total ex.GST " + var._name].grid(row=self.current_frame_number+1, column=1)
            self.fee_dic["Total in.GST " + var._name].grid(row=self.current_frame_number+1, column=2)
            self.current_frame_number += 1
        else:
            self.scope_frames[var._name]["main"].grid_forget()
            self.fee_dic["Service " + var._name].grid_forget()
            self.app.data["Fee Proposal Page"]["Details"][var._name].set("")
            self.app.data["Fee Proposal Page"]["Scope of Work"][var._name]["on"] = False
            self.fee_dic["Total ex.GST " + var._name].grid_forget()
            self.fee_dic["Total in.GST " + var._name].grid_forget()


    def _on_mousewheel(self, event):
        self.main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")


    def append_value(self, *args):
        if self.append_service.get() == "":
            tk.messagebox.showwarning(title="Error", message="You need to choose a service type")
            return
        elif not self.append_service.get() in self.app.data["Fee Proposal Page"]["Scope of Work"].keys():
            tk.messagebox.showwarning(title="Error", message="You cant add a service that you didnt choose")
            return
        elif self.app.data["Fee Proposal Page"]["Scope of Work"][self.append_service.get()]["on"] == False:
            tk.messagebox.showwarning(title="Error", message="You cant add a service that you didnt choose")
            return
        elif self.append_extra.get() == "":
            tk.messagebox.showwarning(title="Error", message="You need to choose a extra type")
            return
        elif self.append_context.get() == "":
            tk.messagebox.showwarning(title="Error", message="You cant need to enter some context")
            return

        self.app.data["Fee Proposal Page"]["Scope of Work"][self.append_service.get()][self.append_extra.get()].append((tk.BooleanVar(value=True), tk.StringVar(value=self.append_context.get())))
        tk.Entry(self.scope_frames[self.append_service.get()][self.append_extra.get()], width=100,
                 textvariable=self.app.data["Fee Proposal Page"]["Scope of Work"][self.append_service.get()][self.append_extra.get()][-1][1],
                 font=self.app.font).grid(row=len(self.app.data["Fee Proposal Page"]["Scope of Work"][self.append_service.get()][self.append_extra.get()]), column=1)

        tk.Checkbutton(self.scope_frames[self.append_service.get()][self.append_extra.get()],
                       variable=self.app.data["Fee Proposal Page"]["Scope of Work"][self.append_service.get()][self.append_extra.get()][-1][0])\
            .grid(row=len(self.app.data["Fee Proposal Page"]["Scope of Work"][self.append_service.get()][self.append_extra.get()]), column=0)

        if self.append_database.get():
            self.cur.execute(f"""
                INSERT INTO service_scope(service_type, extra, context)
                VALUES 
                ('{self.convert_map[self.append_service.get()]}','{self.append_extra.get()}','{self.append_context.get()}')
                """)
            self.conn.commit()




        self.append_service.set("")
        self.append_extra.set("")
        self.append_context.set("")
        self.append_database.set(False)

