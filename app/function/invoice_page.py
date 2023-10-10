import tkinter as tk
from tkinter import ttk
class InvoicePage(tk.Frame):
    def __init__(self, parent, app):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.app = app
        self.main_frame = tk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=1, side=tk.LEFT, anchor="nw")

        self.main_canvas = tk.Canvas(self.main_frame)
        self.main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1, anchor="nw")

        self.scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.main_canvas.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y, expand=1)

        self.main_canvas.config(yscrollcommand=self.scrollbar.set)
        self.main_canvas.bind("<Configure>", lambda e:self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")))
        self.app.bind("<MouseWheel>", self._on_mousewheel, add="+")
        self.main_context_frame = tk.LabelFrame(self.main_canvas, text="Main Context")
        self.main_canvas.create_window((0, 0), window=self.main_context_frame, anchor="nw")
        self.main_context_frame.bind("<Configure>", self.reset_scrollregion)
        self.items_frame = {}
        self.invoice_list = {}
        self.fee_frame()

    def fee_frame(self):
        self.fee_frame = tk.LabelFrame(self.main_context_frame, text="Invoice Details", font=self.app.font)
        self.fee_frame.pack(fill=tk.BOTH, expand=1, padx=20)

        self.invoice_frame = tk.LabelFrame(self.fee_frame)
        tk.Label(self.invoice_frame,width=50, text="Items", font=self.app.font).grid(row=0, column=0)
        tk.Label(self.invoice_frame,width=20, text="ex.GST", font=self.app.font).grid(row=0, column=1)
        tk.Label(self.invoice_frame,width=20, text="in.GST", font=self.app.font).grid(row=0, column=2)
        self.invoice_frame.grid(row=0, column=0)

        self.fee_dic = dict()
        self.bool_var_list = []
        self.fee_frame_number = 1

        self.total_frame = tk.LabelFrame(self.fee_frame)
        tk.Label(self.total_frame, width=50, text="Total", font=self.app.font).grid(row=0, column=0)
        self.total_ex_gst_label = tk.Label(self.total_frame, width=20, text="0", font=self.app.font)
        self.total_ex_gst_label.grid(row=0, column=1)
        self.total_in_gst_label = tk.Label(self.total_frame, width=20, text="0", font=self.app.font)
        self.total_in_gst_label.grid(row=0, column=2)
        self.total_frame.grid(row=999, column=0)

    def update_fee(self, var):
        if var.get() == True:
            if not var._name in self.fee_dic.keys():
                self.items_frame[var._name] = tk.LabelFrame(self.fee_frame)
                self.fee_dic[var._name] = {
                    "Service": tk.Label(self.items_frame[var._name],
                                        text=var._name + " design and documentation",
                                        width=50,
                                        font=self.app.font),
                    "ex.GST": tk.Label(self.items_frame[var._name],
                                       textvariable=self.app.data["Fee Proposal Page"]["Details"][var._name]["Fee"],
                                       width=20,
                                       font=self.app.font,
                                       fg="blue"),
                    "in.GST": tk.Label(self.items_frame[var._name], text=0, width=20, font=self.app.font),
                    "Position": 0,
                    "expand frames": [tk.LabelFrame(self.items_frame[var._name]) for _ in range(4)]
                    "expand frame": {
                        "Service": [tk.Label(self.fee_frame[var._name]["expand frames"][i], width=50, textvariable=
                        self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][i][0]) for i in range(3)],
                        "Fee": [tk.Label(self.fee_frame[var._name]["expand frames"][i], width=20, textvariable=
                        self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][i][1]) for i in range(3)],
                        "in.GST": [tk.Label(self.fee_frame[var._name]["expand frames"][i], width=20, text="0", font=self.app.font) for i in range(3)],
                        "Total": tk.Label(self.fee_frame[var._name]["expand frames"][3], width=50, text=var._name + " Total", font=self.app.font)
                    }
                }
                for i in range(3):
                    self.fee_dic[var._name]["expand frame"]["Service"][i].grid(row=0, column=0)
                    self.fee_dic[var._name]["expand frame"]["Fee"][i].grid(row=0, column=1)
                    self.fee_dic[var._name]["expand frame"]["in.GST"][i].grid(row=0, column=2)
                self.fee_dic[var._name]["expand frame"]["Total"].grid(row=0, column=0)



                #update function
                for i in range(3):
                    sum_fun = lambda a, b, c: self.app._sum_update(
                        [item[1] for item in self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"]],
                        self.app.data["Fee Proposal Page"]["Details"][var._name]["Fee"])

                    self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][i][1].trace("w", sum_fun)

                # garbage collection value
                ist_fun_0 = lambda a, b, c: self.app._ist_update(
                    self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][0][1],
                    self.fee_dic[var._name]["expand frame"]["in.GST"][0])
                self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][0][1].trace("w", ist_fun_0)
                ist_fun_1 = lambda a, b, c: self.app._ist_update(
                    self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][1][1],
                    self.fee_dic[var._name]["expand frame"]["in.GST"][1])
                self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][1][1].trace("w", ist_fun_1)
                ist_fun_2 = lambda a, b, c: self.app._ist_update(
                    self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][2][1],
                    self.fee_dic[var._name]["expand frame"]["in.GST"][2])
                self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][2][1].trace("w", ist_fun_2)

                self.app.data["Fee Proposal Page"]["Details"][var._name]["Fee"].trace("w",
                                                                                      lambda a, b, c: self.app._ist_update(
                                                                                          self.app.data[
                                                                                              "Fee Proposal Page"][
                                                                                              "Details"][var._name]["Fee"],
                                                                                          self.fee_dic[var._name][
                                                                                              "in.GST"]))
                self.app.data["Fee Proposal Page"]["Details"][var._name]["Fee"].trace("w",
                                                                                      lambda a, b, c: self.app._sum_update(
                                                                                          [value["Fee"] for value in
                                                                                           self.app.data[
                                                                                               "Fee Proposal Page"][
                                                                                               "Details"].values()],
                                                                                          self.total_ex_gst_label))
                self.app.data["Fee Proposal Page"]["Details"][var._name]["Fee"].trace("w", lambda a, b,
                                                                                                  c: self.app._in_sum_update(
                    [value["Fee"] for value in self.app.data["Fee Proposal Page"]["Details"].values()],
                    self.total_in_gst_label))
                ###
                ###
                # for i in range(3):
                #     self.app.data["Fee Proposal Page"]["Details"][var._name]["Context"][i][0].trace("w", lambda a,b,c: self.extra_context_append(var._name))
                ###

                self.fee_dic[var._name]["Service"].grid(row=0, column=0)
                self.fee_dic[var._name]["ex.GST"].grid(row=0, column=1)
                self.fee_dic[var._name]["in.GST"].grid(row=0, column=2)


            self.fee_dic[var._name]["Position"] = str(self.fee_frame_number)
            self.items_frame[var._name].grid(row=self.fee_frame_number, column=0)
            self.fee_frame_number += 5
        else:
            self.app.data["Fee Proposal Page"]["Details"][var._name]["Fee"].set("")
            if self.app.data["Fee Proposal Page"]["Details"][var._name]["Expanded"].get():
                self.app.data["Fee Proposal Page"]["Details"][var._name]["Expanded"].set(False)

            self.items_frame[var._name].grid_forget()


    def expand(self, service):
        if self.app.data["Fee Proposal Page"]["Details"][service]["Expanded"].get():
            self.app.data["Fee Proposal Page"]["Details"][service]["Fee"].set("")
            self.fee_dic[service]["expand frame"]["Total"].grid(row=int(self.fee_dic[service]["Position"]) + 4, column=1, pady=(0,20))
            self.fee_dic[service]["ex.GST"].grid(row=int(self.fee_dic[service]["Position"]) + 4, column=2, pady=(0,20))
            self.fee_dic[service]["in.GST"].grid(row=int(self.fee_dic[service]["Position"]) + 4, column=3, pady=(0,20))
        else:
            for i in range(3):
                self.app.data["Fee Proposal Page"]["Details"][service]["Context"][i][0].set("")
                self.app.data["Fee Proposal Page"]["Details"][service]["Context"][i][1].set("")
                self.fee_dic[service]["expand frame"]["Service"][i].grid_forget()
                self.fee_dic[service]["expand frame"]["Fee"][i].grid_forget()
                self.fee_dic[service]["expand frame"]["in.GST"][i].grid_forget()
            self.app.data["Fee Proposal Page"]["Details"][service]["Fee"].set("")
            self.fee_dic[service]["expand frame"]["Total"].grid_forget()
            self.fee_dic[service]["ex.GST"].grid(row=int(self.fee_dic[service]["Position"]), column=2,pady=0)
            self.fee_dic[service]["in.GST"].grid(row=int(self.fee_dic[service]["Position"]), column=3,pady=0)

    def extra_context_append(self, service):
        for i in range(3):
            try:
                self.fee_dic[service]["expand frame"]["Service"][i].grid_forget()
            except:
                pass
            try:
                self.fee_dic[service]["expand frame"]["Fee"][i].grid_forget()
            except:
                pass
            try:
                self.fee_dic[service]["expand frame"]["in.GST"][i].grid_forget()
            except:
                pass

            if len(self.app.data["Fee Proposal Page"]["Details"][service]["Context"][i][0].get())==0:
                continue
            self.fee_dic[service]["expand frame"]["Service"][i].grid(row=int(self.fee_dic[service]["Position"]) + i + 1,
                                                                     column=1)
            self.fee_dic[service]["expand frame"]["Fee"][i].grid(row=int(self.fee_dic[service]["Position"]) + i + 1,
                                                                 column=2)
            self.fee_dic[service]["expand frame"]["in.GST"][i].grid(row=int(self.fee_dic[service]["Position"]) + i + 1,
                                                                    column=3)
    def _on_mousewheel(self, event):
        self.main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def reset_scrollregion(self, event):
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))