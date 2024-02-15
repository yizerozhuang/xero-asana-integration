import tkinter as tk
import tkinter.font as tkFont
from tkinter import ttk
import json
import os

from config import CONFIGURATION as conf
from utility import load_data


def sortby(tree, col, descending):
    """sort tree contents when a column header is clicked on"""
    # grab values to sort
    data = [(tree.set(child, col), child) for child in tree.get_children('')]
    # if the data to be sorted is numeric change to float
    # data =  change_numeric(data)
    # now sort the data in place
    data.sort(reverse=descending)
    for ix, item in enumerate(data):
        tree.move(item[1], '', ix)
    # switch the heading so it will sort in the opposite direction
    tree.heading(col, command=lambda col=col: sortby(tree, col, int(not descending)))

def show_status(data_json):
    state = data_json["State"]
    if state["Asana State"] in ["Pending", "DWG drawings", "Done", "Installation", "Construction Phase"]:
        return state["Asana State"]
    if state["Quote Unsuccessful"]:
        return "Quote Unsuccessful"
    elif state["Fee Accepted"]:
        return "Design"
    elif state["Email to Client"]:
        return "Chase Client"
    elif state["Generate Proposal"]:
        return "Email To Client"
    elif state["Set Up"]:
        return "Preview Fee Proposal"
    else:
        return "Set Up"


class SearchBarPage(tk.Frame):
    def __init__(self, parent, app):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.app = app

        self.main_frame = tk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=1)

        self.main_canvas = tk.Canvas(self.main_frame)
        self.main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1, anchor="nw")
        # # self.scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.main_canvas.yview)
        # # self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        # # self.main_canvas.config(yscrollcommand=self.scrollbar.set)
        # self.main_canvas.bind("<Configure>", lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")))
        # # self.app.bind("<MouseWheel>", self._on_mousewheel, add="+")
        self.main_context_frame = tk.LabelFrame(self.main_canvas, text="Main Context")
        self.main_canvas.create_window((0, 0), window=self.main_context_frame, anchor="nw")
        self.main_context_frame.bind("<Configure>", self.reset_scrollregion)
        self.conf = self.app.conf


        self.generate_convert_map()
        self.search_bar()
        self.build_tree()
        self.reset()


    def reset(self):
        self.generate_data()
        self.update_data(self.master_project)


    def generate_convert_map(self):
        if self.app.user in self.conf["engineer_user_list"]:
            self.convert_map = {
                "Quotation Number": lambda data_json: data_json["Project Info"]["Project"]["Quotation Number"],
                "Project Number": lambda data_json: data_json["Project Info"]["Project"]["Project Number"],
                "Project Name": lambda data_json: data_json["Project Info"]["Project"]["Project Name"],
                "Shop Name": lambda data_json: data_json["Project Info"]["Project"]["Shop Name"],
                "Proposal Type": lambda data_json: data_json["Project Info"]["Project"]["Proposal Type"],
                "Project Type": lambda data_json: data_json["Project Info"]["Project"]["Project Type"],
                "Project Status": show_status,
                "Address To": lambda data_json: data_json["Address_to"],
                "Building Feature": lambda data_json: data_json["Project Info"]["Building Features"]["Feature"],
                "Service": lambda data_json: ", ".join(
                    [service["Service"] for service in data_json["Project Info"]["Project"]["Service Type"].values() if
                     service["Include"]]),
                "Fee Proposal Date": lambda data_json: data_json["Email"]["Fee Proposal"],
                "Apt/Room/Area": lambda data_json: data_json["Project Info"]["Building Features"]["Apt"],
                "Basement/Car Spots": lambda data_json: data_json["Project Info"]["Building Features"]["Basement"],
                "Feature/Notes": lambda data_json: data_json["Project Info"]["Building Features"]["Feature"],
                "INV1": lambda data_json: data_json["Invoices Number"][0]["Number"],
                "INV2": lambda data_json: data_json["Invoices Number"][1]["Number"],
                "INV3": lambda data_json: data_json["Invoices Number"][2]["Number"],
                "INV4": lambda data_json: data_json["Invoices Number"][3]["Number"],
                "INV5": lambda data_json: data_json["Invoices Number"][4]["Number"],
                "INV6": lambda data_json: data_json["Invoices Number"][5]["Number"]
            }
        else:
            self.convert_map = {
                "Quotation Number": lambda data_json: data_json["Project Info"]["Project"]["Quotation Number"],
                "Project Number": lambda data_json: data_json["Project Info"]["Project"]["Project Number"],
                "Project Name": lambda data_json: data_json["Project Info"]["Project"]["Project Name"],
                "Shop Name": lambda data_json: data_json["Project Info"]["Project"]["Shop Name"],
                "Proposal Type": lambda data_json: data_json["Project Info"]["Project"]["Proposal Type"],
                "Project Type": lambda data_json: data_json["Project Info"]["Project"]["Project Type"],
                "Project Status": show_status,
                "Address To": lambda data_json: data_json["Address_to"],
                "Building Feature": lambda data_json: data_json["Project Info"]["Building Features"]["Feature"],
                "Service": lambda data_json: ", ".join(
                    [service["Service"] for service in data_json["Project Info"]["Project"]["Service Type"].values() if
                     service["Include"]]),
                "Fee Proposal Date": lambda data_json: data_json["Email"]["Fee Proposal"],
                "Client Name": lambda data_json: data_json["Project Info"]["Client"]["Full Name"],
                "Client Contact Type": lambda data_json: data_json["Project Info"]["Client"]["Contact Type"],
                "Client Company": lambda data_json: data_json["Project Info"]["Client"]["Company"],
                "Main Contact Name": lambda data_json: data_json["Project Info"]["Main Contact"]["Full Name"],
                "Main Contact Contact Type": lambda data_json: data_json["Project Info"]["Main Contact"]["Contact Type"],
                "Main Contact Company": lambda data_json: data_json["Project Info"]["Main Contact"]["Company"],
                "Apt/Room/Area": lambda data_json: data_json["Project Info"]["Building Features"]["Apt"],
                "Basement/Car Spots": lambda data_json: data_json["Project Info"]["Building Features"]["Basement"],
                "Feature/Notes": lambda data_json: data_json["Project Info"]["Building Features"]["Feature"],
                "INV1": lambda data_json: data_json["Invoices Number"][0]["Number"],
                "INV2": lambda data_json: data_json["Invoices Number"][1]["Number"],
                "INV3": lambda data_json: data_json["Invoices Number"][2]["Number"],
                "INV4": lambda data_json: data_json["Invoices Number"][3]["Number"],
                "INV5": lambda data_json: data_json["Invoices Number"][4]["Number"],
                "INV6": lambda data_json: data_json["Invoices Number"][5]["Number"],
                "Paid Amount": lambda data_json: data_json["Invoices"]["Paid Fee"],
                "Total Fee Amount exGST": lambda data_json: data_json["Invoices"]["Fee"],
                "Total Bill Amount exGST": lambda data_json: data_json["Bills"]["Fee"],
            }
        self.mp_header = list(self.convert_map.keys())


    def generate_data(self):
        database_dir = conf["database_dir"]
        res = []
        for dir in os.listdir(database_dir):
            if os.path.isdir(os.path.join(database_dir, dir)):
                data_dir = os.path.join(database_dir, dir, "data.json")
                data_json = json.load(open(data_dir))
                res.append(
                    tuple([self.convert_map[title](data_json) for title in self.mp_header])
                )
        # res.sort(key=lambda e: e[0])
        self.master_project = res

    def search_bar(self):
        # search_bar_frame = tk.LabelFrame(self.main_context_frame)
        # search_bar_frame.pack()

        container = ttk.Frame(self.main_context_frame)
        container.pack(fill='both', expand=True)


        search_frame = tk.Frame(container)
        search_frame.grid(row=0, column=0, sticky="ew")
        self.entry = tk.Entry(search_frame, font=self.conf["font"], fg="blue", width=200)
        self.entry.grid(row=0, column=0, sticky="ew")
        tk.Button(search_frame, text="Refresh", bg="Brown", fg="white", command=self.refresh, font=self.conf["font"]).grid(row=0, column=1)


        self.tree = ttk.Treeview(container, height=35, columns=self.mp_header, show="headings", selectmode="browse")
        treeXScroll = ttk.Scrollbar(container, orient=tk.HORIZONTAL)
        treeXScroll.configure(command=self.main_canvas.xview)
        self.main_canvas.configure(xscrollcommand=treeXScroll.set)
        treeXScroll.grid(row=2, column=0, sticky="ew")
        # treeXScroll.pack(side=tk.BOTTOM)
        # hsb = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        # self.tree.config(xscrollcommand=hsb.set)
        # hsb.config(command=self.tree.xview)
        # hsb.grid(row=2, column=0, sticky='ew')
        self.tree.grid(row=1, column=0, sticky='nsew')


        self.entry.bind("<KeyRelease>", self.check)
        self.tree.bind("<<TreeviewSelect>>", self.load_project)

    def build_tree(self):
        for col in self.mp_header:
            self.tree.heading(col, text=col.title(), command=lambda c=col: sortby(self.tree, c, 0))
            # adjust the column's width to the header string
            self.tree.column(col, width=tkFont.Font().measure(col.title()))
    def update_data(self, data):
        self.tree.delete(*self.tree.get_children())
        for item in data:
            self.tree.insert('', 'end', values=item)
            # adjust column's width if necessary to fit each value
            # for ix, val in enumerate(item):
                # col_w = tkFont.Font().measure(val)
                # if self.tree.column(self.mp_header[ix], width=None)<col_w: self.tree.column(self.mp_header[ix], width=col_w)
    def load_project(self, event):
        self.tree.selection()
        selected = self.tree.focus()
        value = self.tree.item(selected, "values")
        self.entry.delete(0, tk.END)
        # self.app.data["Project Info"]["Project"]["Quotation Number"].set(value[0])
        load_data(self.app, value[0])
    def check(self, e):
        typed = self.entry.get()
        if typed == "":
            data = self.master_project
        else:
            data = []
            for item in self.master_project:
                for title in item:
                    if title is None:
                        print()
                    if typed.lower() in title.lower():
                        data.append(item)
        self.update_data(set(data))

    def refresh(self):
        self.entry.delete(0, "end")
        self.generate_data()
        self.check("None")
        sortby(self.tree, "Quotation Number", True)

    def reset_scrollregion(self, event):
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
    # def left_arrow(self, e):
    #     self.tree.xview_scroll(int(-500000), "units")
    # def right_arrow(self, e):
    #     self.tree.xview_scroll(int(50000), "units")

    # def reset_scrollregion(self, event):
    #     self.tree.configure(scrollregion=self.tree)