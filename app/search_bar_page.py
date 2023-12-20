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


class SearchBarPage(tk.Frame):
    def __init__(self, parent, app):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.app = app

        self.main_frame = tk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=1)
        # self.main_canvas = tk.Canvas(self.main_frame)
        # self.main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1, anchor="nw")
        # # self.scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.main_canvas.yview)
        # # self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        # # self.main_canvas.config(yscrollcommand=self.scrollbar.set)
        # self.main_canvas.bind("<Configure>", lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")))
        # # self.app.bind("<MouseWheel>", self._on_mousewheel, add="+")
        # self.main_context_frame = tk.LabelFrame(self.main_canvas, text="Main Context")
        # self.main_canvas.create_window((0, 0), window=self.main_context_frame, anchor="nw")
        self.conf = self.app.conf


        self.generate_convert_map()
        self.search_bar()
        self.build_tree()
        self.reset()


    def reset(self):
        self.generate_data()
        self.update_data(self.master_project)


    def generate_convert_map(self):
        self.convert_map = {
            "Quotation Number": lambda data_json: data_json["Project Info"]["Project"]["Quotation Number"],
            "Project Number": lambda data_json: data_json["Project Info"]["Project"]["Project Number"],
            "Project Name": lambda data_json: data_json["Project Info"]["Project"]["Project Name"],
            "Shop Name": lambda data_json: data_json["Project Info"]["Project"]["Shop Name"],
            "Proposal Type": lambda data_json: data_json["Project Info"]["Project"]["Proposal Type"],
            "Address To": lambda data_json: data_json["Address_to"],
            "Building Feature": lambda data_json: data_json["Project Info"]["Building Features"]["Feature"],
            "Service": lambda data_json: ", ".join([service["Service"] for service in data_json["Project Info"]["Project"]["Service Type"].values() if service["Include"]]),
            "Fee Proposal data": lambda data_json: data_json["Email"]["Fee Proposal"],
            "Client Name": lambda data_json: data_json["Project Info"]["Client"]["Full Name"],
            "Client Contact Type": lambda data_json: data_json["Project Info"]["Client"]["Contact Type"],
            "Client Company": lambda data_json: data_json["Project Info"]["Client"]["Company"],
            "Main Contact Name": lambda data_json: data_json["Project Info"]["Main Contact"]["Full Name"],
            "Main Contact Contact Type": lambda data_json: data_json["Project Info"]["Main Contact"]["Contact Type"],
            "Main Contact Company": lambda data_json: data_json["Project Info"]["Main Contact"]["Company"]
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
        self.master_project = res

    def search_bar(self):
        # search_bar_frame = tk.LabelFrame(self.main_context_frame)
        # search_bar_frame.pack()

        container = ttk.Frame(self.main_frame)
        container.pack(fill='both', expand=True)

        self.entry = tk.Entry(container, font=self.conf["font"], fg="blue", width=50)
        self.entry.grid(row=0, column=0, sticky="ew")

        self.tree = ttk.Treeview(container, height=20, columns=self.mp_header, show="headings", selectmode="browse")
        # vsb = ttk.Scrollbar(container, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(container, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.config(xscrollcommand=hsb.set)

        # self.tree.config(yscrollcommand=vsb.set)
        #
        # vsb.config(command=self.tree.yview)
        # hsb.config(command=self.tree.xview)
        self.tree.grid(row=1, column=0, sticky='nsew')
        # vsb.grid(row=1, column=1, sticky='ns')
        hsb.grid(row=2, column=0, sticky='ew')
        # container.grid_columnconfigure(0, weight=1)
        # container.grid_rowconfigure(0, weight=1)

        self.entry.bind("<KeyRelease>", self.check)
        self.tree.bind("<<TreeviewSelect>>", self.load_project)
        # self.tree.bind("<Left>", self.left_arrow)
        # self.tree.bind("<Right>", self.right_arrow)
        # self.tree.focus_set()
        # self.tree.bind("<Configure>", self.reset_scrollregion)

        # self.tree.bind("<Double-1>", self.clicker)

        # self.list_box = tk.Listbox(search_bar_frame, width=50)
        # self.list_box.pack(pady=40)
        # self.update_data(self.dummy_list)

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
            for ix, val in enumerate(item):
                col_w = tkFont.Font().measure(val)
                if self.tree.column(self.mp_header[ix], width=None)<col_w: self.tree.column(self.mp_header[ix], width=col_w)
    def load_project(self, event):
        self.tree.selection()
        selected = self.tree.focus()
        value = self.tree.item(selected, "values")
        self.entry.delete(0, tk.END)
        self.app.data["Project Info"]["Project"]["Quotation Number"].set(value[0])
        load_data(self.app)

    def check(self, e):
        typed = self.entry.get()
        if typed == "":
            data = self.master_project
        else:
            data = []
            for item in self.master_project:
                for title in item:
                    if typed.lower() in title.lower():
                        data.append(item)
        self.update_data(data)
    def _on_mousewheel(self, event):
        self.main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # def left_arrow(self, e):
    #     self.tree.xview_scroll(int(-500000), "units")
    # def right_arrow(self, e):
    #     self.tree.xview_scroll(int(50000), "units")

    # def reset_scrollregion(self, event):
    #     self.tree.configure(scrollregion=self.tree)