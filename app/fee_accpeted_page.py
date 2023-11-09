import tkinter as tk
from tkinter import ttk, messagebox
from tkinterdnd2 import DND_FILES

from utility import save, config_log, config_state
from asana_function import update_asana

import os
import shutil


class FeeAcceptedPage(tk.Frame):
    def __init__(self, parent, app):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.app = app
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

        self.main_part()

    def main_part(self):
        self.data["Log Files"] = {
            "Type": tk.StringVar(value="None"),
            "External Files": tk.StringVar()
        }
        self.data["Log Files"]["External Files"].trace("w", self.update_list_box)
        drop_file_frame = tk.Frame(self.main_context_frame)
        drop_file_frame.pack(side=tk.BOTTOM)

        select_type_frame = tk.Frame(self.main_context_frame)
        select_type_frame.pack(side=tk.LEFT)
        tk.Radiobutton(select_type_frame, variable=self.data["Log Files"]["Type"], value="Verbal",
                       text="Verbal").grid(row=0, column=0, columnspan=2, sticky="W")
        tk.Radiobutton(select_type_frame, variable=self.data["Log Files"]["Type"], value="Email",
                       text="Email").grid(row=1, column=0, columnspan=2, sticky="W")
        tk.Radiobutton(select_type_frame, variable=self.data["Log Files"]["Type"], value="PDF",
                       text="PDF").grid(row=2, column=0, columnspan=2, sticky="W")
        tk.Radiobutton(select_type_frame, variable=self.data["Log Files"]["Type"], value="Others").grid(row=3, column=0, sticky="W")
        tk.Entry(select_type_frame).grid(row=3, column=1, sticky="W")


        tk.Button(drop_file_frame, text="Upload", bg="brown", fg="white", command=self.upload_files,
                  font=self.conf["font"]).pack(side=tk.BOTTOM)
        self.drop_file_listbox = tk.Listbox(
            drop_file_frame,
            width=100,
            height=40,
            selectmode=tk.SINGLE
        )
        self.drop_file_listbox.pack(fill=tk.X, side=tk.LEFT)
        self.drop_file_listbox.drop_target_register(DND_FILES)
        self.drop_file_listbox.dnd_bind("<<Drop>>", self.drop_inside_list_box)


    def drop_inside_list_box(self, event):
        self.data["Log Files"]["External Files"].set(event.data.split("{")[-1].split("}")[0])
    def update_list_box(self, *args):
        self.drop_file_listbox.delete(0)
        self.drop_file_listbox.insert(0, self.data["Log Files"]["External Files"].get().split("/")[-1])
    def upload_files(self):
        database_dir = os.path.join(self.conf["database_dir"], self.data["Project Info"]["Project"]["Quotation Number"].get())
        if not self.data["State"]["Fee Accepted"].get():
            messagebox.showerror("Error", "You cant update the fee acceptance before you email to client")
            return
        elif self.data["Log Files"]["Type"].get() == "None":
            messagebox.showerror("Error", "Please select a type first")
            return

        if self.data["Log Files"]["Type"].get() == "Verbal":
            yes = messagebox.askyesno("Warning", "Are you sure Client give the verbal approve and move this project into design state?")
            if yes:
                self.data["State"]["Done"].set(True)
                save(self.app)
                self.app.log.log_fee_accept_file(self.app)
                config_log(self.app)
                config_state(self.app)

        else:
            rewrite = True
            if len(self.data["Log Files"]["External Files"].get()) == 0:
                messagebox.showerror("Error", "Please upload a file before uploading")
                return
            elif os.path.exists(os.path.join(database_dir, "Fee accepted Files")):
                rewrite = messagebox.askyesno("Overwrite", "Previous Fee Accepted found, do you want to overwrite")

            if rewrite:
                if os.path.exists(os.path.join(database_dir, "Fee accepted Files")):
                    shutil.rmtree(os.path.join(database_dir, "Fee accepted Files"))
                os.mkdir(os.path.join(database_dir, "Fee accepted Files"))
                shutil.copy(self.data["Log Files"]["External Files"].get(), os.path.join(database_dir, "Fee accepted Files"))
                update_asana(self.app)
                messagebox.showinfo("Update", "Successfully log the file into database and update asana")
                self.data["State"]["Done"].set(True)
                save(self.app)
                self.app.log.log_fee_accept_file(self.app)
                config_log(self.app)
                config_state(self.app)

    def _on_mousewheel(self, event):
        self.main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
