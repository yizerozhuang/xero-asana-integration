import tkinter as tk
from tkinter import simpledialog, ttk, messagebox

# from main import CONFIGURATION

import os
import shutil


class AppDialog:
    def __init__(self, master=None, **options):
        if not master:
            master = options.get('parent')
        self.master = master
        self.options = options
    def show(self):
        pass

class FileSelectDialog(simpledialog.Dialog):
    def __init__(self, app, dir_list, title=None):
        self.dir_list = dir_list
        self.app = app
        self.conf = app.conf
        super().__init__(app, title)

    def body(self, master):
        self.rename_dir = tk.StringVar()
        ttk.Combobox(self, textvariable=self.rename_dir, values=self.dir_list).pack()

    def ok(self):

        resource_dir = self.conf["resource_dir"]
        calculation_sheet = self.conf["calculation_sheet"]
        mode = 0o666
        folder_name = self.app.data["Project Info"]["Project"]["Quotation Number"].get() + "-" + \
                      self.app.data["Project Info"]["Project"]["Project Name"].get()
        folder_path = os.path.join(self.app.conf["working_dir"], folder_name)
        os.mkdir(folder_path, mode)
        shutil.move(self.rename_dir.get(), os.path.join(folder_path, "External"))
        os.mkdir(os.path.join(folder_path, "Photos"), mode)
        os.mkdir(os.path.join(folder_path, "Plot"), mode)
        os.mkdir(os.path.join(folder_path, "SS"), mode)

        shutil.copyfile(os.path.join(resource_dir, "xlsx", calculation_sheet),
                        os.path.join(folder_path, calculation_sheet))
        self.app.data["State"]["Folder Renamed"].set(True)
        self.destroy()
        messagebox.showinfo(title="Folder renamed", message=f"Rename Folder {self.rename_dir.get()} to {folder_path}")
    def cancel(self):
        self.destroy()
#
# class message_box(tk.messagebox)

# class ConfirmDialog(simpledialog.Dialog):
#     def __init__(self, master, title=None):
#         super().__init__(master, title)
#         self.exit=False
#
#     def body(self, master):
#         tk.Label(self, text="Do you want to save the change before exist this program").pack()
#
#     def buttonbox(self):
#         self.ok_button = tk.Button(self, text="Ok", width=20, command=self.ok_pressed)
#         self.ok_button.pack(side=tk.LEFT)
#         self.no_button = tk.Button(self, text="NO", width=20, command=self.no_pressed)
#         self.no_button.pack(side=tk.RIGHT)
#         self.bind("<Return>", lambda event:self.ok_pressed())
#         self.bind("<Escape>", lambda event:self.no_pressed())
#
#     def ok_pressed(self):
#         # data_json = convert_to_json(self.master.data)
#         # if not os.path.exists(os.getcwd()+"\\database\\"+data_json["Project Info"]["Project"]["Quotation Number"]):
#         #     os.mkdir(os.getcwd()+"\\database\\"+data_json["Project Info"]["Project"]["Quotation Number"])
#         # with open(os.getcwd()+"\\database\\"+data_json["Project Info"]["Project"]["Quotation Number"]+"\\data.json", "w") as f:
#         #     json_object = json.dumps(data_json, indent=4)
#         #     f.write(json_object)
#         self.master.destroy()
#     def no_pressed(self):
#         self.master.destroy()

# class KillProcessDialog(simpledialog.Dialog):
#     def __init__(self, master, title=None):
#         super().__init__(master, title)
#     def body(self, master):
#         tk.Label(self, text="Python is about to close the excel background process, please save before use this").pack()
#     def ok(self):
#         for proc in psutil.process_iter():
#             if proc.name() == "excel.exe" or proc.name() == "EXCEL.EXE":
#                 proc.kill()
#         self.destroy()
#     def cancel(self):
#         self.destroy()
