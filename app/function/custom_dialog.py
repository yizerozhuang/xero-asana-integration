import tkinter as tk
from tkinter import simpledialog, ttk
import os
import shutil


class FileSelectDialog(simpledialog.Dialog):
    def __init__(self, master, app, dir_list, title=None):
        self.dir_list = dir_list
        self.app = app
        super().__init__(master, title)

    def body(self, master):
        self.rename_dir = tk.StringVar()
        ttk.Combobox(self, textvariable=self.rename_dir, values=self.dir_list).pack()

    def ok(self):
        external_files = os.listdir(self.rename_dir.get())
        mode = 0o666
        os.mkdir(self.rename_dir.get()+"/External", mode)
        for file in external_files:
            shutil.move(self.rename_dir.get()+"/"+file, self.rename_dir.get()+"/External")
        os.mkdir(self.rename_dir.get()+"/Photos", mode)
        os.mkdir(self.rename_dir.get()+"/Plot", mode)
        os.mkdir(self.rename_dir.get()+"/SS", mode)
        shutil.copyfile("resource/xlsx/Preliminary Calculation v2.5.xlsx",
                        self.rename_dir.get()+"/Preliminary Calculation v2.5.xlsx")


        os.rename(self.rename_dir.get(),
                  self.app.data["Project Information"]["Quotation Number"].get() + "-" + self.app.data["Project Information"]["Project Name"].get())
        self.destroy()

    def cancel(self):
        self.destroy()
