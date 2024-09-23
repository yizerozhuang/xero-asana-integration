import tkinter as tk
from tkinter import messagebox
import os
import json
import socket

class Login(tk.Tk):
    def __init__(self, conf, user, user_email, login, admin, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        # Variables for Tkinter
        self.title("Login Bridge")
        self.geometry("250x100")
        self.conf = conf
        self.user = user
        self.user_email = user_email
        self.login = login
        self.admin = admin
        self.login_session_path = os.path.join(self.conf['database_dir'], "login_session.json")
        self.login_session = json.load(open(self.login_session_path))
        self.current_user = socket.gethostbyname(socket.gethostname())
        if self.current_user in self.login_session.keys():
            self.username = tk.StringVar(value=self.login_session[self.current_user])
        else:
            self.username = tk.StringVar()
        self.password = tk.StringVar()
        tk.Label(self, text="Username").grid(row=0, column=0)
        tk.Entry(self, relief=tk.FLAT, textvariable=self.username).grid(row=0, column=1)
        tk.Label(self, text="Password").grid(row=1, column=0)
        tk.Entry(self, show="*", relief=tk.FLAT, textvariable=self.password).grid(row=1, column=1)
        # Actual Variable
        self.bind("<Return>", self.validate)
        tk.Button(self, text="Submit", pady=5, padx=20, command=self.validate).grid(row=2, column=0)
    def validate(self, e=None):
        database_dir = os.path.join(self.conf["database_dir"], "login.json")
        data_json = json.load(open(database_dir))
        if not self.username.get() in data_json.keys():
            messagebox.showerror("Error", "Cant find the username")
        elif data_json[self.username.get()]["password"] != self.password.get():
            messagebox.showerror("Error", "Wrong Password")
        else:
            self.user[0] = data_json[self.username.get()]["user_name"]
            self.user_email[0] = self.username.get()
            self.login[0] = True
            if data_json[self.username.get()]["admin"]:
                self.admin[0] = True
            self.login_session[self.current_user] = self.username.get()
            with open(self.login_session_path, "w") as f:
                json_object = json.dumps(self.login_session, indent=4)
                f.write(json_object)

            self.destroy()
