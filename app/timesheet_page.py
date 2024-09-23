import os
import json
import tkinter as tk
from tkinter import ttk
from datetime import date, timedelta
import asana
from asana_function import clean_response, flatter_custom_fields
from config import CONFIGURATION as conf

asana_configuration = asana.Configuration()
asana_configuration.access_token = '2/1203283895754383/1206354773081941:c116d68430be7b2832bf5d7ea2a0a415'
asana_api_client = asana.ApiClient(asana_configuration)

task_api_instance = asana.TasksApi(asana_api_client)
user_task_lists_api_instance = asana.UserTaskListsApi(asana_api_client)
workspace_gid = '1198726743417674'


class TimeSheetPage(tk.Frame):

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
        self.user_map = json.load(open(os.path.join(conf["database_dir"], "login.json")))

        self.time_sheet()

    def get_weekday_from_date(self, date):
        return int(date.strftime("%w"))

    # def get_date_from_weeks(self):
    #     since = -1
    #     return self.get_week_from_date(date.today())

    def get_date_of_sunday(self, week):
        today = date.today()
        weekday = int(today.strftime("%w"))
        return today - timedelta(weeks=-week, days=weekday)

    def time_sheet(self):
        container = ttk.Frame(self.main_context_frame)
        container.pack(fill='both', expand=True)
        utility_frame = tk.Frame(container)
        utility_frame.grid(row=0, column=0, sticky="ew")

        tk.Button(utility_frame, text="Sync", bg="Brown", fg="white", command=self.sync,
                  font=self.conf["font"]).grid(row=0, column=0)

        # self.sync()

    def sync(self, user_asana_id):
        # user_asana_id = self.user_map[self.app.user]["asana_id"]

        preview_week_sunday = self.get_date_of_sunday(-1)
        this_week_sunday = self.get_date_of_sunday(0)
        next_week_sunday = self.get_date_of_sunday(1)
        the_week_after_sunday = self.get_date_of_sunday(2)

        opt_list = ["name", "parent.name", "custom_fields.display_value", "custom_fields.name", "due_on",
                    "completed", "permalink_url", "actual_time_minutes"]
        attribute_list = ["name", "parent_name", "Estimated time", "due_on", "completed", "permalink_url", "actual_time_minutes"]
        all_tasks = clean_response(task_api_instance.get_tasks(workspace=workspace_gid, assignee=user_asana_id,
                                                               completed_since=preview_week_sunday,
                                                               opt_fields=opt_list))

        # user_task_list_id = clean_response(user_task_lists_api_instance.get_user_task_list_for_user(user_gid=user_asana_id, workspace=workspace_gid))["gid"]

        # all_tasks = clean_response(
        #     task_api_instance.get_tasks_for_user_task_list(user_task_list_id, completed_since=preview_week_sunday, opt_fields=opt_list))

        self.preview_week = [[] for _ in range(7)]
        self.current_week = [[] for _ in range(7)]
        self.next_week = [[] for _ in range(7)]
        for task in all_tasks:
            task = flatter_custom_fields(task)
            if "Estimated time" in task.keys():
                if not task["due_on"] is None:
                    if task["due_on"] <= the_week_after_sunday:
                        weekday = self.get_weekday_from_date(task["due_on"])

                        task["parent_name"] = task["parent"]["name"]

                        new_task = {k: v for k, v in task.items() if k in attribute_list}

                        if task["due_on"] >= next_week_sunday:
                            self.next_week[weekday].append(new_task)
                        elif task["due_on"] >= this_week_sunday:
                            self.current_week[weekday].append(new_task)
                        else:
                            self.preview_week[weekday].append(new_task)
        return self.preview_week, self.current_week, self.next_week


    def reset_scrollregion(self, event):
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
