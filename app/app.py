import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
import os
import shutil

def create_new_folder():
    try:
        os.mkdir("./new folder")
        os.mkdir("./new folder/External")
        os.mkdir("./new folder/Photos")
        os.mkdir("./new folder/Plot")
        os.mkdir("./new folder/SS")
        shutil.copyfile("./Resource/xlsx/Preliminary Calculation v2.5.xlsx", "./new folder/Preliminary Calculation v2.5.xlsx")
    except:
        messagebox.showwarning(title="Error", message="The new folder already exist")

def update_asana():
    pass

window = tk.Tk()
window.title("Data Entry Form")
frame = tk.Frame(window)
frame.pack()

#Utility Part
utility_frame = tk.LabelFrame(frame, text="Utility")
utility_frame.grid(row=0, column=0, padx=20, pady=20)


create_folder = tk.Button(utility_frame, text="Create new Folder", command=create_new_folder, bg="brown", fg="white")
create_folder.grid(row=0, column=0)

update_asana = tk.Button(utility_frame, text="Update Asana", command=update_asana, bg="brown", fg="white")
update_asana.grid(row=0, column=1)

#User_information
user_info_frame=tk.LabelFrame(frame, text="User Information")
user_info_frame.grid(row=1, column=0, padx=20)

user_info = ["Project Name", "Quotation Number", "Project Number"]
user_info_dict = dict()

for i, info in enumerate(user_info):
    user_info_dict[info+" Label"] = tk.Label(user_info_frame, text=info)
    user_info_dict[info+" Label"].grid(row=i, column=0)
    user_info_dict[info+" Entry"] = tk.Entry(user_info_frame, width=100)
    user_info_dict[info+" Entry"].grid(row=i,column=1, columnspan = 7)


project_type_label = tk.Label(user_info_frame, text="Project Type")
project_type_label.grid(row=3, column=0)
options = ["Restaurant", "Office", "Commercial", "Group house", "Apartment", "Mixed-use complex", "School", "Others"]
drop = ttk.Combobox(user_info_frame,values=options)
drop.grid(row=3,column=1,columnspan = 7)

service_type = ["Fire Service", "Kitchen Ventilation", "Mechanical Service", "CFD Service", "Mech Review", "Installation", "Miscellaneous"]

service_label = tk.Label(user_info_frame, text="Service Type")
service_label.grid(row=4, column=0)

services = dict()
for i, service in enumerate(service_type):
        services[service] = tk.Checkbutton(user_info_frame, text=service)
        services[service].grid(row=4, column=i+1)


submit = tk.Button(user_info_frame, text="Submit",bg="brown", fg="white")
submit.grid(row=5, column=0)

window.mainloop()