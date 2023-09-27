import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
import os
import shutil
from function.app import App


# def create_new_folder():
#     try:
#         os.mkdir("./new folder")
#         os.mkdir("./new folder/External")
#         os.mkdir("./new folder/Photos")
#         os.mkdir("./new folder/Plot")
#         os.mkdir("./new folder/SS")
#         shutil.copyfile("resource/xlsx/Preliminary Calculation v2.5.xlsx",
#                         "./new folder/Preliminary Calculation v2.5.xlsx")
#     except:
#         messagebox.showwarning(title="Error", message="The new folder already exist")
#
#
# def update_asana():
#     pass
#
#
# def area_update(*args):
#     sum = 0
#     for area in area_list:
#         if area.get() == "":
#             continue
#         try:
#             sum += float(area.get())
#         except:
#             sum = "error"
#             break
#     total_area_label.config(text=str(sum))
#
#
# window = tk.Tk()
# window.title("Data Entry Form")
# frame = tk.Frame(window)
# frame.pack()
#
# # Utility Part
# utility_frame = tk.LabelFrame(frame, text="Utility")
# utility_frame.grid(row=0, column=0, padx=20)
#
# create_folder = tk.Button(utility_frame, text="Create new Folder", command=create_new_folder, bg="brown", fg="white")
# create_folder.grid(row=0, column=0)
#
# update_asana = tk.Button(utility_frame, text="Update Asana", command=update_asana, bg="brown", fg="white")
# update_asana.grid(row=0, column=1)
#
# # User_information
# user_info_frame = tk.LabelFrame(frame, text="User Information")
# user_info_frame.grid(row=1, column=0, padx=20)
#
# user_info = ["Project Name", "Quotation Number", "Project Number"]
# user_info_dict = dict()
#
# for i, info in enumerate(user_info):
#     user_info_dict[info + " Label"] = tk.Label(user_info_frame, text=info)
#     user_info_dict[info + " Label"].grid(row=i, column=0)
#     user_info_dict[info + " Entry"] = tk.Entry(user_info_frame, width=100)
#     user_info_dict[info + " Entry"].grid(row=i, column=1, columnspan=7)
#
# project_type_label = tk.Label(user_info_frame, text="Project Type(Single)")
# project_type_label.grid(row=3, column=0)
# project_types = ["Restaurant", "Office", "Commercial", "Group house", "Apartment", "Mixed-use complex", "School", "Others"]
# project_type_dic = dict()
# project_type = tk.StringVar()
#
# for i, p_type in enumerate(project_types):
#     project_type_dic["type"+str(i)] = tk.Radiobutton(user_info_frame, text=p_type, variable=project_type, value=p_type)
#     project_type_dic["type"+str(i)].grid(row=3, column=i+1)
#
# project_type_dic["None"]=tk.Radiobutton(user_info_frame, text="None", variable=project_type, value = "None")
# project_type.set("None")
#
#
# service_type = ["Fire Service", "Kitchen Ventilation", "Mechanical Service", "CFD Service", "Mech Review",
#                 "Installation", "Miscellaneous"]
#
# service_label = tk.Label(user_info_frame, text="Service Type(Multiple)")
# service_label.grid(row=4, column=0)
#
# services = dict()
# for i, service in enumerate(service_type):
#     services[service] = tk.Checkbutton(user_info_frame, text=service)
#     services[service].grid(row=4, column=i + 1)
#
# client_frame = tk.LabelFrame(frame, text="Client")
# client_frame.grid(row=2, column=0, padx=20)
#
# client_info = ["Client Full Name", "Client Company", "Client Address", "Client ABN", "Contact Number", "Contact Email"]
# for i, info in enumerate(client_info):
#     user_info_dict[info + " Label"] = tk.Label(client_frame, text=info)
#     user_info_dict[info + " Label"].grid(row=i, column=0)
#     user_info_dict[info + " Entry"] = tk.Entry(client_frame, width=100)
#     user_info_dict[info + " Entry"].grid(row=i, column=1)
#
# contact_frame = tk.LabelFrame(frame, text="Main Contact")
# contact_frame.grid(row=3, column=0, padx=20)
#
# contact_info = ["Main Contact Full Name", "Main Contact Company", "Main Contact Address", "Main Contact ABN",
#                 "Main Contact Number", "Main Contact Email"]
# for i, info in enumerate(contact_info):
#     user_info_dict[info + " Label"] = tk.Label(contact_frame, text=info)
#     user_info_dict[info + " Label"].grid(row=i, column=0)
#     user_info_dict[info + " Entry"] = tk.Entry(contact_frame, width=100)
#     user_info_dict[info + " Entry"].grid(row=i, column=1)
#
# build_feature_frame = tk.LabelFrame(frame, text="Building Features")
# build_feature_frame.grid(row=4, column=0, padx=20)
#
# levels_label = tk.Label(build_feature_frame, text="Levels")
# levels_label.grid(row=0, column=0)
# space_label = tk.Label(build_feature_frame, text="Space/room Description")
# space_label.grid(row=0, column=1)
# area_label = tk.Label(build_feature_frame, text="Area/m^2")
# area_label.grid(row=0, column=2)
#
# n_building = 5
# building_dic = dict()
# area_list = []
# for i in range(n_building):
#     building_dic["Level " + str(i + 1)] = tk.Entry(build_feature_frame, width=30)
#     building_dic["Level " + str(i + 1)].grid(row=i + 1, column=0)
#     building_dic["Space " + str(i + 1)] = tk.Entry(build_feature_frame, width=50)
#     building_dic["Space " + str(i + 1)].grid(row=i + 1, column=1)
#     area_list.append(tk.StringVar())
#     building_dic["Area " + str(i + 1)] = tk.Entry(build_feature_frame, textvariable=area_list[i], width=20)
#     building_dic["Area " + str(i + 1)].grid(row=i + 1, column=2)
#
# total_label = tk.Label(build_feature_frame, text="Total")
# total_label.grid(row=n_building + 1, column=0)
# total_area_label = tk.Label(build_feature_frame, text="0")
# total_area_label.grid(row=n_building+1, column=2)
#
# [s.trace("w", area_update) for s in area_list]
#
# drawing_number_frame = tk.LabelFrame(frame, text="Drawing Number")
# drawing_number_frame.grid(row=5, column=0, padx=20)
#
# levels_label = tk.Label(drawing_number_frame, text="Drawing Number")
# levels_label.grid(row=0, column=0)
# space_label = tk.Label(drawing_number_frame, text="Drawing Name")
# space_label.grid(row=0, column=1)
# area_label = tk.Label(drawing_number_frame, text="Revision")
# area_label.grid(row=0, column=2)
#
# n_drawing = 5
# drawing_dic = dict()
# for i in range(n_drawing):
#     drawing_dic["Number " + str(i + 1)] = tk.Entry(drawing_number_frame, width=30)
#     drawing_dic["Number " + str(i + 1)].grid(row=i + 1, column=0)
#     drawing_dic["Name " + str(i + 1)] = tk.Entry(drawing_number_frame, width=50)
#     drawing_dic["Name " + str(i + 1)].grid(row=i + 1, column=1)
#     drawing_dic["Revision " + str(i + 1)] = tk.Entry(drawing_number_frame, width=20)
#     drawing_dic["Revision " + str(i + 1)].grid(row=i + 1, column=2)
#
# submit = tk.Button(frame, text="Submit", bg="brown", fg="white")
# submit.grid(row=6, column=0)
#
#
# window.mainloop()

if __name__ == '__main__':
    app = App()
    app.mainloop()