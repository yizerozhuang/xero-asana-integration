import os
import shutil
from tkinter import messagebox

def create_new_folder():
    try:
        os.mkdir("./new folder")
        os.mkdir("./new folder/External")
        os.mkdir("./new folder/Photos")
        os.mkdir("./new folder/Plot")
        os.mkdir("./new folder/SS")
        shutil.copyfile("resource/xlsx/Preliminary Calculation v2.5.xlsx",
                        "./new folder/Preliminary Calculation v2.5.xlsx")
    except:
        messagebox.showwarning(title="Error", message="The new folder already exist")


def update_asana():
    pass

