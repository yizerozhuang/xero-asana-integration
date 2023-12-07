from tkinter import messagebox

class AppMessagebox:
    def __init__(self):
        pass
    def show_error(self, message):
        message = self.format_message(message)
        messagebox.showerror("Error", message)
    def show_update(self, message):
        message = self.format_message(message)
        messagebox.showinfo("Update", message)
    def file_info(self, command, old_folder, new_folder, others=None):
        if others is None:
            message = self.format_message(f"{command} from \n {old_folder} \n To {new_folder}")
        else:
            message = self.format_message(f"{command} from \n {old_folder} \n To {new_folder} \n {others}")
        messagebox.showinfo("Rename", message)
    def show_info(self, message):
        message = self.format_message(message)
        messagebox.showinfo("Info", message)
    def ask_yes_no(self, message):
        message = self.format_message(message)
        yes = messagebox.askyesno("Warning", message)
        return yes
    def file_ask_yes_no(self, command, old_folder, new_folder, update):
        message = self.format_message(f"{command} from \n {old_folder} \n To {new_folder} \n Do you want to Update {update}")
        yes = messagebox.askyesno("Warning", message)
        return yes

    def format_message(self, message):
        return "\n" + message + "\n" + "\n"