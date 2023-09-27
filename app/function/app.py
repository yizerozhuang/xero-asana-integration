import tkinter as tk
from function.utility import *
from function.project_info_page import ProjectInfoPage
from function.fee_proposal_page import FeeProposalPage

class App(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        self.title("Data Entry Form")
        frame = tk.Frame(self)
        frame.grid(row=0, column=0)

        # Utility Part
        utility_frame = tk.LabelFrame(frame, text="Utility")
        utility_frame.grid(row=0, column=0, padx=20)

        create_folder = tk.Button(utility_frame, text="Create new Folder", command=create_new_folder, bg="brown",
                                  fg="white")
        create_folder.grid(row=0, column=0)

        update_asana = tk.Button(utility_frame, text="Update Asana", bg="brown", fg="white")
        update_asana.grid(row=0, column=1)

        print_button = tk.Button(utility_frame, text="Print", bg="brown", fg="white")
        print_button.grid(row=0, column=2)
        #main page
        container = tk.Frame(self)
        container.grid(row=1, column=0)

        self.frames= dict()
        for F in (ProjectInfoPage, FeeProposalPage):
            page = F(container)
            self.frames[F] = page
            page.grid(row=0, column=0, sticky="nsew")

        self.show_frame(ProjectInfoPage)

        #change page
        change_page_frame = tk.Frame(self)
        change_page_frame.grid(row=2, column=0)

        proj_info_button = tk.Button(change_page_frame, text="Project Info", command=lambda: self.show_frame(ProjectInfoPage),bg="brown", fg="white")
        proj_info_button.grid(row=0, column=0)
        fee_proposal_button = tk.Button(change_page_frame, text="Fee Proposal", command=lambda: self.show_frame(FeeProposalPage),bg="brown", fg="white")
        fee_proposal_button.grid(row=0, column=1)
    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()

    def update_scope(self, var):
        self.frames[FeeProposalPage].update_scope_frame(var)


