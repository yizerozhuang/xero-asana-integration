import tkinter as tk
from tkinter import ttk
from function.utility import *
from function.project_info_page import ProjectInfoPage
from function.fee_proposal_page import FeeProposalPage

class App(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        self.title("Data Entry Form")
        self.font = ("Calibri", 11)
        self.data = {}
        # the main frame
        main_frame = tk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=1)

        # Utility Part
        utility_frame = tk.LabelFrame(main_frame, text="Utility", font=self.font)
        utility_frame.pack(side=tk.TOP)

        create_folder = tk.Button(utility_frame, text="Create new Folder", command=create_new_folder, bg="brown",
                                  fg="white", font=self.font)
        create_folder.grid(row=0, column=0)

        update_asana = tk.Button(utility_frame, text="Update Asana", bg="brown", fg="white", font=self.font)
        update_asana.grid(row=0, column=1)

        print_button = tk.Button(utility_frame, text="Print", bg="brown", fg="white", font=self.font)
        print_button.grid(row=0, column=2)

        #change page
        change_page_frame = tk.LabelFrame(main_frame, text="Change Page")
        change_page_frame.pack(side=tk.BOTTOM)

        proj_info_button = tk.Button(change_page_frame, text="Project Info", command=lambda: self.show_frame(ProjectInfoPage), bg="brown", fg="white", font=self.font)
        proj_info_button.grid(row=0, column=0)
        fee_proposal_button = tk.Button(change_page_frame, text="Fee Proposal", command=lambda: self.show_frame(FeeProposalPage), bg="brown", fg="white", font=self.font)
        fee_proposal_button.grid(row=0, column=1)

        #main frame page
        self.main_canvas = tk.Canvas(main_frame)
        self.main_canvas.pack(side=tk.LEFT,  fill=tk.BOTH, expand=1)

        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.main_canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.main_canvas.config(yscrollcommand=scrollbar.set)
        self.main_canvas.bind("<Configure>", lambda e:self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")))
        self.main_canvas.bind("<MouseWheel>", self._on_mousewheel)
        main_context_frame = tk.LabelFrame(self.main_canvas, text="Main Context")
        x_0 = main_context_frame.winfo_screenwidth()/2
        y_0 = main_context_frame.winfo_screenheight()/2
        self.main_canvas.create_window((x_0, y_0), window=main_context_frame, anchor="center")
        self.frames = dict()
        for F in (ProjectInfoPage, FeeProposalPage):
            page = F(main_context_frame, self)
            self.frames[F] = page
            page.grid(row=0, column=0, sticky="nsew")
        self.show_frame(ProjectInfoPage)

    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()

    def update_scope(self, var):
        self.frames[FeeProposalPage].update_scope_frame(var)

    def _on_mousewheel(self, event):
        self.main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
