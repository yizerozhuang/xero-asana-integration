import tkinter as tk
import pdfkit
from function.utility import *
from function.project_info_page import ProjectInfoPage
from function.fee_proposal_page import FeeProposalPage
import psycopg2
import os
class App(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        self.title("Data Entry Form")
        self.font = ("Calibri", 11)
        self.data = {}
        # the main frame
        self.main_frame = tk.Frame(self, width=600)
        self.main_frame.pack(fill=tk.BOTH, expand=1)
        self.conn = psycopg2.connect(
            host="127.0.0.1",
            database="postgres",
            user="postgres",
            password="Zero0929"
        )
        self.cur = self.conn.cursor()

        self.utility_frame()
        self.change_page_frame()
        self.main_context_frame()

        self.show_frame(self.page_info_page)



    def utility_frame(self):
        # Utility Part
        utility_frame = tk.LabelFrame(self.main_frame, text="Utility", font=self.font)
        utility_frame.pack(side=tk.TOP)

        create_folder = tk.Button(utility_frame, text="Create new Folder", command=create_new_folder, bg="brown",
                                  fg="white", font=self.font)
        create_folder.grid(row=0, column=0)

        rename_folder = tk.Button(utility_frame, text="Rename folder", command=self.rename_new_folder, bg="brown",
                                  fg="white", font=self.font)
        rename_folder.grid(row=0, column=1)

        update_asana = tk.Button(utility_frame, text="Update Asana", bg="brown", fg="white", font=self.font)
        update_asana.grid(row=0, column=2)

        print_button = tk.Button(utility_frame, text="Print", bg="brown", command = self._print_pdf, fg="white", font=self.font)
        print_button.grid(row=0, column=3)

    def change_page_frame(self):
        #change page
        change_page_frame = tk.LabelFrame(self.main_frame, text="Change Page")
        change_page_frame.pack(side=tk.BOTTOM)

        proj_info_button = tk.Button(change_page_frame, text="Project Info", command=lambda: self.show_frame(self.page_info_page), bg="brown", fg="white", font=self.font)
        proj_info_button.grid(row=0, column=0)
        fee_proposal_button = tk.Button(change_page_frame, text="Fee Proposal", command=lambda: self.show_frame(self.fee_proposal_page), bg="brown", fg="white", font=self.font)
        fee_proposal_button.grid(row=0, column=1)

    def main_context_frame(self):
        #main frame page
        self.page_info_page = ProjectInfoPage(self.main_frame, self)
        self.fee_proposal_page = FeeProposalPage(self.main_frame, self)

    def show_frame(self, cont):
        self.page_info_page.pack_forget()
        self.fee_proposal_page.pack_forget()
        cont.pack(fill=tk.BOTH, expand=1)

    def update_service_type(self, var):
        self.fee_proposal_page.update_scope_frame(var)


    def update_client(self, *args):
        self.fee_proposal_page.update_client_name()

    def rename_new_folder(self):
        os.replace("./new folder", f"./{self.data['Project Information']['Quotation Number'].get()}-{self.data['Project Information']['Project Name'].get()}")


    def _print_pdf(self):
        options = {
            "--header-html": "resource/html/header.html",
            "--footer-html": "resource/html/footer.html"
        }
        with open("resource/html/cur_html.html", "w") as f:
            f.write(open("resource/html/sample_html.html").read().format(
                Client_Full_Name=self.data["Client"]["Client Full Name"].get(),
                Client_Company=self.data["Client"]["Client Company"].get(),
                Reference=self.data["Project Information"]["Quotation Number"].get(),
                Date=self.data["Fee Proposal Page"]["Date"].get(),
                Revision=self.data["Fee Proposal Page"]["Revision"].get(),
                Project_Name=self.data["Project Information"]["Project Name"].get(),
                Fee_Stage_Start=self.data["Fee Proposal Page"]["Time"]["Fee proposal stage Duration"][0].get(),
                Fee_Stage_End=self.data["Fee Proposal Page"]["Time"]["Fee proposal stage Duration"][1].get(),
                Design_Stage_Start=self.data["Fee Proposal Page"]["Time"]["Pre-design information collection Duration"][0].get(),
                Design_Stage_End=self.data["Fee Proposal Page"]["Time"]["Pre-design information collection Duration"][1].get(),
                Issue_Stage_Start=self.data["Fee Proposal Page"]["Time"]["Documentation and issue Duration"][0].get(),
                Issue_Stage_End=self.data["Fee Proposal Page"]["Time"]["Documentation and issue Duration"][1].get(),
                Scope_of_Work=self._format_scope(),
                Past_Project=self._format_experience(),
                Details=self._format_details()
                )
            )
        pdfkit.from_file("resource/html/cur_html.html", "output.pdf", options=options)

    def _format_details(self):
        res = ""
        for key, value in self.data["Fee Proposal Page"]["Details"].items():
            if value[0].get():
                ingst = int(float(value[1].get())*1.1) if value[1].get() != "" else 0
                res+=f"""
                 </tr>
                  <tr height=20 style='mso-height-source:userset;height:15.0pt'>
                  <td height=20 class=xl1525308 style='height:15.0pt'></td>
                  <td colspan=4 class=xl11625308 style='border-right:.5pt solid black'>{key} design and documentation</td>
                  <td class=xl9625308 align=right style='border-top:none;border-left:none'>{value[1].get()}</td>
                  <td class=xl9625308 align=right style='border-top:none;border-left:none'>${ingst}</td>
                  <td class=xl8825308></td>
                  <td class=xl7325308></td>
                 </tr>
                """
        return res
    def _format_scope(self):
        data = self._return_scope()
        res = ""
        i = 1
        for service in data.keys():
            for extra in data[service].keys():
                i+=1
                res+=f"""
                <tr height=20 style='mso-height-source:userset;height:15.0pt'>
                <td height=20 class=xl8625308 style='height:15.0pt'>2.{i}</td>
                <td class=xl6925308 colspan=3>{service} - {extra}</td>
                <td class=xl7425308></td>
                <td class=xl7425308></td>
                <td class=xl7425308></td>
                <td class=xl1525308></td>
                <td class=xl8625308></td>
                </tr>
                """
                for context in data[service][extra]:
                    res+=f"""
                     <tr height=20 style='mso-height-source:userset;height:15.0pt'>
                      <td height=20 class=xl6825308 style='height:15.0pt'>•</td>
                      <td class=xl1525308 colspan=4>{context}</td>
                      <td class=xl1525308></td>
                      <td class=xl6825308></td>
                      <td class=xl1525308></td>
                      <td class=xl6825308></td>
                     </tr>
                    """
        return res


    def _format_experience(self):
        res = ""
        project_type = self.data["Project Information"]["Project Type"].get()
        self.cur.execute(
            f"""
                SELECT *
                FROM past_project
                WHERE project_type='{project_type}'
            """
        )
        past_project = self.cur.fetchall()
        for project in past_project:
            res += f"""
             <tr height=20 style='mso-height-source:userset;height:15.0pt'>
              <td height=20 class=xl6825308 style='height:15.0pt'>•</td>
              <td class=xl1525308 colspan=4>{project[1]}</td>
              <td class=xl1525308></td>
              <td class=xl6825308></td>
              <td class=xl1525308></td>
              <td class=xl6825308></td>
             </tr>
            """
        return res

    def _return_scope(self):
        res = dict()
        for service_name, service_context in self.data["Fee Proposal Page"]["Scope of Work"].items():
            if service_context["on"]:
                res[service_name] = dict()
                for extra in ["Extend", "Exclusion", "Deliverables"]:
                    res[service_name][extra] = []
                    for context in self.data["Fee Proposal Page"]["Scope of Work"][service_name][extra]:
                        if context[0].get():
                            res[service_name][extra].append(context[1].get())
        return res

