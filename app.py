import os
import asana
from flask import Flask, render_template, request, send_file
import openpyxl
def clear(data):
    data = data.to_dict()["data"]
    if type(data)==list:
        res = []
        for dic in data:
            cur = {}
            for key, value in dic.items():
                if not value is None:
                    cur[key]=value
            res.append(cur)
    else:
        res = dict()
        for key, value in data.items():
            if not value is None:
                res[key]=value
    return res

def custom_fields_finder(custom_fields, name):
    for field in custom_fields["custom_fields"]:
        if field["name"]==name:
            return field["display_value"]

# Configure OAuth2 access token for authorization: oauth2
configuration = asana.Configuration()
configuration.access_token = '1/1205463377927189:4825d7e7924a9dd8dd44a9c826e45455'
api_client = asana.ApiClient(configuration)
project_api_instance = asana.ProjectsApi(api_client)
task_api_instance = asana.TasksApi(api_client)
workspace_gid = '1205045058713243'

app = Flask(__name__)
Project_Type = None


@app.route("/")
def index():
    return render_template("index.html")

@app.route('/project')
def get_all_project():
    data = clear(project_api_instance.get_projects_for_workspace(workspace_gid))
    return render_template("project.html",data=data)

@app.route('/select_project', methods=["POST"])
def select_project():
    data = clear(task_api_instance.get_tasks(project=request.form["project_id"]))
    return render_template("select_project.html",data=data)

@app.route("/get_task", methods=["POST"])
def get_task():
    task_id = request.form["task_id"]
    task = clear(task_api_instance.get_task(task_id))
    path = "./Resource/xlsx/fee_proposal_template.xlsx"
    wb = openpyxl.load_workbook(path)
    project_info = wb.worksheets[0]

    project_info["B1"]=task["name"][7:]
    project_info["B3"]=task["name"][:6]
    project_info["B5"]=task["projects"][0]["name"]
    for i, service in enumerate(custom_fields_finder(task,"Service").split(", ")):
        project_info["B"+str(i+6)]=service
    invoices=[]
    sub_task = clear(task_api_instance.get_subtasks_for_task(task_id))
    for t in sub_task:
        if t["name"].startswith("INV"):
            invoices.append(t["gid"])
    for i in range(2,len(invoices)+1):
        target = wb["INV-1"]
        wb.copy_worksheet(target)
        copy_sheet = wb["INV-1 Copy"]
        copy_sheet.title = "INV-"+str(i)
    fee_proposal = wb["02b-Fee Proposal"]
    for i, inv in enumerate(invoices):
        invoice_sheet = wb["INV-"+str(i+1)]
        inv_task = clear(task_api_instance.get_task(inv))
        invoice_date = inv_task["due_on"]
        invoice_number = inv_task["name"].split(" ")[1]
        gross = float(custom_fields_finder(inv_task, "Gross"))
        note = custom_fields_finder(inv_task, "Notes")
        invoice_sheet["J4"]=invoice_date.strftime("%d-%b-%Y")
        invoice_sheet["J5"]=invoice_number
        invoice_sheet["A14"]=note
        invoice_sheet["I16"]=gross

        fee_proposal["B"+str(136+i)]=note
        fee_proposal["F"+str(136+i)]=gross

    path = "./Resource/xlsx/"+"Mechanical Fee Proposal for "+task["name"][7:]+".xlsx"
    wb.save(path)
    return send_file(path, as_attachment=True)

if __name__ == '__main__':
    task_id = "1205513500879876"
    body = asana.TasksTaskGidBody({"1205059451939777":"1205059451939779"})
    api_response = task_api_instance.update_task(body, task_id)
    app.run(port=4000)