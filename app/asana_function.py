import asana
from tkinter import messagebox
from utility import remove_none, config_log

asana_configuration = asana.Configuration()
asana_configuration.access_token = '1/1205463377927189:4825d7e7924a9dd8dd44a9c826e45455'
asana_api_client = asana.ApiClient(asana_configuration)

project_api_instance = asana.ProjectsApi(asana_api_client)
task_api_instance = asana.TasksApi(asana_api_client)
custom_fields_api_instance = asana.CustomFieldsApi(asana_api_client)
template_api_instance = asana.TaskTemplatesApi(asana_api_client)
workspace_gid = '1205045058713243'

def clearn_response(response):
    response = response.to_dict()["data"]
    return remove_none(response)

def name_id_map(api_list):
    res = dict()
    for item in api_list:
        res[item["name"]] = item["gid"]
    return res

def update_asana(app, *args):
    data = app.data
    all_projects = clearn_response(project_api_instance.get_projects_for_workspace(workspace_gid))
    projects_id_map = name_id_map(all_projects)

    all_task = clearn_response(task_api_instance.get_tasks(project=projects_id_map[data["Project Info"]["Project"]["Project Type"].get()]))

    task_id_map = name_id_map(all_task)
    if data["Project Info"]["Project"]["Quotation Number"].get() + "-" + data["Project Info"]["Project"]["Project Name"].get() in task_id_map.keys():
        rewrite=messagebox.askyesno("Duplicate", "Existing Asana Project Found, Do you want to overwrite")
        if rewrite:
            task_id = task_id_map[data["Project Info"]["Project"]["Quotation Number"].get() + "-" + data["Project Info"]["Project"]["Project Name"].get()]
        else:
            return
    else:
        current_project_template = clearn_response(template_api_instance.get_task_templates(
            project=projects_id_map[data["Project Info"]["Project"]["Project Type"].get()]))

        template_id_map = name_id_map(current_project_template)
        api_respond = clearn_response(template_api_instance.instantiate_task(template_id_map["P:\\300000-XXXX"]))
        task_id = api_respond["new_task"]["gid"]

    all_custom_fields = clearn_response(task_api_instance.get_task(task_id))["custom_fields"]
    custom_field_id_map = name_id_map(all_custom_fields)

    status_field = clearn_response(custom_fields_api_instance.get_custom_field(custom_field_id_map["Status"]))
    status_id_map = name_id_map(status_field["enum_options"])

    service_filed = clearn_response(custom_fields_api_instance.get_custom_field(custom_field_id_map["Services"]))
    service_id_map = name_id_map(service_filed["enum_options"])

    contact_field = clearn_response(custom_fields_api_instance.get_custom_field(custom_field_id_map["Contact Type"]))
    contact_id_map = name_id_map(contact_field["enum_options"])

    body = asana.TasksTaskGidBody(
        {
            "name": data["Project Info"]["Project"]["Quotation Number"].get() + "-" + data["Project Info"]["Project"]["Project Name"].get(),
            "custom_fields": {
                custom_field_id_map["Status"]: status_id_map["Fee Proposal"],
                custom_field_id_map["Services"]: [service_id_map[key] for key, value in
                                                  data["Project Info"]["Project"]["Service Type"].items() if value["Include"].get() and not value["Archive"].get()],
                custom_field_id_map["Shop name"]: data["Project Info"]["Project"]["Shop Name"].get(),
                custom_field_id_map["Apt/Room/Area"]: data["Project Info"]["Building Features"]["Total Area"].get() + "m2",
                custom_field_id_map["Feature/Notes"]: data["Project Info"]["Building Features"]["Feature/Notes"].get(),
                custom_field_id_map["Client"]: data["Project Info"]["Client"]["Client Full Name"].get(),
                custom_field_id_map["Main contact"]: data["Project Info"]["Main Contact"]["Main Contact Full Name"].get(),
                custom_field_id_map["Contact Type"]: contact_id_map[
                    data["Project Info"]["Main Contact"]["Main Contact Contact Type"].get()]
            }
        }
    )
    task_api_instance.update_task(task_gid=task_id, body=body)
    messagebox.showinfo("Success", "Update/Create Asana Success")
    app.log.log_update_asana(app)
    config_log(app)

def rename_asana_project(app, old_folder, *args):
    data = app.data
    all_projects = clearn_response(project_api_instance.get_projects_for_workspace(workspace_gid))
    projects_id_map = name_id_map(all_projects)

    all_task = clearn_response(task_api_instance.get_tasks(project=projects_id_map[data["Project Info"]["Project"]["Project Type"].get()]))



    task_id_map = name_id_map(all_task)
    if old_folder in task_id_map.keys():
        task_id = task_id_map[old_folder]

        body = asana.TasksTaskGidBody(
            {
                "name": data["Project Info"]["Project"]["Quotation Number"].get() + "-" +
                        data["Project Info"]["Project"]["Project Name"].get()
            }
        )
        task_api_instance.update_task(task_gid=task_id, body=body)
        return True
    return False

def change_asana_quotation(app, new_quotation):
    data = app.data
    all_projects = clearn_response(project_api_instance.get_projects_for_workspace(workspace_gid))
    projects_id_map = name_id_map(all_projects)
    all_task = clearn_response(task_api_instance.get_tasks(project=projects_id_map[data["Project Info"]["Project"]["Project Type"].get()]))
    task_id_map = name_id_map(all_task)
    # try:
    old_name = data["Project Info"]["Project"]["Quotation Number"].get() + "-" + data["Project Info"]["Project"]["Project Name"].get()
    task_id = task_id_map[old_name]

    all_custom_fields = clearn_response(task_api_instance.get_task(task_id))["custom_fields"]
    custom_field_id_map = name_id_map(all_custom_fields)

    status_field = clearn_response(custom_fields_api_instance.get_custom_field(custom_field_id_map["Status"]))
    status_id_map = name_id_map(status_field["enum_options"])

    body = asana.TasksTaskGidBody(
        {
            "name": new_quotation + "-" + data["Project Info"]["Project"]["Project Name"].get(),
            "custom_fields": {custom_field_id_map["Status"]: status_id_map["Design"]}
        }
    )
    task_api_instance.update_task(task_gid=task_id, body=body)
    # except KeyError as e:
    #     print(e)
    #     messagebox.showerror("Error", "Cannot found the asana project, please contact the admin")

def update_asana_invoices(app):
    data = app.data

    all_projects = clearn_response(project_api_instance.get_projects_for_workspace(workspace_gid))
    projects_id_map = name_id_map(all_projects)
    all_task = clearn_response(task_api_instance.get_tasks(project=projects_id_map[data["Project Info"]["Project"]["Project Type"].get()]))
    task_id_map = name_id_map(all_task)
    task_id = task_id_map[app.data["Project Info"]["Project"]["Quotation Number"].get()+"-"+app.data["Project Info"]["Project"]["Project Name"].get()]

    all_subtask = clearn_response(task_api_instance.get_subtasks_for_task(task_id))
    subtask_id_map = name_id_map(all_subtask)

    for key, value in subtask_id_map.items():
        if key.startswith("INV "):
            inv_id = value
            break

    # inv_id = subtask_id_map["INV 3xxxxx, 1 of 1, refer checklist inside."]

    invoices = {
        "INV1": [],
        "INV2": [],
        "INV3": [],
        "INV4": [],
        "INV5": [],
        "INV6": []
    }

    for inv in ["INV1", "INV2", "INV3", "INV4", "INV5", "INV6"]:
        for key, service in data["Invoices"]["Details"].items():
            if service["Expand"].get():
                for i in range(app.conf["n_items"]):
                    if service["Content"][i]["Number"].get() == inv:
                        invoices[inv].append(
                            {
                                "Service": service["Content"][i]["Service"].get(),
                                "Fee": service["Content"][i]["Fee"].get()
                            }
                        )
            else:
                if service["Number"].get() == inv:
                    invoices[inv].append(
                        {
                            "Service": service["Service"].get(),
                            "Fee": service["Fee"].get()
                        }
                    )
    invoices_list = []
    for key, value in invoices.items():
        if len(value) != 0:
            invoices_list.append(value)

    all_custom_fields = clearn_response(task_api_instance.get_task(inv_id))["custom_fields"]
    custom_field_id_map = name_id_map(all_custom_fields)

    status_field = clearn_response(custom_fields_api_instance.get_custom_field(custom_field_id_map["Invoice Status"]))
    status_id_map = name_id_map(status_field["enum_options"])

    body = asana.TasksTaskGidBody(
        {
            "name": "INV " + data["Financial Panel"]["Invoice Details"]["INV1"]["Number"].get(),
            "custom_fields":{
                custom_field_id_map["Invoice Status"]: status_id_map["Draft"],
                custom_field_id_map["Net"]: str(sum([float(item["Fee"]) for item in invoices["INV1"]]))
            }
        }
    )
    task_api_instance.update_task(task_gid=inv_id, body=body)
    for i in range(len(invoices_list)-1):
        name = "INV " + data["Financial Panel"]["Invoice Details"][f"INV{i+2}"]["Number"].get() if len(data["Financial Panel"]["Invoice Details"][f"INV{i+2}"]["Number"].get()) !=0 else "INV 3xxxxx"
        body = asana.TasksTaskGidBody(
            {
                "name": name,
                "include":["parent"]
            }
        )
        task_api_instance.duplicate_task(body=body, task_gid=inv_id)


