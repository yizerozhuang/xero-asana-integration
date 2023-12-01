import asana
from tkinter import messagebox
from utility import remove_none, config_log, get_invoice_item

asana_configuration = asana.Configuration()
asana_configuration.access_token = '1/1205463377927189:4825d7e7924a9dd8dd44a9c826e45455'
asana_api_client = asana.ApiClient(asana_configuration)

project_api_instance = asana.ProjectsApi(asana_api_client)
task_api_instance = asana.TasksApi(asana_api_client)
custom_fields_api_instance = asana.CustomFieldsApi(asana_api_client)
custom_fields_setting_api_instance = asana.CustomFieldSettingsApi(asana_api_client)
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
        # rewrite=messagebox.askyesno("Duplicate", "Existing Asana Project Found, Do you want to overwrite")
        # if rewrite:
        task_id = task_id_map[data["Project Info"]["Project"]["Quotation Number"].get() + "-" + data["Project Info"]["Project"]["Project Name"].get()]
        # else:
        #     return
    else:
        current_project_template = clearn_response(template_api_instance.get_task_templates(
            project=projects_id_map[data["Project Info"]["Project"]["Project Type"].get()]))

        template_id_map = name_id_map(current_project_template)
        api_respond = clearn_response(template_api_instance.instantiate_task(template_id_map["P:\\300000-XXXX"]))
        task_id = api_respond["new_task"]["gid"]
        # task_id = api_respond["gid"]

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
                                                  data["Project Info"]["Project"]["Service Type"].items() if value["Include"].get()],
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
    # messagebox.showinfo("Success", "Update/Create Asana Success")
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


    custom_fields_setting = clearn_response(custom_fields_setting_api_instance.get_custom_field_settings_for_project(projects_id_map["Invoice status"]))
    all_custom_fields = [custom_field["custom_field"] for custom_field in custom_fields_setting]


    all_subtask = clearn_response(task_api_instance.get_subtasks_for_task(task_id))
    subtask_id_map = name_id_map(all_subtask)
    custom_field_id_map = name_id_map(all_custom_fields)

    status_field = clearn_response(
        custom_fields_api_instance.get_custom_field(custom_field_id_map["Invoice Status"]))
    status_id_map = name_id_map(status_field["enum_options"])

    asana_inv_list = {}
    for key, value in subtask_id_map.items():
        if key.startswith("INV "):
            asana_inv_list[key] = value

    invoice_item = {}
    for key, value in get_invoice_item(app).items():
        if len(value) != 0:
            invoice_item[key] = value



    invoice_status_templates = clearn_response(template_api_instance.get_task_templates(project=projects_id_map["Invoice status"]))
    template_id_map = name_id_map(invoice_status_templates)

    # invoice_template = clearn_response(template_api_instance.instantiate_task(template_id_map["INV Template"]))
    invoice_template_id = template_id_map["INV Template"]

    for i in range(len(invoice_item) - len(asana_inv_list)):
        body = asana.TasksTaskGidBody(
            {
                "name": f"INV 3xxxxx {len(asana_inv_list)+1 +i} of {len(invoice_item)}",
            }
        )
        response = template_api_instance.instantiate_task(task_template_gid=invoice_template_id, body=body).to_dict()
        new_inv_task_gid = response["data"]['new_task']["gid"]
        asana_inv_list[f"INV 3xxxxx {len(asana_inv_list) +i +1} of {len(invoice_item)}"] = new_inv_task_gid

        body = asana.TasksTaskGidBody(
            {
                "parent": task_id,
                "insert_before": None
            }
        )
        task_api_instance.set_parent_for_task(body=body, task_gid=new_inv_task_gid)
        # asana_inv_list[f"INV 3xxxxx {len(asana_inv_list) +i +1} of {len(invoice_item)}"] = response["data"]["gid"]

    # for task in subtask_id_map:
    #     if task.startswith("INV "):

    # all_custom_fields = clearn_response(project_api_instance.get_project(projects_id_map["Invoice status"]))["custom_field_settings"]


    for i, _ in enumerate(asana_inv_list.items()):
        try:
            key, value = _
            name = "INV " + data["Financial Panel"]["Invoice Details"][f"INV{i+1}"]["Number"].get() if len(data["Financial Panel"]["Invoice Details"][f"INV{i+1}"]["Number"].get())!= 0 else f"INV 3xxxxx {i+1} of {len(invoice_item)}"

            body = asana.TasksTaskGidBody(
                {
                    "name": name,
                    "custom_fields": {
                        custom_field_id_map["Invoice Status"]: status_id_map["Draft"],
                        custom_field_id_map["Net"]: str(sum([float(item["Fee"]) for item in invoice_item[f"INV{i+1}"]]))
                    }
                }
            )
            task_api_instance.update_task(task_gid=value, body=body)
        except KeyError:
            continue

    asana_bill_list = {}
    for key, value in subtask_id_map.items():
        if key.startswith("BIL "):
            asana_bill_list[key] = value

    bill_item = {}
    for key, value in data["Bills"]["Details"].items():
        for i in range(app.conf["n_items"]-1):
            if len(value["Content"][i]["Number"].get()) != 0:
                bill_item[value["Content"][i]["Number"].get()] = value["Content"][i]

    # bill_template = clearn_response(template_api_instance.instantiate_task(template_id_map["BIL Template"]))
    bill_template_id = template_id_map["BIL Template"]

    for key, value in bill_item.items():
        if not f"BIL {data['Project Info']['Project']['Quotation Number'].get() + key}" in asana_bill_list:
            body = asana.TasksTaskGidBody(
                {
                    "name": f"BIL {data['Project Info']['Project']['Quotation Number'].get() + key}"
                }
            )
            response = template_api_instance.instantiate_task(body=body, task_template_gid=bill_template_id).to_dict()
            new_bill_task_gid = response["data"]['new_task']["gid"]
            asana_bill_list[f"BIL {data['Project Info']['Project']['Quotation Number'].get() + key}"] = new_bill_task_gid
            body = asana.TasksTaskGidBody(
                {
                    "parent": task_id,
                    "insert_before": None
                }
            )
            task_api_instance.set_parent_for_task(body=body, task_gid=new_bill_task_gid)
            # asana_bill_list[f"BILL {data['Project Info']['Project']['Quotation Number'].get() + key}"] = response["data"]["gid"]

        bill_task_id = asana_bill_list[f"BIL {data['Project Info']['Project']['Quotation Number'].get()+key}"]


        body = asana.TasksTaskGidBody(
            {
                "parent": task_id,
                "custom_fields": {
                    custom_field_id_map["Invoice Status"]: status_id_map["Draft"],
                    custom_field_id_map["Net"]: str(value["Fee"].get()) if len(str(value["Fee"].get()))!=0 else "0"
                }
            }
        )
        task_api_instance.update_task(task_gid=bill_task_id, body=body)




    all_invoices_task = clearn_response(task_api_instance.get_tasks(project=projects_id_map["Invoice status"]))
    invoice_task_id_map = name_id_map(all_invoices_task)


    for inv in app.data["Financial Panel"]["Invoice Details"].values():
        if len(inv["Number"].get()) == 0:
            continue
        body = asana.TasksTaskGidBody(
            {
                "custom_fields": {
                    custom_field_id_map["Invoice Status"]: status_id_map[inv["State"].get()],
                }
            }
        )
        task_api_instance.update_task(task_gid=invoice_task_id_map["INV "+inv["Number"].get()], body=body)

    for bill in data["Bills"]["Details"].values():
        for item in bill["Content"]:
            if len(item["Number"].get()) != 0:
                body = asana.TasksTaskGidBody(
                    {
                        "custom_fields": {
                            custom_field_id_map["Invoice Status"]: status_id_map[item["State"].get()],
                        }
                    }
                )
                bill_number = app.data["Project Info"]["Project"]["Project Number"].get() + item["Number"].get()
                task_api_instance.update_task(task_gid=invoice_task_id_map["BIL " + bill_number], body=body)
        # if inv[-1].isdigit():
        # else:
        #     task_api_instance.update_task(task_gid=task_id_map["BILL "+inv], body=body)





    # for i in range(len(bill_item) - len(asana_bill_list)):
    #     body = asana.TasksTaskGidBody(
    #         {
    #             "name": f"BILL {data['Project Info']['Project']['Quotation Number'].get()+data['Project Info']['Project']['Quotation Number'].get()}",
    #             "include": ["parent", "dependencies"]
    #         }
    #     )
    #     response = task_api_instance.duplicate_task(body=body, task_gid=inv_id).to_dict()
    #     asana_bill_list[f"INV 3xxxxx {len(asana_inv_list) +i +1} of {len(invoice_item)}"] = response["data"]["gid"]
    # for i in range(len(invoices_list)-1):
    #     name = "INV " + data["Financial Panel"]["Invoice Details"][f"INV{i+2}"]["Number"].get() if len(data["Financial Panel"]["Invoice Details"][f"INV{i+2}"]["Number"].get()) !=0 else "INV 3xxxxx"
    #     body = asana.TasksTaskGidBody(
    #         {
    #             "name": name,
    #             "include": ["parent"]
    #         }
    #     )
    #     task_api_instance.duplicate_task(body=body, task_gid=inv_id)

# def update_invoice_status(app, inv_list):
#     data = app.data
#     inv_list = list(inv_list.keys())
#     all_projects = clearn_response(project_api_instance.get_projects_for_workspace(workspace_gid))
#     projects_id_map = name_id_map(all_projects)
#
#     all_task = clearn_response(task_api_instance.get_tasks(project=projects_id_map["Invoice status"]))
#     task_id_map = name_id_map(all_task)
#
#
#     for inv in inv_list:

