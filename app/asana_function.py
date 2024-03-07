import asana
from utility import remove_none, config_log, get_invoice_item, isfloat

from tkinter import messagebox

import datetime
from datetime import date

asana_configuration = asana.Configuration()
asana_configuration.access_token = '2/1203283895754383/1206354773081941:c116d68430be7b2832bf5d7ea2a0a415'
asana_api_client = asana.ApiClient(asana_configuration)

project_api_instance = asana.ProjectsApi(asana_api_client)
task_api_instance = asana.TasksApi(asana_api_client)
custom_fields_api_instance = asana.CustomFieldsApi(asana_api_client)
custom_fields_setting_api_instance = asana.CustomFieldSettingsApi(asana_api_client)
template_api_instance = asana.TaskTemplatesApi(asana_api_client)
workspace_gid = '1198726743417674'
user_gid_map = {
    "Admin": "1203283895754383",
    "Felix": "1198835648677067",
    "Engineer1": "1203396624150040"
}
#

def clean_response(response):
    response = response.to_dict()["data"]
    # return remove_none(response)
    return response

def name_id_map(api_list):
    res = dict()
    for item in api_list:
        res[item["name"]] = item["gid"]
    return res

def flatter(inv_list):
    res = {**inv_list["Invoices"]["AUTHORISED"], **inv_list["Invoices"]["PAID"]}
    return res

def flatter_custom_fields(asana_task):
    if not "custom_fields" in asana_task.keys():
        return asana_task
    for field in asana_task["custom_fields"]:
        assert not field["name"] in asana_task.keys()
        asana_task[field["name"]] = field["display_value"]
    return asana_task


# def convert_email_content(email):
#     if email is None or len(email)==0:
#         return "<body> </body>"
#     return "<body>" + email.replace("<", "").replace(">", "") + "</body>"
def update_asana(app, *args):
    data = app.data
    if len(app.data["Asana_id"].get()) == 0:
        all_projects = clean_response(project_api_instance.get_projects_for_workspace(workspace_gid))
        projects_id_map = name_id_map(all_projects)
        current_project_template = clean_response(template_api_instance.get_task_templates(
            project=projects_id_map[data["Project Info"]["Project"]["Project Type"].get()]))

        template_id_map = name_id_map(current_project_template)
        api_respond = clean_response(template_api_instance.instantiate_task(template_id_map["P:\\300000-XXXX"]))
        new_task_gid = api_respond["new_task"]["gid"]

        body = asana.TaskGidAddProjectBody(
            {
                "project": projects_id_map["MP"]
            }
        )
        task_api_instance.add_project_for_task(task_gid=new_task_gid, body=body)
        data["Asana_id"].set(new_task_gid)
        asana_task = clean_response(task_api_instance.get_task(data["Asana_id"].get()))
        data["Asana_url"].set(asana_task["permalink_url"])
        project_task = clean_response(task_api_instance.get_task(data["Asana_id"].get()))
    else:
    #     try:
        project_task = clean_response(task_api_instance.get_task(data["Asana_id"].get()))
    #     except Exception as e:
    #         print(f"Unable to Found Project Asana id {data['Asana_id'].get()}, Creating New Asana Project")
    #         all_projects = clean_response(project_api_instance.get_projects_for_workspace(workspace_gid))
    #         projects_id_map = name_id_map(all_projects)
    #         current_project_template = clean_response(template_api_instance.get_task_templates(
    #             project=projects_id_map[data["Project Info"]["Project"]["Project Type"].get()]))
    #
    #         template_id_map = name_id_map(current_project_template)
    #         api_respond = clean_response(template_api_instance.instantiate_task(template_id_map["P:\\300000-XXXX"]))
    #         new_task_gid = api_respond["new_task"]["gid"]
    #         body = asana.TaskGidAddProjectBody(
    #             {
    #                 "project": projects_id_map["MP"]
    #             }
    #         )
    #         task_api_instance.add_project_for_task(task_gid=new_task_gid, body=body)
    #         data["Asana_id"].set(new_task_gid)
    #         asana_task = clean_response(task_api_instance.get_task(data["Asana_id"].get()))
    #         data["Asana_url"].set(asana_task["permalink_url"])
    #     project_task = clean_response(task_api_instance.get_task(data["Asana_id"].get()))
        # all_projects = clean_response(project_api_instance.get_projects_for_workspace(workspace_gid))
        # projects_id_map = name_id_map(all_projects)
        # current_project_template = clean_response(template_api_instance.get_task_templates(
        #     project=projects_id_map[data["Project Info"]["Project"]["Project Type"].get()]))
        #
        # template_id_map = name_id_map(current_project_template)
        # api_respond = clean_response(template_api_instance.instantiate_task(template_id_map["P:\\300000-XXXX"]))
        # new_task_gid = api_respond["new_task"]["gid"]
        # body = asana.TaskGidAddProjectBody(
        #     {
        #         "project": projects_id_map["MP"]
        #     }
        # )
        # task_api_instance.add_project_for_task(task_gid=new_task_gid, body=body)
        # data["Asana_id"].set(new_task_gid)
        # project_task = clean_response(task_api_instance.get_task(data["Asana_id"].get()))
        # data["Asana_url"].set(project_task["permalink_url"])

    task_id = data["Asana_id"].get()

    all_custom_fields = clean_response(task_api_instance.get_task(task_id))["custom_fields"]
    custom_field_id_map = name_id_map(all_custom_fields)

    status_field = clean_response(custom_fields_api_instance.get_custom_field(custom_field_id_map["Status"]))
    status_id_map = name_id_map(status_field["enum_options"])

    service_filed = clean_response(custom_fields_api_instance.get_custom_field(custom_field_id_map["Services"]))
    service_id_map = name_id_map(service_filed["enum_options"])

    contact_field = clean_response(custom_fields_api_instance.get_custom_field(custom_field_id_map["Contact Type"]))
    contact_id_map = name_id_map(contact_field["enum_options"])

    if len(data["Project Info"]["Project"]["Project Number"].get()) != 0:
        name = "P:\\"+data["Project Info"]["Project"]["Project Number"].get() + "-" + data["Project Info"]["Project"]["Project Name"].get()
    else:
        name = "P:\\"+data["Project Info"]["Project"]["Quotation Number"].get() + "-" + data["Project Info"]["Project"]["Project Name"].get()

    status = None
    for custom_field in project_task["custom_fields"]:
        if custom_field["name"] == "Status":
            status = custom_field["display_value"]

    if not status in ["Pending", "DWG drawings", "Done", "Installation", "Construction Phase"]:
        if data["State"]["Quote Unsuccessful"].get():
            status = "Quote Unsuccessful"
        elif data["State"]["Fee Accepted"].get():
            status = "Design"
        else:
            status = "Fee Proposal"
    data["State"]["Asana State"].set(status)

    client_name_list = []
    if len(data["Project Info"]["Client"]["Company"].get()) != 0:
        client_name_list.append(data["Project Info"]["Client"]["Company"].get())
    if len(data["Project Info"]["Client"]["Full Name"].get()) != 0:
        client_name_list.append(data["Project Info"]["Client"]["Full Name"].get())
    client_name = "-".join(client_name_list)

    main_contact_list = []
    if len(data["Project Info"]["Main Contact"]["Company"].get()) !=0:
        main_contact_list.append(data["Project Info"]["Main Contact"]["Company"].get())
    if len(data["Project Info"]["Main Contact"]["Full Name"].get()) != 0:
        main_contact_list.append(data["Project Info"]["Main Contact"]["Full Name"].get())
    main_contact_name = "-".join(main_contact_list)


    asana_update_body = {
        "name": name,
        "custom_fields": {
            custom_field_id_map["Status"]: status_id_map[status],
            custom_field_id_map["Services"]: [service_id_map[key] for key, value in
                                              data["Project Info"]["Project"]["Service Type"].items() if value["Include"].get()],
            custom_field_id_map["Shop name"]: data["Project Info"]["Project"]["Shop Name"].get(),
            custom_field_id_map["Apt/Room/Area"]: data["Project Info"]["Building Features"]["Apt"].get(),
            custom_field_id_map["Basement/Car Spots"]: data["Project Info"]["Building Features"]["Basement"].get(),
            custom_field_id_map["Feature/Notes"]: data["Project Info"]["Building Features"]["Feature"].get(),
            custom_field_id_map["Client"]: client_name,
            custom_field_id_map["Main Contact"]: main_contact_name,
            custom_field_id_map["Contact Type"]: contact_id_map[data["Project Info"]["Main Contact"]["Contact Type"].get()],
            custom_field_id_map["Fee ExGST"]:float(data["Invoices"]["Fee"].get()) if isfloat(data["Invoices"]["Fee"].get()) and len(data["Invoices"]["Fee"].get())!=0 else 0,
            custom_field_id_map["Total Paid ExGST"]: float(data["Invoices"]["Paid Fee"].get())if isfloat(data["Invoices"]["Paid Fee"].get()) and len(data["Invoices"]["Paid Fee"].get())!=0 else 0,
            custom_field_id_map["Overdue Amount"]: float(data["Invoices"]["Over Due Fee"].get()) if isfloat(
                data["Invoices"]["Over Due Fee"].get()) and len(data["Invoices"]["Over Due Fee"].get()) != 0 else 0
        }
    }
    # messagebox.showinfo("Success", "Update/Create Asana Success")

    sub_tasks = clean_response(task_api_instance.get_subtasks_for_task(task_id))
    if not data["State"]["Fee Accepted"].get():
        asana_update_body["notes"] = data["Email_Content"].get()
        body = asana.TasksTaskGidBody(asana_update_body)
        task_api_instance.update_task(task_gid=task_id, body=body)

        body_dict = {}
        first_task_gid = sub_tasks[0]["gid"]
        first_task = clean_response(task_api_instance.get_task(first_task_gid))

        if first_task["assignee"] is None:
            body_dict["assignee"] = user_gid_map[app.user]
        if first_task["due_on"] is None:
            body_dict["due_on"] = date.today().strftime("%Y-%m-%d")
        body = asana.TasksTaskGidBody(body_dict)
        task_api_instance.update_task(task_gid=first_task_gid, body=body)
    else:
        body = asana.TasksTaskGidBody(asana_update_body)
        task_api_instance.update_task(task_gid=task_id, body=body)
        notes = clean_response(task_api_instance.get_task(task_gid=task_id))["notes"]
        data["Email_Content"].set(notes)
        first_task_gid = sub_tasks[0]["gid"]
        # first_task = clearn_response(task_api_instance.get_task(first_task_gid))
        body = asana.TasksTaskGidBody(
            {
                "completed": True
            }
        )
        task_api_instance.update_task(task_gid=first_task_gid, body=body)
        for i in range(1, len(sub_tasks)):
            body_dict = {}
            task_gid = sub_tasks[i]["gid"]
            task = clean_response(task_api_instance.get_task(task_gid, opt_fields=["name", "assignee", "due_on", "is_rendered_as_separator"]))
            if task["is_rendered_as_separator"] or task["name"].startswith("INV") or task["name"].startswith("BIL") or task["name"].startswith("---"):
                continue
            if task["assignee"] is None:
                body_dict["assignee"] = user_gid_map[app.user]
            if task["due_on"] is None:
                body_dict["due_on"] = date.today().strftime("%Y-%m-%d")
            body = asana.TasksTaskGidBody(body_dict)
            task_api_instance.update_task(task_gid=task_gid, body=body)

    app.log.log_update_asana(app)
    config_log(app)


def update_asana_invoices(app, inv_list=None):
    data = app.data
    all_projects = clean_response(project_api_instance.get_projects_for_workspace(workspace_gid))
    projects_id_map = name_id_map(all_projects)

    task_id = app.data["Asana_id"].get()

    custom_fields_setting = clean_response(custom_fields_setting_api_instance.get_custom_field_settings_for_project(projects_id_map["Invoice status"]))
    all_custom_fields = [custom_field["custom_field"] for custom_field in custom_fields_setting]
    custom_field_id_map = name_id_map(all_custom_fields)

    status_field = clean_response(
        custom_fields_api_instance.get_custom_field(custom_field_id_map["Invoice status"]))
    status_id_map = name_id_map(status_field["enum_options"])


    invoice_item = get_invoice_item(app)
    invoice_status_templates = clean_response(
        template_api_instance.get_task_templates(project=projects_id_map["Invoice status"]))
    template_id_map = name_id_map(invoice_status_templates)
    invoice_template_id = template_id_map["INV Template"]

    if not inv_list is None:
        inv_list = flatter(inv_list)

    for i in range(6):
        if len(invoice_item[i])==0:
            continue
        if len(data["Invoices Number"][i]["Asana_id"].get()) == 0:
            body = asana.TasksTaskGidBody(
                {
                    "name": f"INV 4xxxxx",
                }
            )
            response = template_api_instance.instantiate_task(task_template_gid=invoice_template_id,
                                                              body=body).to_dict()
            new_inv_task_gid = response["data"]['new_task']["gid"]

            body = asana.TasksTaskGidBody(
                {
                    "parent": task_id,
                    "insert_before": None
                }
            )
            task_api_instance.set_parent_for_task(body=body, task_gid=new_inv_task_gid)
            data["Invoices Number"][i]["Asana_id"].set(new_inv_task_gid)
        else:

            try:
                task_api_instance.get_task(data["Invoices Number"][i]["Asana_id"].get())
            except Exception as e:
                print(e)
                messagebox.showerror("Error", f'Can not find the invoice with id{data["Invoices Number"][i]["Asana_id"].get()}')
                return
            #     print(f'Can not found Asana Invoice Task {data["Invoices Number"][i]["Asana_id"].get()} Creating New Invoice Task')
            #
            #     body = asana.TasksTaskGidBody(
            #         {
            #             "name": f"INV 4xxxxx",
            #         }
            #     )
            #     response = template_api_instance.instantiate_task(task_template_gid=invoice_template_id,
            #                                                       body=body).to_dict()
            #     new_inv_task_gid = response["data"]['new_task']["gid"]
            #
            #     body = asana.TasksTaskGidBody(
            #         {
            #             "parent": task_id,
            #             "insert_before": None
            #         }
            #     )
            #     task_api_instance.set_parent_for_task(body=body, task_gid=new_inv_task_gid)
            #     data["Invoices Number"][i]["Asana_id"].set(new_inv_task_gid)

        name = "INV " + data["Invoices Number"][i]["Number"].get() if len(data["Invoices Number"][i]["Number"].get())!= 0 else f"INV 4xxxxx"


        task_body = {
            "name": name,
            "custom_fields": {
                custom_field_id_map["Invoice status"]: status_id_map[data["Invoices Number"][i]["State"].get()],
                custom_field_id_map["Net"]: str(sum([float(item["Fee"]) for item in invoice_item[i]])),
                custom_field_id_map["Gross"]: str(sum([float(item["in.GST"]) for item in invoice_item[i]]))
            }
        }

        if data["Invoices Number"][i]["State"].get() == "Sent":
            invoice_task = clean_response(task_api_instance.get_task(data["Invoices Number"][i]["Asana_id"].get()))
            if invoice_task["due_on"] is None:
                task_body["due_on"] = (date.today() + datetime.timedelta(days=7)).strftime("%Y-%m-%d")
            if invoice_task["assignee"] is None:
                task_body["assignee"] = user_gid_map[app.user]
        if not inv_list is None:
            inv_number = data["Invoices Number"][i]["Number"].get()
            if inv_number in inv_list.keys():
                task_body["custom_fields"][custom_field_id_map["Payment Amount"]] = str(inv_list[inv_number]["payment_amount"])
                task_body["custom_fields"][custom_field_id_map["Payment Date"]] = {"date": str(inv_list[inv_number]["payment_date"])}
        body = asana.TasksTaskGidBody(task_body)
        try:
            task_api_instance.update_task(task_gid=data["Invoices Number"][i]["Asana_id"].get(), body=body)
        except Exception as e:
            print(e)


    custom_fields_setting = clean_response(
        custom_fields_setting_api_instance.get_custom_field_settings_for_project(projects_id_map["Bill status"]))
    all_custom_fields = [custom_field["custom_field"] for custom_field in custom_fields_setting]
    custom_field_id_map = name_id_map(all_custom_fields)

    status_field = clean_response(
        custom_fields_api_instance.get_custom_field(custom_field_id_map["Bill status"]))
    status_id_map = name_id_map(status_field["enum_options"])

    invoice_status_templates = clean_response(
        template_api_instance.get_task_templates(project=projects_id_map["Bill status"]))
    template_id_map = name_id_map(invoice_status_templates)
    bill_template_id = template_id_map["BIL 4"]

    for key, value in data["Bills"]["Details"].items():
        for i in range(app.conf["n_bills"]):
            if len(value["Content"][i]["Number"].get()) == 0:
                continue

            if len(value["Content"][i]["Asana_id"].get()) == 0:
                body = asana.TasksTaskGidBody(
                    {
                        "name": f"BIL {data['Project Info']['Project']['Project Number'].get() + value['Content'][i]['Number'].get()}"
                    }
                )
                response = template_api_instance.instantiate_task(body=body, task_template_gid=bill_template_id).to_dict()
                new_bill_task_gid = response["data"]['new_task']["gid"]
                body = asana.TasksTaskGidBody(
                    {
                        "parent": task_id,
                        "insert_before": None
                    }
                )
                task_api_instance.set_parent_for_task(body=body, task_gid=new_bill_task_gid)
                value["Content"][i]["Asana_id"].set(new_bill_task_gid)
                # asana_bill_list[f"BILL {data['Project Info']['Project']['Quotation Number'].get() + key}"] = response["data"]["gid"]
            else:
                try:
                    task_api_instance.get_task(value["Content"][i]["Asana_id"].get())
                except Exception as e:
                    print(f'Can not find the Asana Bill Task {value["Content"][i]["Asana_id"].get()}, Creating New Asana Bill Task')
                    body = asana.TasksTaskGidBody(
                        {
                            "name": f"BIL {data['Project Info']['Project']['Project Number'].get() + value['Content'][i]['Number'].get()}"
                        }
                    )
                    response = template_api_instance.instantiate_task(body=body,
                                                                      task_template_gid=bill_template_id).to_dict()
                    new_bill_task_gid = response["data"]['new_task']["gid"]
                    body = asana.TasksTaskGidBody(
                        {
                            "parent": task_id,
                            "insert_before": None
                        }
                    )
                    task_api_instance.set_parent_for_task(body=body, task_gid=new_bill_task_gid)
                    value["Content"][i]["Asana_id"].set(new_bill_task_gid)


            state = value["Content"][i]["State"].get()
            bill_task_id = value["Content"][i]["Asana_id"].get()
            body = asana.TasksTaskGidBody(
                {
                    "custom_fields": {
                        custom_field_id_map["Bill status"]: status_id_map[state],
                        custom_field_id_map["Amount Excl GST"]: str(value["Fee"].get()) if len(str(value["Fee"].get()))!=0 else "0",
                        custom_field_id_map["Amount Incl GST"]: str(value["in.GST"].get()) if len(str(value["in.GST"].get())) != 0 else "0"
                    }
                }
            )
            task_api_instance.update_task(task_gid=bill_task_id, body=body)
# def update_asana_invoices_from_xero(inv_list):




