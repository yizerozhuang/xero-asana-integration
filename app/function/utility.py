import asana
import win32com.client
import os

configuration = asana.Configuration()
configuration.access_token = '1/1205463377927189:4825d7e7924a9dd8dd44a9c826e45455'
api_client = asana.ApiClient(configuration)

project_api_instance = asana.ProjectsApi(api_client)
task_api_instance = asana.TasksApi(api_client)
custom_fields_api_instance = asana.CustomFieldsApi(api_client)
template_api_instance = asana.TaskTemplatesApi(api_client)
workspace_gid = '1205045058713243'


def remove_none(obj):
    if isinstance(obj, list):
        return [remove_none(x) for x in obj if x is not None]
    elif isinstance(obj, dict):
        return {k: remove_none(v) for k, v in obj.items() if v is not None}
    else:
        return obj


def clearn_response(response):
    response = response.to_dict()["data"]
    return remove_none(response)


def name_id_map(api_list):
    res = dict()
    for item in api_list:
        res[item["name"]] = item["gid"]
    return res


def update_asana(data, *args):
    all_projects = clearn_response(project_api_instance.get_projects_for_workspace(workspace_gid))
    projects_id_map = name_id_map(all_projects)

    current_project_template = clearn_response(template_api_instance.get_task_templates(
        project=projects_id_map[data["Project Information"]["Project Type"].get()]))
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
            "name": data["Project Information"]["Quotation Number"].get() + "-" + data["Project Information"][
                "Project Name"].get(),
            "custom_fields": {
                custom_field_id_map["Status"]: status_id_map["Fee Proposal"],
                custom_field_id_map["Services"]: [service_id_map[service._name] for service in
                                                 data["Project Information"]["Service Type"] if service.get()],
                custom_field_id_map["Shop name"]: data["Project Information"]["Project Name"].get(),
                custom_field_id_map["Apt/Room/Area"]: data["Building Features"]["Total Area"].get() + "m2",
                custom_field_id_map["Feature/Notes"]:data["Building Features"]["Feature/Notes"].get(),
                custom_field_id_map["Client"]:data["Client"]["Client Full Name"].get(),
                custom_field_id_map["Main contact"]:data["Main Contact"]["Main Contact Full Name"].get(),
                custom_field_id_map["Contact Type"]:contact_id_map[data["Main Contact"]["Main Contact Contact Type"].get()]
            }
        }
    )
    task_api_instance.update_task(task_gid=task_id, body=body)
    print(api_respond)

def excel_print_pdf(date, *args):
    pass

def email(data, *args):
    ol = win32com.client.Dispatch("Outlook.Application")
    olmailitem =0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = "Fee Proposal - " + data["Project Information"]["Project Name"].get()
    newmail.To = data["Client"]["Contact Email"].get()
    newmail.CC = "felix@pcen.com.au"
    newmail.Body=f"""
    {data["Client"]["Client Full Name"]}

    I hope this email finds you well. Please find the attached fee proposal to this email.

    If you have any questions or need more information regarding the proposal, please don't hesitate to reach out. Felix and I are happy to provide you with whatever information you need.

    I look forward to hearing from you soon.

    Cheers,

    Iza
    Administrative Assistant
    Premium Consulting Engineers Pty Ltd

    E : admin@pcen.com.au
    A : Suite 802, 299 Sussex Street, Sydney, NSW 2000
    W: www.pcen.com.au
    """
    newmail.Attachments.Add(os.getcwd()+"\output.pdf")
    newmail.Display()

