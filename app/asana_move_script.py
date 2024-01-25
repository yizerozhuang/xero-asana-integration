import asana
from asana_function import clearn_response
asana_configuration = asana.Configuration()
asana_configuration.access_token = '2/1203283895754383/1206354773081941:c116d68430be7b2832bf5d7ea2a0a415'
asana_api_client = asana.ApiClient(asana_configuration)
project_api_instance = asana.ProjectsApi(asana_api_client)
task_api_instance = asana.TasksApi(asana_api_client)
workspace_gid = '1198726743417674'
old_project_gid = "1206098784008530"
new_project_gid = "1206394969582055"

all_new_sub_task = clearn_response(task_api_instance.get_subtasks_for_task(new_project_gid))


for i, e in enumerate(all_new_sub_task):
    new_task_gid = e["gid"]
    new_task = clearn_response(task_api_instance.get_task(new_task_gid))
    print(f"Processing New Task {i}")
    if new_task["completed"]:
        body = asana.TasksTaskGidBody(
            {
                "parent": old_project_gid,
                "insert_before": None
            }
        )
        task_api_instance.set_parent_for_task(body=body, task_gid=new_task_gid)

all_old_sub_task = clearn_response(task_api_instance.get_subtasks_for_task(old_project_gid))
for i, e in enumerate(all_old_sub_task):
    old_task_gid = e["gid"]
    old_task = clearn_response(task_api_instance.get_task(old_task_gid))
    print(f"Processing Old Task {i}")
    if not old_task["completed"]:
        body = asana.TasksTaskGidBody(
            {
                "parent": new_project_gid,
                "insert_before": None
            }
        )
        task_api_instance.set_parent_for_task(body=body, task_gid=old_task_gid)

