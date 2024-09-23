import asana
from config import CONFIGURATION as conf
from asana_function import clean_response
import json
import os

admin_users_list = ["admin@pcen.com.au", "felix@pcen.com.au", "jasper@forgebc.com.au", "joe@forgebc.com.au"]
director_users_list = ["felix@pcen.com.au"]
admin_user_gid = "1203283895754383"


asana_configuration = asana.Configuration()
asana_configuration.access_token = '2/1203283895754383/1206354773081941:c116d68430be7b2832bf5d7ea2a0a415'
asana_api_client = asana.ApiClient(asana_configuration)

user_api_instance = asana.UsersApi(asana_api_client)

workspace_gid = '1198726743417674'
database_dir = conf['database_dir']
opt_fields = ["email", "name"]
all_users = clean_response(user_api_instance.get_users(opt_fields=opt_fields))

user_json = {}

for user in all_users:
    name = user["name"]
    email = user["email"]

    user_json[email] = {
        "user_name":name,
        "password": "123456",
        "admin": True if email in admin_users_list else False,
        "asana_id": user["gid"],
        "supervisor_asana_id": user["gid"] if email in director_users_list else admin_user_gid
    }

user_json_dir = os.path.join(database_dir, "login.json")
with open(user_json_dir, "w") as f:
    json_object = json.dumps(user_json, indent=4)
    f.write(json_object)
print("Done")


