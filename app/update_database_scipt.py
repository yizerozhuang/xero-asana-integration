from config import CONFIGURATION as conf
import os
import json

database_dir = "./database"
# database_dir = conf["database_dir"]
scope_dir = os.path.join(database_dir, "scope_of_work.json")
scope = json.load(open(scope_dir))

stage_dir = os.path.join(database_dir, "general_scope_of_staging.json")
stage_json = json.load(open(stage_dir))

for dir in os.listdir(database_dir):
    if os.path.isdir(os.path.join(database_dir, dir)):
        data_dir = os.path.join(database_dir, dir, "data.json")
        data_json = json.load(open(data_dir))

        for service in data_json["Invoices"]["Details"].values():
            service["Fee"] = service.pop("Fee")
            service["in.GST"] = service.pop("in.GST")

        with open(data_dir, "w") as f:
            json_object = json.dumps(data_json, indent=4)
            f.write(json_object)


