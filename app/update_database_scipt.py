from config import CONFIGURATION as conf
import os
import json


database_dir = conf["database_dir"]
scope_dir = os.path.join(database_dir, "scope_of_work.json")
scope = json.load(open(scope_dir))
for dir in os.listdir(database_dir):
    if os.path.isdir(os.path.join(database_dir, dir)):
        data_dir = os.path.join(database_dir, dir, "data.json")
        data_json = json.load(open(data_dir))
        # for service in data_json["Fee Proposal"]["Scope"]["Minor"]:

        # data_json["Fee Proposal"]["Scope"]["Minor"]["Mechanical Service"]["Extend"] = [
        #     {
        #         "Include": True,
        #         "Item": item
        #     } for item in scope["Minor"]["Mechanical Service"]["Extend"]
        # ]
        # if "Type" in data_json["Project Info"]["Building Features"].keys():
        #     data_json["Project Info"]["Building Features"].pop("Type")
        # for service in data_json["Bills"]["Details"].values():
        #     if "Ori" in service:
        #         service["Origin"] = service.pop("Ori")

        with open(data_dir, "w") as f:
            json_object = json.dumps(data_json, indent=4)
            f.write(json_object)


