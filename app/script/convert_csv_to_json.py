import csv
import json
import os

from config import CONFIGURATION as conf

database_dir = conf["database_dir"]
csv_dir = os.path.join(database_dir, "Bridge_MP_csv.csv")
manuel_assign_json = {}
with open(csv_dir, "r") as csv_file:
    csvFile = csv.reader(csv_file)
    for lines in csvFile:
        if len(lines[2]) != 0:
            manuel_assign_json[lines[3]] = lines[2]

with open(os.path.join(database_dir, "manuel_assign_quotation.json"), "w") as f:
    json_object = json.dumps(manuel_assign_json, indent=4)
    f.write(json_object)