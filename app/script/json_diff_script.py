import json
import jsondiff

old_json_dir = "C:\\Users\\Admin\\Desktop\\test_back_up\\20240402\\database_backup\\mp.json"
current_json_dir = "P:\\app\\database\\mp.json"
report_dir = "A:\\01-Bridge Management Report"
different_json_dir = "A:\\01-Bridge Management Report\\different_json.json"


old_json = json.load(open(old_json_dir))
current_json = json.load(open(current_json_dir))

different = dict(jsondiff.diff(old_json, current_json))

# with open(different_json_dir, "w") as f:
#     json.dump(different, f, indent=4)
