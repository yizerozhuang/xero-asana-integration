import os
import shutil

scan_folder_dir = "S:\\01.Expense invoice"
destination_folder = "S:\\20240402-bill_classification"

all_folders = ["2022"+str(i+1).rjust(2, "0") for i in range(12)] + ["2023"+str(i+1).rjust(2, "0") for i in range(12)]
print(all_folders)

for folder in all_folders:
    sr = os.path.join(scan_folder_dir, folder)
    for root_file in os.listdir(sr):
        if os.path.isdir(os.path.join(sr, root_file)):
            for file in os.listdir(os.path.join(sr, root_file)):
                if file == 'Thumbs.db':
                    continue
                if not os.path.exists(os.path.join(destination_folder, file)):
                    shutil.copy(os.path.join(sr, root_file, file), os.path.join(destination_folder, file))
                else:
                    print()
        else:
            if root_file == 'Thumbs.db':
                continue
            if not os.path.exists(os.path.join(destination_folder, root_file)):
                shutil.copy(os.path.join(sr, root_file), os.path.join(destination_folder, root_file))
            else:
                print()