import os
from pathlib import Path


CONFIGURATION = {
    "font": ["Calibri", 11],
    "tax rates": 1.1,
    "n_building": 5,
    "n_drawing": 5,
    "n_items": 4,
    "n_invoice": 6,
    "n_variation": 4,
    "bridge_email": "bridge@pcen.com.au",
    "email_username": "bridge@pcen.com.au",
    "email_password": "PcE$yD2023",
    "imap_server": "mail.pcen.com.au",
    "smap_server": "mail.pcen.com.au",
    "smap_port": 587,
    "xero_bill_email": "bills.of4xmk.wt2xjjy1w2n5vceb@xerofiles.com"
}
CONFIGURATION["working_dir"] = str(Path(os.getcwd()).parent)
CONFIGURATION["database_dir"] = os.path.join(os.path.join(Path(CONFIGURATION['working_dir']), "app"), "database")
CONFIGURATION["resource_dir"] = os.path.join(os.path.join(Path(CONFIGURATION['working_dir']), "app"), "resource")
CONFIGURATION["recycle_bin_dir"] = os.path.join(os.path.join(Path(CONFIGURATION['working_dir']), "app"), "recycle_bin")