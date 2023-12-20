import os
from pathlib import Path


CONFIGURATION = {
    "font": ["Calibri", 11],
    "proposal_list": ["Mechanical Service", "CFD Service", "Electrical Service", "Hydraulic Service", "Fire Service"],
    "service_list":["Mechanical Service", "CFD Service", "Electrical Service", "Hydraulic Service", "Fire Service", "Installation"],
    "invoice_list": ["Mechanical Service", "CFD Service", "Electrical Service", "Hydraulic Service", "Fire Service", "Installation", "Variation"],
    "extra_list":["Extend", "Clarifications", "Deliverables"],
    "sub_title": ["Design Development", "Construction Documents", "Construction Phase Service"],
    "tax rates": 1.1,
    "n_building": 5,
    "n_major_building": 8,
    "n_drawing": 5,
    "n_items": 4,
    "n_bills": 3,
    "n_invoice": 6,
    "n_variation": 4,
    "len_per_line": 90,
    "bridge_email": "bridge@pcen.com.au",
    "email_username": "bridge@pcen.com.au",
    "email_password": "PcE$yD2023",
    "imap_server": "mail.pcen.com.au",
    "smap_server": "mail.pcen.com.au",
    "smap_port": 587,
    "xero_client_id": "876EFEC2F1AC4729812A3B39152A2DD3",
    "xero_client_secret": "jtohs0Oqcoezje-bjYn8n9KaTa9hCm2taATzBIbS3RpaXmOl",
    "xero_bill_email": "bills.of4xmk.wt2xjjy1w2n5vceb@xerofiles.com"
}




CONFIGURATION["working_dir"] = str(Path(os.getcwd()).parent)
CONFIGURATION["database_dir"] = os.path.join(os.path.join(Path(CONFIGURATION['working_dir']), "app"), "database")
CONFIGURATION["resource_dir"] = os.path.join(os.path.join(Path(CONFIGURATION['working_dir']), "app"), "resource")
CONFIGURATION["recycle_bin_dir"] = os.path.join(os.path.join(Path(CONFIGURATION['working_dir']), "app"), "recycle_bin")