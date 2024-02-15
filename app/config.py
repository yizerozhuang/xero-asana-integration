import os
from pathlib import Path


CONFIGURATION = {
    "font": ["Calibri", 11],
    "proposal_list": ["Mechanical Service", "Mechanical Review", "CFD Service", "Electrical Service", "Hydraulic Service", "Fire Service", "Miscellaneous"],
    "service_list":["Mechanical Service", "Mechanical Review", "CFD Service", "Electrical Service", "Hydraulic Service", "Fire Service", "Miscellaneous", "Installation"],
    "invoice_list": ["Mechanical Service", "Mechanical Review", "CFD Service", "Electrical Service", "Hydraulic Service", "Fire Service", "Miscellaneous", "Installation", "Variation"],
    "extra_list":["Extent", "Clarifications", "Deliverables"],
    "major_stage": ["Design Application", "Design Development", "Construction Documentation", "Construction Phase Service"],
    "tax rates": 1.1,
    "n_building": 5,
    "n_major_building": 20,
    "n_car_park": 8,
    "n_drawing": 12,
    "n_items": 4,
    "n_bills": 3,
    "n_invoice": 6,
    "n_variation": 4,
    "len_per_line": 90,
    "car_park_row": 6,
    "bridge_email": "bridge@pcen.com.au",
    "email_username": "bridge@pcen.com.au",
    "email_password": "PcE$yD2023",
    "imap_server": "mail.pcen.com.au",
    "smap_server": "mail.pcen.com.au",
    "smap_port": 587,
    "xero_client_id": "92582E6BA77A41F0B5076D3E5B442A24",
    "xero_client_secret": "YmhTPLEHqGhjYFOK0uPowcpVsgdLJ2ZKYD_PKq-rjGJVQIml",
    "xero_bill_email": "bills.vwkv1.b68g90zti0h38qcd@xerofiles.com",
    "engineer_user_list": ["Engineer1"]
    # "xero_bill_email":"bills.of4xmk.wt2xjjy1w2n5vceb@xerofiles.com"
}





CONFIGURATION["working_dir"] = str(Path(os.getcwd()).parent)
# CONFIGURATION["database_dir"] = "A:\\00-Bridge Database"
CONFIGURATION["database_dir"] = os.path.join(os.path.join(Path(CONFIGURATION['working_dir']), "app"), "database")
CONFIGURATION["accounting_dir"] = os.path.join(os.path.join(Path(CONFIGURATION['working_dir']), "app"), "accounting")

CONFIGURATION["bills_dir"] = "S:\\01.Expense invoice"
CONFIGURATION["remittances_dir"] = "S:\\01.Expense invoice"
CONFIGURATION["resource_dir"] = "T:\\00-Template-Do Not Modify\\00-Bridge template"
# CONFIGURATION["resource_dir"] = os.path.join(os.path.join(Path(CONFIGURATION['working_dir']), "app"), "resource")
CONFIGURATION["xero_access_token_dir"] = os.path.join(CONFIGURATION["resource_dir"], "txt", "xero_access_token.txt")
CONFIGURATION["xero_refresh_token_dir"] = os.path.join(CONFIGURATION["resource_dir"], "txt", "xero_refresh_token.txt")
CONFIGURATION["recycle_bin_dir"] = os.path.join(os.path.join(Path(CONFIGURATION['working_dir']), "app"), "recycle_bin")