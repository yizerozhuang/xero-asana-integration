import os
from pathlib import Path


CONFIGURATION = {
    "font": ["Calibri", 11],
    "proposal_list": ["Mechanical Service", "CFD Service", "Electrical Service", "Hydraulic Service", "Fire Service"],
    "service_list":["Mechanical Service", "CFD Service", "Electrical Service", "Hydraulic Service", "Fire Service", "Installation"],
    "invoice_list": ["Mechanical Service", "CFD Service", "Electrical Service", "Hydraulic Service", "Fire Service", "Installation", "Variation"],
    "extra_list":["Extent", "Clarifications", "Deliverables"],
    "major_stage": ["Design Application", "Design Development", "Construction Documentation", "Construction Phase"],
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
    "xero_client_id": "92582E6BA77A41F0B5076D3E5B442A24",
    "xero_client_secret": "YmhTPLEHqGhjYFOK0uPowcpVsgdLJ2ZKYD_PKq-rjGJVQIml",
    "xero_bill_email": "bills.vwkv1.b68g90zti0h38qcd@xerofiles.com"
}




CONFIGURATION["working_dir"] = str(Path(os.getcwd()).parent)
# CONFIGURATION["database_dir"] = os.path.join(os.path.join(Path(CONFIGURATION['working_dir']), "app"), "database")
CONFIGURATION["database_dir"] = "A:\\00-Bridge Database"
CONFIGURATION["bills_address"] = os.path.join(os.path.join(Path(CONFIGURATION['working_dir']), "app"), "bills")
CONFIGURATION["remittances_address"] = os.path.join(os.path.join(Path(CONFIGURATION['working_dir']), "app"), "remittances")
CONFIGURATION["resource_dir"] = os.path.join(os.path.join(Path(CONFIGURATION['working_dir']), "app"), "resource")
CONFIGURATION["recycle_bin_dir"] = os.path.join(os.path.join(Path(CONFIGURATION['working_dir']), "app"), "recycle_bin")