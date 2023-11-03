import os
from pathlib import Path


CONFIGURATION = {
    "font": ("Calibri", 11),
    "tax rates": 1.1,
    "n_building": 5,
    "n_drawing": 5,
    "n_items": 3,
    "n_invoice": 6,
    "working_dir": str(Path(os.getcwd()).parent)
}
CONFIGURATION["database_dir"] = os.path.join(os.path.join(Path(CONFIGURATION['working_dir']), "app"), "database")
CONFIGURATION["resource_dir"] = os.path.join(os.path.join(Path(CONFIGURATION['working_dir']), "app"), "resource")