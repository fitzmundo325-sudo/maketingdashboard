"""Shared Excel -> app mappings, used by the import scripts and the app's edit routes."""

FILES = {
    "COINBANK DATABASE.xlsx": {
        "sheets": {
            "Coinbank Database": {
                "header_row": 1,
                "name": "Recipient store",
                "ref": "Transaction No.",
                "created_by": "Prepared By",
                "status": "Status",
                "instructions": ["Date Transmitted", "Total Quantity"],
                "notes": "col7",
                "labels": {"col7": "Notes"},
            },
            "2025 Coinbank": {
                "header_row": 1,
                "name": "STORE NAME",
                "created_by": "MANAGER",
                "status": "Status",
                "instructions": ["CLUSTER", "# of Rooms", "SCHOOL NAME", "NET SALES"],
            },
            "2025 Project Reach": {
                "header_row": 1,
                "name": "STORE NAME",
                "created_by": "MANAGER",
                "status": "Status",
                "instructions": ["CLUSTER", "PARTICULARS", "SCHOOL NAME", "NET SALES"],
            },
            "2026 Coinbank - TARGETS": {
                "header_row": 1,
                "name": "STORE NAME",
                "created_by": "MANAGER",
                "status": "STATUS",
                "instructions": ["CLUSTER", "# of Rooms", "SCHOOL NAME", "NET SALES", "CUSTOMER CLASSIFICATION"],
            },
            "2026 INVENTORY": {
                "header_row": 1,
                "name": "STORE NAME",
                "created_by": "MANAGER",
                "status": "Status",
                "instructions": ["CLUSTER", "2025 # of Rooms", "PCS @ STORE END", "PCS FORWARDED TO HQ", "NET SALES"],
            },
        }
    },
    "Coin Bank QR System.xlsx": {
        "sheets": {
            "COIN BANK QR CODE": {
                "header_row": 0,
                "name": "///",
                "instructions": ["Cluster", "Store", "Condition "],
                "labels": {"///": "Serial Number"},
                "exclude": ["QR Image", "col6"],
            },
            "Cluster and Store": {
                "header_row": 0,
                "name": "Store",
                "instructions": ["Cluster", "PENDING", "ORIGINAL REQUEST ", "DONE", "DONE DELIVERY", "DONE PRINTING OF QR CODE"],
                "notes": "NOTE",
                "labels": {"DONE": "DONE COUNT"},
            },
            "Coinbank Monitoring ": {
                "header_row": 0,
                "name": "Store",
                "instructions": ["col1", "DONE DELIVERY", "DONE PRINTING OF QR CODE"],
                "notes": "col5",
                "labels": {"col1": "Cluster", "col5": "NOTE"},
            },
            "EXPENSES": {
                "header_row": 1,
                "name": "Particular",
                "ref": "No.",
                "created_by": "Received by:",
                "status": "Status",
                "instructions": ["Date", "Amount", "Change", "TOTAL MONEY LEFT"],
            },
        }
    },
    "Coin Bank Request.xlsx": {
        "sheets": {
            "DATABASE": {
                "header_row": 1,
                "name": "Store",
                "all_columns": True,
            },
            "Control Number": {
                "header_row": 1,
                "name": "Control Number",
                "status": "STATUS",
                "all_columns": True,
                "labels": {"col8": "Seq", "col9": "# Distributed", "col10": "# Returned", "col11": "# Lacking"},
            },
            "Coinbank Expenses": {
                "header_row": 2,
                "name": "Item",
                "all_columns": True,
                "labels": {"Google Drive Link:": "Google Drive Link", "Date_2": "Release Date"},
            },
            "Coin Bank Monitoring Form": {
                "header_row": 2,
                "name": "Store",
                "all_columns": True,
            },
        }
    },
    "COINBANK LIQUIDATION 2025.xlsx": {
        "sheets": {
            "NEW FORMAT": {
                "header_row": 4,
                "name": "Payee",
                "all_columns": True,
            },
            "format": {
                "header_row": 4,
                "name": "Payee",
                "all_columns": True,
            },
            "LIQUIDATION": {
                "header_row": 4,
                "name": "Payee",
                "all_columns": True,
            },
        }
    },
    "STORE VIEW DATABASE.xlsx": {
        "sheets": {
            "Store View Database": {
                "header_row": 0,
                "name": "Store",
                "ref": "col1",
                "status": "Delivery Status",
                "instructions": ["Cluster", "Deployment Status"],
                "notes": "col6",
                "exclude": ["Funds Collected as of August 15, 2025", "Funds Collected as of August 30, 2025"],
                "labels": {"col1": "Serial Number", "col6": "Note"},
            },
        }
    },
}

PARENT_CATEGORIES = {
    "COINBANK DATABASE.xlsx": "Coinbank Release Database",
    "Coin Bank QR System.xlsx": "Coinbank QR Code System",
    "Coin Bank Request.xlsx": "Coinbank 2024",
    "COINBANK LIQUIDATION 2025.xlsx": "Coinbank Replenishment Database",
    "STORE VIEW DATABASE.xlsx": "Coinbank Store View Database",
}


def mapping_for(sheet_name):
    for file_cfg in FILES.values():
        sheet = file_cfg["sheets"].get(sheet_name)
        if sheet:
            return sheet
    return None