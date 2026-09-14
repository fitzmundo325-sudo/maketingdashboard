"""Import COINBANK LIQUIDATION 2025.xlsx into the 'Coinbank Replenishment Database' project."""
from .excel_import import run_import

if __name__ == "__main__":
    run_import(r"c:\Users\325\Downloads\COINBANK LIQUIDATION 2025.xlsx", update=True)
