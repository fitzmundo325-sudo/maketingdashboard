"""Import COINBANK DATABASE.xlsx into the 'Coinbank Release Database' category."""
import sys

from excel_import import run_import

if __name__ == "__main__":
    run_import(r"c:\Users\325\Downloads\COINBANK DATABASE.xlsx", update=True)