"""Import STORE VIEW DATABASE.xlsx into the 'Coinbank Store View Database' category."""
import sys

from excel_import import run_import

if __name__ == "__main__":
    run_import(r"c:\Users\325\Downloads\STORE VIEW DATABASE.xlsx", update=True)