"""Import Coin Bank Request.xlsx into the 'Coinbank 2024' project."""
import sys

from .excel_import import run_import

if __name__ == "__main__":
    run_import(r"c:\Users\325\Downloads\Coin Bank Request.xlsx", backfill="--backfill" in sys.argv)