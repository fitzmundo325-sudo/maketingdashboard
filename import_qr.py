"""Import Coin Bank QR System.xlsx into the 'Coinbank QR Code System' category."""
import sys

from excel_import import run_import

if __name__ == "__main__":
    run_import(r"c:\Users\325\Downloads\Coin Bank QR System.xlsx", backfill="--backfill" in sys.argv)