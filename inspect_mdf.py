
from asammdf import MDF
import sys

if len(sys.argv) != 2:
    print("Usage: python inspect_mdf.py <mdf_file_path>")
    sys.exit(1)

mdf_file = sys.argv[1]

try:
    with MDF(mdf_file) as mdf:
        print(mdf.info())
except Exception as e:
    print(f"Error reading MDF file: {e}")
