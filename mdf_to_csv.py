
from asammdf import MDF
import sys
import os

if len(sys.argv) != 2:
    print("Usage: python mdf_to_csv.py <mdf_file_path>")
    sys.exit(1)

mdf_file = sys.argv[1]

# Create the output filename by changing the extension to .csv
output_filename = os.path.splitext(mdf_file)[0] + ".csv"

try:
    with MDF(mdf_file) as mdf:
        mdf.export(fmt='csv', filename=output_filename)
        print(f"Successfully converted {mdf_file} to {output_filename}")
except Exception as e:
    print(f"Error converting MDF file: {e}")
