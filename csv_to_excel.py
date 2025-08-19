
import pandas as pd
import sys
import os

if len(sys.argv) != 2:
    print("Usage: python csv_to_excel.py <csv_file_path>")
    sys.exit(1)

csv_file = sys.argv[1]

# Create the output filename by changing the extension to .xlsx
output_filename = os.path.splitext(csv_file)[0] + ".xlsx"

try:
    df = pd.read_csv(csv_file)
    df.to_excel(output_filename, index=False)
    print(f"Successfully converted {csv_file} to {output_filename}")
except Exception as e:
    print(f"Error converting CSV file: {e}")
