import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import os
from asammdf import MDF
import pandas as pd

class MDFConverterApp:
    def __init__(self, master):
        self.master = master
        master.title("MDF to XLSX Converter")

        self.default_load_location = Path.home() / "Downloads"
        self.mdf_file_path = ""

        self.label = tk.Label(master, text="Select an MDF file to convert.")
        self.label.pack(pady=10)

        self.select_button = tk.Button(master, text="Select MDF File", command=self.select_mdf_file)
        self.select_button.pack(pady=5)

        self.selected_file_label = tk.Label(master, text="")
        self.selected_file_label.pack(pady=5)

        self.convert_button = tk.Button(master, text="Convert to XLSX", command=self.convert_to_xlsx, state=tk.DISABLED)
        self.convert_button.pack(pady=10)

        self.status_label = tk.Label(master, text="")
        self.status_label.pack(pady=5)

    def select_mdf_file(self):
        self.mdf_file_path = filedialog.askopenfilename(
            initialdir=self.default_load_location,
            title="Select MDF File",
            filetypes=(("MDF files", "*.mdf"), ("All files", "*.*" ))
        )
        if self.mdf_file_path:
            self.selected_file_label.config(text=f"Selected file: {os.path.basename(self.mdf_file_path)}")
            self.convert_button.config(state=tk.NORMAL)
            self.status_label.config(text="")

    def convert_to_xlsx(self):
        if not self.mdf_file_path:
            messagebox.showerror("Error", "Please select an MDF file first.")
            return

        output_filename = Path(self.mdf_file_path).with_suffix(".xlsx")

        try:
            self.status_label.config(text="Converting...")
            self.master.update_idletasks()

            with MDF(self.mdf_file_path) as mdf:
                with pd.ExcelWriter(output_filename) as writer:
                    for i, group in enumerate(mdf.groups):
                        sheet_name = f"Channel Group {i}"
                        # The to_dataframe method can be slow for large data, this is a known issue.
                        df = mdf.to_dataframe(groups=[i])
                        df.to_excel(writer, sheet_name=sheet_name, index=False)

            self.status_label.config(text=f"Successfully converted to {output_filename.name}")
            messagebox.showinfo("Success", f"File converted successfully to:\n{output_filename}")

        except Exception as e:
            self.status_label.config(text="Conversion failed.")
            messagebox.showerror("Error", f"An error occurred during conversion:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MDFConverterApp(root)
    root.geometry("400x250")
    root.mainloop()