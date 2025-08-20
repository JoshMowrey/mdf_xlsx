import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import os
import threading
from asammdf import MDF
import pandas as pd

class MDFConverterApp:
    def __init__(self, master):
        self.master = master
        master.title("MDF to XLSX Converter")

        self.style = ttk.Style(master)
        self.style.theme_use("clam")

        self.frame = ttk.Frame(master, padding="10")
        self.frame.pack(fill=tk.BOTH, expand=True)

        self.default_load_location = Path.home() / "Downloads"
        self.mdf_file_path = ""

        self.label = ttk.Label(self.frame, text="Select an MDF file to convert.")
        self.label.pack(pady=10)

        self.select_button = ttk.Button(self.frame, text="Select MDF File", command=self.select_mdf_file)
        self.select_button.pack(pady=5)

        self.selected_file_label = ttk.Label(self.frame, text="")
        self.selected_file_label.pack(pady=5)

        self.output_dir_label = ttk.Label(self.frame, text="Output directory:")
        self.output_dir_label.pack(pady=5)

        self.output_dir_entry = ttk.Entry(self.frame, width=50)
        self.output_dir_entry.pack(pady=5)

        self.select_output_dir_button = ttk.Button(self.frame, text="Select Output Directory", command=self.select_output_dir)
        self.select_output_dir_button.pack(pady=5)

        self.convert_button = ttk.Button(self.frame, text="Convert to XLSX", command=self.start_conversion_thread, state=tk.DISABLED)
        self.convert_button.pack(pady=10)

        self.progress_bar = ttk.Progressbar(self.frame, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.pack(pady=10)

        self.status_label = ttk.Label(self.frame, text="")
        self.status_label.pack(pady=5)

        self.status_bar = ttk.Label(self.master, text="Ready", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.create_menu()

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
            self.status_bar.config(text=f"Ready to convert {os.path.basename(self.mdf_file_path)}")

    def select_output_dir(self):
        output_dir = filedialog.askdirectory(
            initialdir=self.default_load_location,
            title="Select Output Directory"
        )
        if output_dir:
            self.output_dir_entry.delete(0, tk.END)
            self.output_dir_entry.insert(0, output_dir)

    def start_conversion_thread(self):
        self.convert_button.config(state=tk.DISABLED)
        self.progress_bar["value"] = 0
        self.status_label.config(text="Converting...")
        self.status_bar.config(text="Converting...")
        self.conversion_thread = threading.Thread(target=self.convert_to_xlsx)
        self.conversion_thread.start()
        self.master.after(100, self.check_conversion_thread)

    def check_conversion_thread(self):
        if self.conversion_thread.is_alive():
            self.master.after(100, self.check_conversion_thread)
        else:
            self.convert_button.config(state=tk.NORMAL)

    def convert_to_xlsx(self):
        if not self.mdf_file_path:
            messagebox.showerror("Error", "Please select an MDF file first.")
            return

        output_dir = self.output_dir_entry.get()
        if not output_dir:
            output_dir = Path(self.mdf_file_path).parent
        else:
            output_dir = Path(output_dir)

        output_filename = output_dir / Path(self.mdf_file_path).with_suffix(".xlsx").name

        #try:
        with MDF(self.mdf_file_path) as mdf:
            if not mdf.groups:
                messagebox.showerror("Error", "The selected MDF file has no channel groups.")
                return

            # Convert the entire MDF to dataframe and create a single sheet
            try:
                df = mdf.to_dataframe()
                if not df.empty:
                    with pd.ExcelWriter(output_filename) as writer:
                        df.to_excel(writer, sheet_name="All_Channels", index=False)
                        sheets_created = 1
                        self.progress_bar["value"] = 100
                        self.status_bar.config(text="Converting... 100%")
                        self.master.update_idletasks()
                else:
                    # Create a file with info about empty MDF
                    with pd.ExcelWriter(output_filename) as writer:
                        pd.DataFrame({'Info': ['MDF file contains no data']}).to_excel(writer, sheet_name="Info", index=False)
                        sheets_created = 1
            except Exception as e:
                print(f"Error converting MDF: {e}")
                # Create a file with error info
                with pd.ExcelWriter(output_filename) as writer:
                    pd.DataFrame({'Error': [f'Error converting MDF: {str(e)}']}).to_excel(writer, sheet_name="Error", index=False)

        self.status_label.config(text=f"Successfully converted to {output_filename.name}")
        self.status_bar.config(text="Conversion successful!")
        messagebox.showinfo("Success", f"File converted successfully to:\n{output_filename}")
        
        # Reset progress bar
        self.progress_bar["value"] = 0

        #except Exception as e:
            #self.status_label.config(text="Conversion failed.")
            #self.status_bar.config(text="Conversion failed!")
            #messagebox.showerror("Error", f"An error occurred during conversion:\n{e}")
        #finally:
            #self.progress_bar["value"] = 0

    def create_menu(self):
        menubar = tk.Menu(self.master)
        self.master.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Exit", command=self.master.quit)
        menubar.add_cascade(label="File", menu=file_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)

    def show_about(self):
        messagebox.showinfo("About", "MDF to XLSX Converter\n\nVersion 1.0\n\nCreated by Gemini")

if __name__ == "__main__":
    root = tk.Tk()
    app = MDFConverterApp(root)
    root.geometry("500x400")
    root.mainloop()