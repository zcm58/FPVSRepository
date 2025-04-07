import os
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


class CompilerGUI:
    def __init__(self, master):
        self.master = master
        master.title("PyInstaller Build GUI")

        # --- Main Python File Selection ---
        frame_main = ttk.LabelFrame(master, text="Main Python File")
        frame_main.pack(fill="x", padx=10, pady=5)
        self.main_file_entry = ttk.Entry(frame_main, width=50)
        self.main_file_entry.insert(0, "fpvs_image_resizer.py")
        self.main_file_entry.pack(side="left", padx=5, pady=5)
        self.browse_main_button = ttk.Button(frame_main, text="Browse", command=self.browse_main_file)
        self.browse_main_button.pack(side="left", padx=5, pady=5)

        # --- Output Name ---
        frame_name = ttk.LabelFrame(master, text="Output Name")
        frame_name.pack(fill="x", padx=10, pady=5)
        self.name_entry = ttk.Entry(frame_name, width=50)
        self.name_entry.insert(0, "FPVSImageResizer")
        self.name_entry.pack(side="left", padx=5, pady=5)

        # --- Options: Onefile and Windowed ---
        frame_options = ttk.LabelFrame(master, text="Options")
        frame_options.pack(fill="x", padx=10, pady=5)
        self.onefile_var = tk.BooleanVar(value=True)
        self.windowed_var = tk.BooleanVar(value=True)
        self.onefile_check = ttk.Checkbutton(frame_options, text="One File", variable=self.onefile_var)
        self.onefile_check.pack(side="left", padx=5, pady=5)
        self.windowed_check = ttk.Checkbutton(frame_options, text="Windowed", variable=self.windowed_var)
        self.windowed_check.pack(side="left", padx=5, pady=5)

        # --- Optional Version File ---
        frame_version = ttk.LabelFrame(master, text="Version File (Optional)")
        frame_version.pack(fill="x", padx=10, pady=5)
        self.version_file_entry = ttk.Entry(frame_version, width=50)
        self.version_file_entry.pack(side="left", padx=5, pady=5)
        self.browse_version_button = ttk.Button(frame_version, text="Browse", command=self.browse_version_file)
        self.browse_version_button.pack(side="left", padx=5, pady=5)

        # --- Build Button ---
        self.build_button = ttk.Button(master, text="Build Executable", command=self.start_build)
        self.build_button.pack(pady=10)

        # --- Output Log ---
        self.output_text = tk.Text(master, height=10)
        self.output_text.pack(fill="both", padx=10, pady=5, expand=True)

    def browse_main_file(self):
        file_path = filedialog.askopenfilename(title="Select Main Python File", filetypes=[("Python Files", "*.py")])
        if file_path:
            self.main_file_entry.delete(0, tk.END)
            self.main_file_entry.insert(0, file_path)

    def browse_version_file(self):
        file_path = filedialog.askopenfilename(title="Select Version File",
                                               filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")])
        if file_path:
            self.version_file_entry.delete(0, tk.END)
            self.version_file_entry.insert(0, file_path)

    def start_build(self):
        self.build_button.config(state="disabled")
        threading.Thread(target=self.build_executable, daemon=True).start()

    def build_executable(self):
        main_file = self.main_file_entry.get().strip()
        output_name = self.name_entry.get().strip()
        onefile = self.onefile_var.get()
        windowed = self.windowed_var.get()
        version_file = self.version_file_entry.get().strip()

        if not os.path.isfile(main_file):
            self.log_output("Error: Main file not found.")
            self.build_button.config(state="normal")
            return

        command = ["pyinstaller"]
        if onefile:
            command.append("--onefile")
        if windowed:
            command.append("--windowed")
        command.extend(["--name", output_name])
        if version_file:
            if os.path.isfile(version_file):
                command.extend(["--version-file", version_file])
            else:
                self.log_output("Warning: Version file not found. Ignoring version file option.")
        command.append(main_file)

        self.log_output("Running command: " + " ".join(command))

        try:
            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                       universal_newlines=True)
            for line in process.stdout:
                self.log_output(line.strip())
            process.wait()
            if process.returncode == 0:
                self.log_output("Build completed successfully. Check the 'dist' folder for the executable.")
            else:
                self.log_output("Build failed with return code " + str(process.returncode))
        except Exception as e:
            self.log_output("Build failed: " + str(e))
        self.build_button.config(state="normal")

    def log_output(self, message):
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)


if __name__ == "__main__":
    root = tk.Tk()
    app = CompilerGUI(root)
    root.mainloop()
