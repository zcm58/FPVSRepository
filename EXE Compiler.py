import os
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image
import sv_ttk  # Import the Sun Valley theme library


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

        # --- App Icon Selection ---
        frame_icon = ttk.LabelFrame(master, text="App Icon (png, jpg, jpeg)")
        frame_icon.pack(fill="x", padx=10, pady=5)
        self.icon_file_entry = ttk.Entry(frame_icon, width=50)
        self.icon_file_entry.pack(side="left", padx=5, pady=5)
        self.browse_icon_button = ttk.Button(frame_icon, text="Browse", command=self.browse_icon_file)
        self.browse_icon_button.pack(side="left", padx=5, pady=5)

        # --- Build and Open Directory Buttons ---
        frame_buttons = ttk.Frame(master)
        frame_buttons.pack(pady=10)
        self.build_button = ttk.Button(frame_buttons, text="Build Executable", command=self.start_build)
        self.build_button.pack(side="left", padx=5)
        self.open_dist_button = ttk.Button(frame_buttons, text="Open Compiled Directory",
                                           command=self.open_compiled_directory)
        self.open_dist_button.pack(side="left", padx=5)

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

    def browse_icon_file(self):
        file_path = filedialog.askopenfilename(title="Select App Icon",
                                               filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")])
        if file_path:
            self.icon_file_entry.delete(0, tk.END)
            self.icon_file_entry.insert(0, file_path)

    def start_build(self):
        self.build_button.config(state="disabled")
        threading.Thread(target=self.build_executable, daemon=True).start()

    def convert_icon(self, icon_path):
        # Open the image and ensure it's in RGBA mode
        try:
            img = Image.open(icon_path)
        except Exception as e:
            self.log_output("Error opening icon image: " + str(e))
            return None
        if img.mode != "RGBA":
            img = img.convert("RGBA")
        # Crop to a centered square
        width, height = img.size
        min_dim = min(width, height)
        left = (width - min_dim) // 2
        top = (height - min_dim) // 2
        box = (left, top, left + min_dim, top + min_dim)
        img_cropped = img.crop(box)
        # Resize cropped image to 24x24
        img_small = img_cropped.resize((24, 24), Image.LANCZOS)
        # Create a new 40x40 transparent image
        img_icon = Image.new("RGBA", (40, 40), (0, 0, 0, 0))
        offset = ((40 - 24) // 2, (40 - 24) // 2)
        img_icon.paste(img_small, offset)
        # Save to a temporary .ico file
        temp_icon_path = "temp_icon.ico"
        try:
            img_icon.save(temp_icon_path, format="ICO")
        except Exception as e:
            self.log_output("Error saving temporary icon file: " + str(e))
            return None
        return temp_icon_path

    def build_executable(self):
        main_file = self.main_file_entry.get().strip()
        output_name = self.name_entry.get().strip()
        onefile = self.onefile_var.get()
        windowed = self.windowed_var.get()
        version_file = self.version_file_entry.get().strip()
        icon_file = self.icon_file_entry.get().strip()

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
        if icon_file:
            if os.path.isfile(icon_file):
                temp_icon = self.convert_icon(icon_file)
                if temp_icon:
                    command.extend(["--icon", temp_icon])
                else:
                    self.log_output("Warning: Failed to convert icon. Proceeding without an icon.")
            else:
                self.log_output("Warning: Icon file not found. Ignoring icon option.")
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

    def open_compiled_directory(self):
        dist_path = os.path.join(os.getcwd(), "dist")
        if os.path.exists(dist_path):
            try:
                if os.name == "nt":
                    os.startfile(dist_path)
                elif sys.platform.startswith("darwin"):
                    subprocess.call(["open", dist_path])
                else:
                    subprocess.call(["xdg-open", dist_path])
            except Exception as e:
                messagebox.showerror("Error", "Unable to open the directory: " + str(e))
        else:
            messagebox.showerror("Error", "Compiled directory ('dist') not found.")

    def log_output(self, message):
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)


if __name__ == "__main__":
    root = tk.Tk()
    app = CompilerGUI(root)
    root.mainloop()
