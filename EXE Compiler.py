# easy way to quickly compile new app versions into an .exe file
# allows you to choose an app icon for compilation

import os
import sys
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
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

        # --- App Icon Selection & Interactive Cropping ---
        frame_icon = ttk.LabelFrame(master, text="App Icon (png, jpg, jpeg)")
        frame_icon.pack(fill="x", padx=10, pady=5)
        self.icon_file_entry = ttk.Entry(frame_icon, width=50)
        self.icon_file_entry.pack(side="left", padx=5, pady=5)
        self.browse_icon_button = ttk.Button(frame_icon, text="Browse", command=self.browse_icon_file)
        self.browse_icon_button.pack(side="left", padx=5, pady=5)

        # Replace static preview with an interactive Canvas:
        self.icon_canvas_size = 200  # Canvas size in pixels
        self.icon_canvas = tk.Canvas(frame_icon, width=self.icon_canvas_size, height=self.icon_canvas_size,
                                     bg="#CCCCCC")
        self.icon_canvas.pack(side="left", padx=5, pady=5)
        # We'll display the image on the canvas and overlay a cropping rectangle.
        self.original_icon = None  # PIL.Image; the original icon image
        self.canvas_img = None  # Resized image for display (PhotoImage)
        self.icon_canvas_scale = 1.0  # Scale factor (display size / original size)
        self.image_offset = (0, 0)  # Top-left position where the image is drawn in the canvas

        # Crop rectangle parameters in canvas coordinates (fixed size)
        self.crop_rect_size = 80  # fixed crop rectangle size on canvas
        self.crop_rect_coords = [(self.icon_canvas_size - self.crop_rect_size) // 2,
                                 (self.icon_canvas_size - self.crop_rect_size) // 2,
                                 (self.icon_canvas_size - self.crop_rect_size) // 2 + self.crop_rect_size,
                                 (self.icon_canvas_size - self.crop_rect_size) // 2 + self.crop_rect_size]
        self.crop_rect_id = self.icon_canvas.create_rectangle(*self.crop_rect_coords, outline="red", width=2)
        # For dragging:
        self.drag_data = {"x": 0, "y": 0, "start_coords": None}
        self.icon_canvas.bind("<ButtonPress-1>", self.on_crop_press)
        self.icon_canvas.bind("<B1-Motion>", self.on_crop_drag)

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

    # --- File Browsing Methods ---
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
            self.load_icon_for_crop(file_path)

    # --- Icon Loading and Cropping ---
    def load_icon_for_crop(self, file_path):
        try:
            self.original_icon = Image.open(file_path)
        except Exception as e:
            self.log_output("Error opening icon image: " + str(e))
            return
        # Ensure RGBA
        if self.original_icon.mode != "RGBA":
            self.original_icon = self.original_icon.convert("RGBA")
        orig_w, orig_h = self.original_icon.size
        # Compute scale factor to fit within the canvas while preserving aspect ratio:
        scale = min(self.icon_canvas_size / orig_w, self.icon_canvas_size / orig_h)
        self.icon_canvas_scale = scale
        new_w = int(orig_w * scale)
        new_h = int(orig_h * scale)
        resized_img = self.original_icon.resize((new_w, new_h), Image.LANCZOS)
        self.canvas_img = ImageTk.PhotoImage(resized_img)
        # Center image on canvas:
        x_offset = (self.icon_canvas_size - new_w) // 2
        y_offset = (self.icon_canvas_size - new_h) // 2
        self.image_offset = (x_offset, y_offset)
        self.icon_canvas.delete("IMG")  # Remove any previous image with tag "IMG"
        self.icon_canvas.create_image(x_offset, y_offset, image=self.canvas_img, anchor="nw", tags="IMG")
        # Reset crop rectangle to default (centered)
        self.crop_rect_coords = [(self.icon_canvas_size - self.crop_rect_size) // 2,
                                 (self.icon_canvas_size - self.crop_rect_size) // 2,
                                 (self.icon_canvas_size - self.crop_rect_size) // 2 + self.crop_rect_size,
                                 (self.icon_canvas_size - self.crop_rect_size) // 2 + self.crop_rect_size]
        self.icon_canvas.coords(self.crop_rect_id, *self.crop_rect_coords)

    def on_crop_press(self, event):
        # Check if click is inside crop rectangle; if so, record the offset.
        x1, y1, x2, y2 = self.icon_canvas.coords(self.crop_rect_id)
        if x1 <= event.x <= x2 and y1 <= event.y <= y2:
            self.drag_data["x"] = event.x
            self.drag_data["y"] = event.y
            self.drag_data["start_coords"] = (x1, y1, x2, y2)
        else:
            self.drag_data["start_coords"] = None

    def on_crop_drag(self, event):
        if self.drag_data["start_coords"] is None:
            return
        # Calculate how far the mouse has moved:
        dx = event.x - self.drag_data["x"]
        dy = event.y - self.drag_data["y"]
        # Get the new coordinates:
        x1, y1, x2, y2 = self.drag_data["start_coords"]
        new_x1 = x1 + dx
        new_y1 = y1 + dy
        new_x2 = x2 + dx
        new_y2 = y2 + dy
        # Constrain rectangle to stay within the canvas boundaries:
        if new_x1 < 0:
            new_x1 = 0
            new_x2 = self.crop_rect_size
        if new_y1 < 0:
            new_y1 = 0
            new_y2 = self.crop_rect_size
        if new_x2 > self.icon_canvas_size:
            new_x2 = self.icon_canvas_size
            new_x1 = self.icon_canvas_size - self.crop_rect_size
        if new_y2 > self.icon_canvas_size:
            new_y2 = self.icon_canvas_size
            new_y1 = self.icon_canvas_size - self.crop_rect_size
        self.crop_rect_coords = [new_x1, new_y1, new_x2, new_y2]
        self.icon_canvas.coords(self.crop_rect_id, *self.crop_rect_coords)

    def get_cropped_icon(self):
        """
        Maps the crop rectangle from canvas coordinates to the original image coordinates.
        Then crops the original image, resizes it to 24x24, centers it in a 40x40 canvas,
        and saves it as a temporary .ico file.
        """
        if self.original_icon is None:
            return None
        # Get crop coordinates in canvas:
        crop_x1, crop_y1, crop_x2, crop_y2 = self.crop_rect_coords
        # The displayed image is drawn at self.image_offset on the canvas.
        offset_x, offset_y = self.image_offset
        # The crop rectangle relative to the displayed image:
        img_crop_x1 = crop_x1 - offset_x
        img_crop_y1 = crop_y1 - offset_y
        img_crop_x2 = crop_x2 - offset_x
        img_crop_y2 = crop_y2 - offset_y
        # Ensure the crop lies within the displayed image:
        # (Clamp coordinates if necessary)
        display_w = self.canvas_img.width()
        display_h = self.canvas_img.height()
        img_crop_x1 = max(0, img_crop_x1)
        img_crop_y1 = max(0, img_crop_y1)
        img_crop_x2 = min(display_w, img_crop_x2)
        img_crop_y2 = min(display_h, img_crop_y2)
        # Map to original image coordinates by dividing by the scale factor:
        scale = self.icon_canvas_scale
        orig_crop_box = (int(img_crop_x1 / scale),
                         int(img_crop_y1 / scale),
                         int(img_crop_x2 / scale),
                         int(img_crop_y2 / scale))
        try:
            cropped = self.original_icon.crop(orig_crop_box)
        except Exception as e:
            self.log_output("Error cropping icon: " + str(e))
            return None
        # Resize cropped image to 24x24:
        try:
            resized_cropped = cropped.resize((24, 24), Image.LANCZOS)
        except Exception as e:
            self.log_output("Error resizing cropped icon: " + str(e))
            return None
        # Create a new 40x40 transparent image and paste the resized crop centered:
        final_icon = Image.new("RGBA", (40, 40), (0, 0, 0, 0))
        offset = ((40 - 24) // 2, (40 - 24) // 2)
        final_icon.paste(resized_cropped, offset)
        temp_icon_path = "temp_icon.ico"
        try:
            final_icon.save(temp_icon_path, format="ICO")
        except Exception as e:
            self.log_output("Error saving temporary icon file: " + str(e))
            return None
        return temp_icon_path

    def start_build(self):
        self.build_button.config(state="disabled")
        threading.Thread(target=self.build_executable, daemon=True).start()

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
        if icon_file and os.path.isfile(icon_file):
            # Use the interactive crop data to generate the .ico file.
            temp_icon = self.get_cropped_icon()
            if temp_icon:
                command.extend(["--icon", temp_icon])
            else:
                self.log_output("Warning: Failed to process icon. Proceeding without an icon.")
        else:
            if icon_file:
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
