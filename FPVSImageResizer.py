import os
import sys
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from PIL import Image
import sv_ttk  # Import the Sun Valley theme library

# Set DPI awareness on Windows
if sys.platform.startswith('win'):
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

# --- Modular image processing function ---
def process_images_in_folder(input_folder, output_folder, target_width, target_height, desired_ext, update_callback, cancel_flag):
    valid_exts = [".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"]
    files = [f for f in os.listdir(input_folder) if os.path.isfile(os.path.join(input_folder, f))]
    total_files = len(files)
    processed = 0
    webp_list = []

    for file in files:
        if cancel_flag():
            update_callback("Processing cancelled.\n", processed, total_files)
            return webp_list, processed

        file_path = os.path.join(input_folder, file)
        _, ext = os.path.splitext(file)
        ext = ext.lower()

        # Skip webp files
        if ext == ".webp":
            webp_list.append(file)
            processed += 1
            update_callback("", processed, total_files)
            continue

        # Process only valid image extensions
        if ext in valid_exts:
            try:
                img = Image.open(file_path)
            except Exception as e:
                update_callback(f"Could not read {file}: {e}\n", processed + 1, total_files)
                processed += 1
                continue

            # Resize logic
            orig_width, orig_height = img.size
            scale = max(target_width / orig_width, target_height / orig_height)
            new_size = (round(orig_width * scale), round(orig_height * scale))
            resized_img = img.resize(new_size, resample=Image.Resampling.LANCZOS)
            left = (new_size[0] - target_width) // 2
            top = (new_size[1] - target_height) // 2
            right = left + target_width
            bottom = top + target_height
            final_img = resized_img.crop((left, top, right, bottom))

            base_name, _ = os.path.splitext(file)
            new_file_name = f"{base_name} Resized.{desired_ext}"
            output_path = os.path.join(output_folder, new_file_name)

            try:
                final_img.save(output_path)
                update_callback(f"Processed {file} -> {new_file_name}\n", processed + 1, total_files)
            except Exception as e:
                update_callback(f"Error writing {file}: {e}\n", processed + 1, total_files)
        else:
            update_callback(f"Skipped {file} (unsupported format).\n", processed + 1, total_files)

        processed += 1
        update_callback("", processed, total_files)

    return webp_list, processed


# --- Main UI Class ---
class FPVSImageResizerApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Image Resizer App")
        self.master.geometry("800x600")
        self.cancel_requested = False

        # Create a simple menubar to switch themes dynamically.
        self.create_theme_menu()

        # Create all other widgets.
        self.create_widgets()

        # Set the default Sun Valley theme to "dark"
        sv_ttk.set_theme("light")

    def create_theme_menu(self):
        """Creates a menubar with a 'Theme' menu to switch between light and dark Sun Valley themes."""
        menubar = tk.Menu(self.master)
        theme_menu = tk.Menu(menubar, tearoff=False)
        theme_menu.add_command(label="Light Theme", command=lambda: sv_ttk.set_theme("light"))
        theme_menu.add_command(label="Dark Theme", command=lambda: sv_ttk.set_theme("dark"))
        menubar.add_cascade(label="Choose your Theme", menu=theme_menu)
        self.master.config(menu=menubar)

    def create_widgets(self):


        # ----- Folder Selection Frame -----
        self.folder_frame = ttk.Frame(self.master, padding=10)
        self.folder_frame.pack(fill="x")

        # Input folder frame
        self.input_frame = ttk.Frame(self.folder_frame)
        self.input_frame.pack(fill="x", padx=5, pady=5)
        self.input_button = ttk.Button(self.input_frame, text="Select Input Folder", command=self.select_input_folder,width=18)
        self.input_button.pack(side="left")
        self.input_label = ttk.Label(self.input_frame, text="Input Folder: Not selected", anchor="w")
        self.input_label.pack(side="left", fill="x", expand=True)

        # Output folder frame
        self.output_frame = ttk.Frame(self.folder_frame)
        self.output_frame.pack(fill="x", padx=5, pady=5)
        self.output_button = ttk.Button(self.output_frame, text="Select Output Folder", command=self.select_output_folder,width=18)
        self.output_button.pack(side="left")
        self.output_label = ttk.Label(self.output_frame, text="Output Folder: Not selected", anchor="w")
        self.output_label.pack(side="left", fill="x", expand=True)

        # ----- Settings Frame -----
        self.settings_frame = ttk.Frame(self.master, padding=10)
        self.settings_frame.pack(fill="x")

        self.width_label = ttk.Label(self.settings_frame, text="Image Width:")
        self.width_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.width_entry = ttk.Entry(self.settings_frame, width=10)
        self.width_entry.insert(0, "512")
        self.width_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        self.height_label = ttk.Label(self.settings_frame, text="Image Height:")
        self.height_label.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.height_entry = ttk.Entry(self.settings_frame, width=10)
        self.height_entry.insert(0, "512")
        self.height_entry.grid(row=0, column=3, padx=5, pady=5, sticky="w")

        self.default_size_label = ttk.Label(
            self.settings_frame,
            text="512x512 is the default option for FPVS image sizes."
        )
        self.default_size_label.grid(row=0, column=4, padx=5, pady=5, sticky="w")

        self.extension_label = ttk.Label(self.settings_frame, text="Desired File Extension:")
        self.extension_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.extension_var = tk.StringVar(value=".jpg")
        self.extension_combobox = ttk.Combobox(
            self.settings_frame, textvariable=self.extension_var, state="readonly", width=8
        )
        self.extension_combobox['values'] = (".jpg", ".jpeg", ".png")
        self.extension_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # ----- Processing Frame -----
        self.process_frame = ttk.Frame(self.master, padding=10)
        self.process_frame.pack(fill="x")

        self.process_button = ttk.Button(self.process_frame, text="Process Images", command=self.start_processing)
        self.process_button.pack(side="left", padx=5, pady=5)

        self.cancel_button = ttk.Button(self.process_frame, text="Cancel", command=self.cancel_processing, state="disabled")
        self.cancel_button.pack(side="left", padx=5, pady=5)

        self.open_folder_button = ttk.Button(self.process_frame, text="Open Resized Images Folder", command=self.open_resized_folder)
        self.open_folder_button.pack(side="left", padx=5, pady=5)
        self.open_folder_button.pack_forget()  # Hide until processing completes

        self.progress = ttk.Progressbar(self.process_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=5, pady=5, expand=True)

        self.progress_label = ttk.Label(self.process_frame, text="Processed 0 of 0")
        self.progress_label.pack(side="left", padx=5)

        # ----- Status Frame -----
        self.status_frame = ttk.Frame(self.master, padding=10)
        self.status_frame.pack(fill="both", expand=True)

        self.status_text = tk.Text(self.status_frame, height=10)
        self.status_text.pack(fill="both", expand=True, padx=5, pady=5)

    def select_input_folder(self):
        folder = filedialog.askdirectory(title="Select input folder containing images")
        if folder:
            self.input_folder = folder
            self.input_label.config(text=f"Input Folder: {folder}")

    def select_output_folder(self):
        if not hasattr(self, 'input_folder'):
            messagebox.showerror("Error", "Select an Input Folder before selecting an Output Folder.")
            return

        answer = messagebox.askquestion(
            "Output Folder Suggestion",
            "We suggest creating a new folder inside the input folder to store your newly resized images. Would you like to do that?"
        )
        if answer == "yes":
            new_folder = os.path.join(self.input_folder, "Resized Images")
            if not os.path.exists(new_folder):
                os.makedirs(new_folder)
            self.output_folder = new_folder
            self.output_label.config(text=f"Output Folder: {new_folder}")
        else:
            messagebox.showinfo("Notice", "Ah. You don't follow the rules. I guess you can select your own folder.")
            folder = filedialog.askdirectory(initialdir=self.input_folder, title="Select output folder for saving images")
            if folder:
                self.output_folder = folder
                self.output_label.config(text=f"Output Folder: {folder}")

    def update_progress(self, message, processed, total):
        if message:
            self.status_text.insert(tk.END, message)
            self.status_text.see(tk.END)
        self.progress['value'] = processed
        self.progress['maximum'] = total
        self.progress_label.config(text=f"Processed {processed} of {total}")

    def cancel_processing(self):
        self.cancel_requested = True
        self.cancel_button.config(state="disabled")

    def start_processing(self):
        if not hasattr(self, 'input_folder'):
            messagebox.showerror("Error", "Please select an input folder.")
            return
        if not hasattr(self, 'output_folder'):
            messagebox.showerror("Error", "Please select an output folder.")
            return

        # Parse width/height
        try:
            target_width = int(self.width_entry.get())
            target_height = int(self.height_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid target dimensions. Please enter valid numbers.")
            return

        self.cancel_requested = False
        self.cancel_button.config(state="normal")
        self.status_text.delete("1.0", tk.END)

        try:
            desired_ext = self.extension_var.get().replace(".", "").lower()
        except Exception:
            desired_ext = "jpg"

        # Run the processing in a separate thread
        thread = threading.Thread(
            target=self.run_processing,
            args=(target_width, target_height, desired_ext)
        )
        thread.start()

    def run_processing(self, target_width, target_height, desired_ext):
        total_files = len([
            f for f in os.listdir(self.input_folder)
            if os.path.isfile(os.path.join(self.input_folder, f))
        ])

        webp_list, processed = process_images_in_folder(
            self.input_folder,
            self.output_folder,
            target_width,
            target_height,
            desired_ext,
            update_callback=self.update_progress,
            cancel_flag=lambda: self.cancel_requested
        )

        self.cancel_button.config(state="disabled")

        if self.cancel_requested:
            messagebox.showinfo("Cancelled", "Processing was cancelled.")
        else:
            success_msg = (
                f"Congrats! Your images have been successfully resized to "
                f"{target_width}x{target_height} and converted to .{desired_ext} format."
            )
            messagebox.showinfo("Success", success_msg)
            self.open_folder_button.pack(side="left", padx=5, pady=5)

            if webp_list:
                alert_msg = (
                    "Note: The following .webp files were skipped:\n"
                    + "\n".join(webp_list)
                    + "\n\n.webp filetypes will not work with PsychoPy. "
                    "Please replace these images with .jpg, .jpeg, or .png images only."
                )
                messagebox.showwarning("Warning", alert_msg)

    def open_resized_folder(self):
        if not hasattr(self, 'output_folder'):
            messagebox.showerror("Error", "Output folder not selected.")
            return

        messagebox.showinfo(
            "Attention",
            "IMPORTANT: Please manually check your images. Ensure that images are still "
            "high quality and that cropping wasn't too aggressive. Images used for FPVS "
            "should still be easily recognizable after resizing."
        )
        try:
            if sys.platform.startswith('win'):
                os.startfile(self.output_folder)
            elif sys.platform.startswith('darwin'):
                subprocess.call(["open", self.output_folder])
            else:
                subprocess.call(["xdg-open", self.output_folder])
        except Exception as e:
            messagebox.showerror("Error", f"Unable to open the folder: {e}")


if __name__ == "__main__":
    root = tk.Tk()  # Create the main window
    # Adjust the scaling factor as needed for High DPI screens
    root.tk.call("tk", "scaling", 1.5)

    app = FPVSImageResizerApp(root)
    root.mainloop()
