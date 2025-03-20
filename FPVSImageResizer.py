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

    # To keep track of skipped files and write errors
    skip_details = []  # list of tuples: (filename, reason)
    write_failures = []

    for file in files:
        if cancel_flag():
            update_callback("Processing cancelled.\n", processed, total_files)
            return skip_details, write_failures, processed

        file_path = os.path.join(input_folder, file)
        _, ext = os.path.splitext(file)
        ext = ext.lower()

        # Skip webp files
        if ext == ".webp":
            skip_details.append((file, ".webp files are not processed (incompatible with PsychoPy)"))
            processed += 1
            update_callback(f"Skipped {file} (unsupported format: .webp).\n", processed, total_files)
            continue

        # Process only valid image extensions
        if ext in valid_exts:
            try:
                img = Image.open(file_path)
            except Exception as e:
                skip_details.append((file, f"Could not read image: {e}"))
                processed += 1
                update_callback(f"Could not read {file}: {e}\n", processed, total_files)
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

            # Confirm before overwriting an existing file.
            if os.path.exists(output_path):
                answer = messagebox.askyesno(
                    "Overwrite Confirmation",
                    f"{new_file_name} already exists in the output folder. Overwrite?"
                )
                if not answer:
                    skip_details.append((file, "User chose not to overwrite existing file."))
                    processed += 1
                    update_callback(f"Skipped {file} (file exists and user declined to overwrite).\n", processed, total_files)
                    continue

            try:
                final_img.save(output_path)
                update_callback(f"Processed {file} -> {new_file_name}\n", processed + 1, total_files)
            except Exception as e:
                write_failures.append((file, str(e)))
                update_callback(f"Error writing {file}: {e}\n", processed + 1, total_files)
        else:
            skip_details.append((file, "Unsupported file format."))
            update_callback(f"Skipped {file} (unsupported format).\n", processed + 1, total_files)

        processed += 1
        update_callback("", processed, total_files)

    return skip_details, write_failures, processed

# --- Main UI Class ---
class FPVSImageResizerApp:
    def __init__(self, master):
        self.master = master
        self.master.title("FPVS Image Resizer")
        self.master.minsize(600, 400)
        self.master.geometry("800x600")
        self.cancel_requested = False

        self.create_menus()
        self.create_widgets()

        sv_ttk.set_theme("light")

    def create_menus(self):
        # Create menubar and add File and Help menus
        menubar = tk.Menu(self.master)

        # File Menu
        file_menu = tk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="Exit", command=self.master.quit)
        menubar.add_cascade(label="File", menu=file_menu)

        # Help Menu with a tabbed help window
        help_menu = tk.Menu(menubar, tearoff=False)
        help_menu.add_command(label="Help Contents", command=self.show_help_window)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.master.config(menu=menubar)

    def show_help_window(self):
        help_win = tk.Toplevel(self.master)
        help_win.title("Help")
        help_win.geometry("600x400")

        notebook = ttk.Notebook(help_win)
        notebook.pack(fill="both", expand=True)

        # User Guide Tab
        user_guide_frame = ttk.Frame(notebook)
        notebook.add(user_guide_frame, text="User Guide")
        self._create_scrollable_text(user_guide_frame,
            "All FPVS images need to be the same size, shape, and filetype before running an experiment in PsychoPy.\n\n"
            "This app helps you quickly standardize all of your images. However, this does NOT mean that all of the images "
            "are suitable for FPVS; Make sure the images aren't cropped too aggressively and that they are still recognizable with a high DPI."
        )

        # FAQ Tab
        faq_frame = ttk.Frame(notebook)
        notebook.add(faq_frame, text="FAQ")
        self._create_scrollable_text(faq_frame,
            "Q: Why are .webp files skipped?\n"
            "A: .webp files are not compatible with PsychoPy. Please replace these images with .jpg, .jpeg, or .png file formats."
        )

        # About Tab
        about_frame = ttk.Frame(notebook)
        notebook.add(about_frame, text="About")
        self._create_scrollable_text(about_frame,
            "FPVS Image Resizer\nVersion 0.5\n\n"
            "Developer: Zack Murphy\nEmail: zmurphy@abe.msstate.edu"
        )

    def _create_scrollable_text(self, parent, content):
        # Create a frame for the Text widget and scrollbar
        text_frame = ttk.Frame(parent)
        text_frame.pack(fill="both", expand=True, padx=10, pady=10)
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")
        text_widget = tk.Text(text_frame, wrap="word", yscrollcommand=scrollbar.set)
        text_widget.insert("1.0", content)
        text_widget.config(state="disabled")
        text_widget.pack(fill="both", expand=True)
        scrollbar.config(command=text_widget.yview)

    def create_widgets(self):
        # ----- Folder Selection Frame -----
        self.folder_frame = ttk.Frame(self.master, padding=10)
        self.folder_frame.pack(fill="x")

        # Input folder frame
        self.input_frame = ttk.Frame(self.folder_frame)
        self.input_frame.pack(fill="x", padx=5, pady=5)
        self.input_button = ttk.Button(self.input_frame, text="Select Input Folder", command=self.select_input_folder, width=18)
        self.input_button.pack(side="left")
        self.input_label = ttk.Label(self.input_frame, text="Input Folder: Not selected", anchor="w")
        self.input_label.pack(side="left", fill="x", expand=True, padx=(10, 0))

        # Output folder frame
        self.output_frame = ttk.Frame(self.folder_frame)
        self.output_frame.pack(fill="x", padx=5, pady=5)
        self.output_button = ttk.Button(self.output_frame, text="Select Output Folder", command=self.select_output_folder, width=18)
        self.output_button.pack(side="left")
        self.output_label = ttk.Label(self.output_frame, text="Output Folder: Not selected", anchor="w")
        self.output_label.pack(side="left", fill="x", expand=True, padx=(10, 0))

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

        # Reset to Default Button
        self.reset_button = ttk.Button(self.settings_frame, text="Reset to Defaults", command=self.reset_defaults)
        self.reset_button.grid(row=0, column=5, padx=5, pady=5)

        # Desired file extension row
        self.extension_label = ttk.Label(self.settings_frame, text="Desired File Extension:")
        self.extension_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.extension_var = tk.StringVar(value=".jpg")
        self.extension_combobox = ttk.Combobox(self.settings_frame, textvariable=self.extension_var, state="readonly", width=8)
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

    def reset_defaults(self):
        self.width_entry.delete(0, tk.END)
        self.width_entry.insert(0, "512")
        self.height_entry.delete(0, tk.END)
        self.height_entry.insert(0, "512")
        self.extension_var.set(".jpg")
        self.status_text.delete("1.0", tk.END)

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
        # Validate numeric inputs
        try:
            target_width = int(self.width_entry.get())
            target_height = int(self.height_entry.get())
            if target_width < 1 or target_height < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Invalid target dimensions. Please enter positive integers.")
            return

        # Check if input folder has valid image files
        valid_files = [f for f in os.listdir(self.input_folder)
                       if os.path.isfile(os.path.join(self.input_folder, f)) and os.path.splitext(f)[1].lower() in
                       [".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"]]
        if not valid_files:
            messagebox.showwarning("No Images Found", "No valid image files were found in the selected input folder.")
            return

        # Check for same input and output folder
        if hasattr(self, 'input_folder') and hasattr(self, 'output_folder'):
            if os.path.abspath(self.input_folder) == os.path.abspath(self.output_folder):
                answer = messagebox.askyesno(
                    "Same Folder Warning",
                    "You selected the same folder for input and output. Existing files may be overwritten.\n\nContinue?"
                )
                if not answer:
                    return

        self.cancel_requested = False
        self.cancel_button.config(state="normal")
        self.status_text.delete("1.0", tk.END)

        try:
            desired_ext = self.extension_var.get().replace(".", "").lower()
        except Exception:
            desired_ext = "jpg"

        # Run processing in a separate thread with exception handling
        thread = threading.Thread(target=self.run_processing, args=(target_width, target_height, desired_ext))
        thread.start()

    def run_processing(self, target_width, target_height, desired_ext):
        try:
            total_files = len([f for f in os.listdir(self.input_folder)
                               if os.path.isfile(os.path.join(self.input_folder, f))])
            skip_details, write_failures, processed = process_images_in_folder(
                self.input_folder,
                self.output_folder,
                target_width,
                target_height,
                desired_ext,
                update_callback=self.update_progress,
                cancel_flag=lambda: self.cancel_requested
            )
            self.cancel_button.config(state="disabled")
            # Prepare summary
            summary = f"Processing complete.\n\nTotal files processed: {processed}\n"
            if skip_details:
                summary += f"\nSkipped {len(skip_details)} files:\n"
                for fname, reason in skip_details:
                    summary += f"  - {fname}: {reason}\n"
            if write_failures:
                summary += f"\nWrite failures for {len(write_failures)} files:\n"
                for fname, error in write_failures:
                    summary += f"  - {fname}: {error}\n"
            messagebox.showinfo("Summary", summary)
            self.open_folder_button.pack(side="left", padx=5, pady=5)
        except Exception as e:
            messagebox.showerror("Processing Error", f"An error occurred during processing:\n{e}")
            self.cancel_button.config(state="disabled")

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
    root.tk.call("tk", "scaling", 1.5)
    app = FPVSImageResizerApp(root)
    root.mainloop()
