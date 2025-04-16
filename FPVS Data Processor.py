#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EEG FPVS Analysis GUI using MNE-Python and CustomTkinter.

Version: 1.10 (April 2025) - Revised to use annotations for event extraction and
string-based reference channels with dynamic GUI resizing.

Key functionalities:
- Modern GUI using CustomTkinter.
- Load BioSemi EEG data (.BDF, .set).
- Process single files or multiple files at once.
- Preprocessing Steps (adapted from EEGlab workflow)
    - Import the file using mastoid references (default: EXG1 and EXG2).
    - Downsample if necessary (default 256 Hz).
    - Apply standard 10-20 electrode layout
    - After importing and re-referencing, remove channels EXG1 through EXG8
    - Bandpass filter (default 0.1–50 Hz).
    - Kurtosis-based channel rejection & interpolation.
    - Re-reference to the average reference
- Extract epochs based on textual annotations (e.g., "condition 1", "condition 2", etc.).
- Post-processing using FFT, SNR, Z-score, BCA computation.
- Saves Excel files with separate sheets per metric.
- Background processing with progress updates.
"""

# === Dependencies ===
# Standard Libraries:
import os
import glob
import tkinter as tk
from tkinter import filedialog, messagebox
import sys
import traceback
import threading
import queue
import gc  # Garbage Collector

# Third-Party Libraries:
import numpy as np
import pandas as pd
from scipy.stats import kurtosis

try:
    import customtkinter as ctk
except ImportError:
    messagebox.showerror("Dependency Error", "CustomTkinter required. pip install customtkinter")
    sys.exit(1)

try:
    import mne
except ImportError:
    messagebox.showerror("Dependency Error", "MNE-Python required. pip install mne")
    sys.exit(1)

try:
    import xlsxwriter
except ImportError:
    messagebox.showerror("Dependency Error", "XlsxWriter required. pip install xlsxwriter")
    sys.exit(1)
# === End Dependencies ===

# =====================================================
# Fixed electrode names for 64 electrode biosemi layout
# =====================================================
TARGET_FREQUENCIES = np.arange(1.2, 16.8 + 1.2, 1.2)
DEFAULT_ELECTRODE_NAMES_64 = [
    'Fp1', 'AF7', 'AF3', 'F1', 'F3', 'F5', 'F7', 'FT7', 'FC5', 'FC3',
    'FC1', 'C1', 'C3', 'C5', 'T7', 'TP7', 'CP5', 'CP3', 'CP1', 'P1',
    'P3', 'P5', 'P7', 'P9', 'PO7', 'PO3', 'O1', 'Iz', 'Oz', 'POz', 'Pz',
    'CPz', 'Fpz', 'Fp2', 'AF8', 'AF4', 'AFz', 'Fz', 'F2', 'F4', 'F6',
    'F8', 'FT8', 'FC6', 'FC4', 'FC2', 'FCz', 'Cz', 'C2', 'C4', 'C6',
    'T8', 'TP8', 'CP6', 'CP4', 'CP2', 'P2', 'P4', 'P6', 'P8', 'P10',
    'PO8', 'PO4', 'O2'
]

# =====================================================
# GUI Configuration
# =====================================================
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")
CORNER_RADIUS = 8
PAD_X = 5
PAD_Y = 5
ENTRY_WIDTH = 100


class FPVSApp(ctk.CTk):
    """ Main application class replicating MATLAB FPVS analysis workflow. """

    def __init__(self):
        super().__init__()
        today_str = pd.Timestamp.now().strftime('%Y-%m-%d')
        self.title(f"EEG FPVS Analysis Tool (v1.10 - {today_str})")
        self.geometry("1000x950")  # initial geometry; will adjust dynamically later

        # Data structures
        self.preprocessed_data = {}
        self.condition_entries = []  # List of strings for condition (e.g., "condition 1")
        self.current_event_ids_process = []  # Will hold condition strings extracted from GUI
        self.data_paths = []
        self.processing_thread = None
        self.detection_thread = None
        self.gui_queue = queue.Queue()
        self._max_progress = 1
        self.validated_params = {}  # To store validated parameters

        self.create_menu()
        self.create_widgets()
        self.log("Welcome to the EEG FPVS Analysis Tool (Python Version)!")
        self.log(f"Appearance Mode: {ctk.get_appearance_mode()}")

        # Dynamically adjust the window size to ensure all widgets (including process button) are visible.
        self.update_idletasks()
        req_width = self.winfo_reqwidth()
        req_height = self.winfo_reqheight()
        self.geometry(f"{req_width}x{req_height}")

        if self.condition_entries:
            self.condition_entries[0].focus_set()

    # --- Menu Methods ---
    def create_menu(self):
        self.menubar = tk.Menu(self)
        self.config(menu=self.menubar)
        file_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="File", menu=file_menu)
        appearance_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="Appearance", menu=appearance_menu)
        appearance_menu.add_command(label="Dark Mode", command=lambda: self.set_appearance_mode("Dark"))
        appearance_menu.add_command(label="Light Mode", command=lambda: self.set_appearance_mode("Light"))
        appearance_menu.add_command(label="System Default", command=lambda: self.set_appearance_mode("System"))
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)
        tools_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Placeholder Tool...", command=lambda: messagebox.showinfo("Placeholder", "Placeholder tool."))
        help_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About...", command=self.show_about_dialog)

    def set_appearance_mode(self, mode):
        self.log(f"Setting appearance mode to: {mode}")
        ctk.set_appearance_mode(mode)

    def show_about_dialog(self):
        messagebox.showinfo("About EEG FPVS Analysis Tool",
                            f"Version: 1.10 ({pd.Timestamp.now().strftime('%B %Y')})")

    def quit(self):
        if (self.processing_thread and self.processing_thread.is_alive()) or \
           (self.detection_thread and self.detection_thread.is_alive()):
            if messagebox.askyesno("Exit Confirmation", "Processing or detection ongoing. Stop and exit?"):
                self.destroy()
            else:
                return
        else:
            self.destroy()

    # --- Validation Methods ---
    # For numeric entries we still retain these, but for condition and ref channel entries we allow strings.
    def _validate_numeric_input(self, P):
        if P == "" or P == "-":
            return True
        try:
            float(P)
            return True
        except ValueError:
            self.bell()
            return False

    # --- GUI Creation ---
    def create_widgets(self):
        """ Builds and arranges GUI components. """
        # For condition text fields, no validation command is required
        validate_num_cmd = (self.register(self._validate_numeric_input), '%P')

        main_frame = ctk.CTkFrame(self, corner_radius=0)
        main_frame.pack(fill="both", expand=True, padx=PAD_X*2, pady=PAD_Y*2)
        main_frame.grid_rowconfigure(3, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        # Options Frame inside the GUI
        self.options_frame = ctk.CTkFrame(main_frame)
        self.options_frame.pack(fill="x", padx=PAD_X, pady=PAD_Y)
        ctk.CTkLabel(
            self.options_frame, text="Processing Options",
            font=ctk.CTkFont(weight="bold")
        ).grid(row=0, column=0, columnspan=4, sticky="w", padx=PAD_X, pady=(PAD_Y, PAD_Y*2))
        ctk.CTkLabel(self.options_frame, text="Mode:").grid(row=1, column=0, sticky="w", padx=PAD_X, pady=PAD_Y)

        self.file_mode = tk.StringVar(value="Single")
        self.radio_single = ctk.CTkRadioButton(
            self.options_frame, text="Single File", variable=self.file_mode,
            value="Single", command=self.update_select_button_text, corner_radius=CORNER_RADIUS
        )
        self.radio_single.grid(row=1, column=1, padx=PAD_X, pady=PAD_Y, sticky="w")
        self.radio_batch = ctk.CTkRadioButton(
            self.options_frame, text="Batch Folder", variable=self.file_mode,
            value="Batch", command=self.update_select_button_text, corner_radius=CORNER_RADIUS
        )
        self.radio_batch.grid(row=1, column=2, padx=PAD_X, pady=PAD_Y, sticky="w")
        ctk.CTkLabel(self.options_frame, text="File Type:").grid(row=2, column=0, sticky="w", padx=PAD_X, pady=PAD_Y)

        self.file_type = tk.StringVar(value=".BDF")
        rb_bdf = ctk.CTkRadioButton(
            self.options_frame, text=".BDF", variable=self.file_type,
            value=".BDF", corner_radius=CORNER_RADIUS
        )
        rb_bdf.grid(row=2, column=1, padx=PAD_X, pady=PAD_Y, sticky="w")
        rb_set = ctk.CTkRadioButton(
            self.options_frame, text=".set", variable=self.file_type,
            value=".set", corner_radius=CORNER_RADIUS
        )
        rb_set.grid(row=2, column=2, padx=PAD_X, pady=PAD_Y, sticky="w")

        self.select_button_text = tk.StringVar()
        self.select_button = ctk.CTkButton(
            self.options_frame, textvariable=self.select_button_text,
            command=self.select_data_source, corner_radius=CORNER_RADIUS
        )
        self.select_button.grid(row=1, column=3, rowspan=2, padx=PAD_X*2, pady=PAD_Y, sticky="ew")
        self.options_frame.grid_columnconfigure(3, weight=1)

        # Params Frame
        self.params_frame = ctk.CTkFrame(main_frame)
        self.params_frame.pack(fill="x", padx=PAD_X, pady=PAD_Y)
        ctk.CTkLabel(
            self.params_frame, text="Preprocessing Parameters",
            font=ctk.CTkFont(weight="bold")
        ).grid(row=0, column=0, columnspan=6, sticky="w", padx=PAD_X, pady=(PAD_Y, PAD_Y*2))

        ctk.CTkLabel(self.params_frame, text="Low Pass (Hz):").grid(row=1, column=0, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.low_pass_entry = ctk.CTkEntry(
            self.params_frame, width=ENTRY_WIDTH, validate='key',
            validatecommand=validate_num_cmd, corner_radius=CORNER_RADIUS
        )
        self.low_pass_entry.insert(0, "0.1")
        self.low_pass_entry.grid(row=1, column=1, padx=PAD_X, pady=PAD_Y)

        ctk.CTkLabel(self.params_frame, text="High Pass (Hz):").grid(row=1, column=2, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.high_pass_entry = ctk.CTkEntry(
            self.params_frame, width=ENTRY_WIDTH, validate='key',
            validatecommand=validate_num_cmd, corner_radius=CORNER_RADIUS
        )
        self.high_pass_entry.insert(0, "50")
        self.high_pass_entry.grid(row=1, column=3, padx=PAD_X, pady=PAD_Y)

        ctk.CTkLabel(self.params_frame, text="Downsample (Hz):").grid(row=2, column=0, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.downsample_entry = ctk.CTkEntry(
            self.params_frame, width=ENTRY_WIDTH, validate='key',
            validatecommand=validate_num_cmd, corner_radius=CORNER_RADIUS
        )
        self.downsample_entry.insert(0, "256")
        self.downsample_entry.grid(row=2, column=1, padx=PAD_X, pady=PAD_Y)

        ctk.CTkLabel(self.params_frame, text="Epoch Start (s):").grid(row=2, column=2, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.epoch_start_entry = ctk.CTkEntry(
            self.params_frame, width=ENTRY_WIDTH, validate='key',
            validatecommand=validate_num_cmd, corner_radius=CORNER_RADIUS
        )
        self.epoch_start_entry.insert(0, "-1")
        self.epoch_start_entry.grid(row=2, column=3, padx=PAD_X, pady=PAD_Y)

        ctk.CTkLabel(self.params_frame, text="Epoch End (s):").grid(row=2, column=4, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.epoch_end_entry = ctk.CTkEntry(
            self.params_frame, width=ENTRY_WIDTH+20, validate='key',
            validatecommand=validate_num_cmd, corner_radius=CORNER_RADIUS
        )
        self.epoch_end_entry.insert(0, "125")
        self.epoch_end_entry.grid(row=2, column=5, padx=PAD_X, pady=PAD_Y)  # Default 125s

        ctk.CTkLabel(self.params_frame, text="Rejection Z-Threshold:").grid(row=3, column=0, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.reject_thresh_entry = ctk.CTkEntry(
            self.params_frame, width=ENTRY_WIDTH, validate='key',
            validatecommand=validate_num_cmd, corner_radius=CORNER_RADIUS
        )
        self.reject_thresh_entry.insert(0, "5")
        self.reject_thresh_entry.grid(row=3, column=1, padx=PAD_X, pady=PAD_Y)

        # Reference Channels now as strings; defaults are "EXG1" and "EXG2"
        ctk.CTkLabel(self.params_frame, text="Reference Channel 1:").grid(row=4, column=0, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.ref_channel1_entry = ctk.CTkEntry(
            self.params_frame, width=ENTRY_WIDTH, corner_radius=CORNER_RADIUS
        )
        self.ref_channel1_entry.insert(0, "EXG1")
        self.ref_channel1_entry.grid(row=4, column=1, padx=PAD_X, pady=PAD_Y)

        ctk.CTkLabel(self.params_frame, text="Reference Channel 2:").grid(row=4, column=2, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.ref_channel2_entry = ctk.CTkEntry(
            self.params_frame, width=ENTRY_WIDTH, corner_radius=CORNER_RADIUS
        )
        self.ref_channel2_entry.insert(0, "EXG2")
        self.ref_channel2_entry.grid(row=4, column=3, padx=PAD_X, pady=PAD_Y)

        ctk.CTkLabel(self.params_frame, text="Max EEG Channels Keep:").grid(row=4, column=4, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.max_idx_keep_entry = ctk.CTkEntry(
            self.params_frame, width=ENTRY_WIDTH, validate='key',
            validatecommand=validate_num_cmd, corner_radius=CORNER_RADIUS
        )
        self.max_idx_keep_entry.insert(0, "64")
        self.max_idx_keep_entry.grid(row=4, column=5, padx=PAD_X, pady=PAD_Y)

        # Event IDs Frame – now these are free-text fields (no numeric validation)
        conditions_frame_outer = ctk.CTkFrame(main_frame)
        conditions_frame_outer.pack(fill="both", expand=True, padx=PAD_X, pady=PAD_Y)
        ctk.CTkLabel(
            conditions_frame_outer, text="Event IDs (Conditions) to Extract",
            font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w", padx=PAD_X, pady=(PAD_Y, 0))
        self.conditions_scroll_frame = ctk.CTkScrollableFrame(conditions_frame_outer, label_text="")
        self.conditions_scroll_frame.pack(fill="both", expand=True, padx=PAD_X, pady=(0, PAD_Y))
        self.conditions_scroll_frame.grid_columnconfigure(0, weight=1)
        self.condition_entries = []
        self.add_event_id_entry()  # Pre-populate a single field

        condition_button_frame = ctk.CTkFrame(conditions_frame_outer, fg_color="transparent")
        condition_button_frame.pack(fill="x", pady=(0, PAD_Y), padx=PAD_X)
        self.detect_button = ctk.CTkButton(
            condition_button_frame, text="Detect Event IDs",
            command=self.detect_and_populate_event_ids, corner_radius=CORNER_RADIUS
        )
        self.detect_button.pack(side="left", padx=(0, PAD_X))
        self.add_cond_button = ctk.CTkButton(
            condition_button_frame, text="Add Event ID Field",
            command=self.add_event_id_entry, corner_radius=CORNER_RADIUS
        )
        self.add_cond_button.pack(side="left")

        # Save Location Frame
        self.save_frame = ctk.CTkFrame(main_frame)
        self.save_frame.pack(fill="x", padx=PAD_X, pady=PAD_Y)
        ctk.CTkLabel(
            self.save_frame, text="Excel Output Save Location",
            font=ctk.CTkFont(weight="bold")
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=PAD_X, pady=(PAD_Y, PAD_Y*2))
        self.save_folder_path = tk.StringVar()
        btn_select_save = ctk.CTkButton(
            self.save_frame, text="Select Output Folder",
            command=self.select_save_folder, corner_radius=CORNER_RADIUS
        )
        btn_select_save.grid(row=1, column=0, padx=PAD_X, pady=PAD_Y)
        self.save_folder_display = ctk.CTkEntry(
            self.save_frame, textvariable=self.save_folder_path,
            state="readonly", corner_radius=CORNER_RADIUS
        )
        self.save_folder_display.grid(row=1, column=1, sticky="ew", padx=PAD_X, pady=PAD_Y)
        self.save_frame.grid_columnconfigure(1, weight=1)

        # Bottom Controls Frame – Process button should now be visible
        bottom_controls_frame = ctk.CTkFrame(main_frame)
        bottom_controls_frame.pack(fill="x", side="bottom", padx=PAD_X, pady=(PAD_Y, 0))
        log_frame_outer = ctk.CTkFrame(bottom_controls_frame)
        log_frame_outer.pack(fill="both", expand=True, padx=0, pady=(0, PAD_Y))
        ctk.CTkLabel(log_frame_outer, text="Log", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=PAD_X, pady=(PAD_Y, 0))
        self.log_text = ctk.CTkTextbox(log_frame_outer, height=150, wrap="word", state="disabled", corner_radius=CORNER_RADIUS)
        self.log_text.pack(fill="both", expand=True, padx=PAD_X, pady=PAD_Y)
        progress_start_frame = ctk.CTkFrame(bottom_controls_frame, fg_color="transparent")
        progress_start_frame.pack(fill="x", padx=0, pady=PAD_Y)
        progress_start_frame.grid_columnconfigure(0, weight=1)
        self.progress_bar = ctk.CTkProgressBar(progress_start_frame, orientation="horizontal", height=20)
        self.progress_bar.grid(row=0, column=0, padx=(0, PAD_X), pady=PAD_Y, sticky="ew")
        self.progress_bar.set(0)
        self.start_button = ctk.CTkButton(
            progress_start_frame, text="Start Processing", command=self.start_processing,
            corner_radius=CORNER_RADIUS, height=30, font=ctk.CTkFont(weight="bold")
        )
        self.start_button.grid(row=0, column=1, padx=(PAD_X, 0), pady=PAD_Y)
        self.update_select_button_text()

    # --- GUI Update/Action Methods ---
    def add_event_id_entry(self, event=None):
        entry_frame = ctk.CTkFrame(self.conditions_scroll_frame, fg_color="transparent")
        entry_frame.pack(fill="x", pady=1, padx=1)
        entry = ctk.CTkEntry(entry_frame, width=100, corner_radius=CORNER_RADIUS)
        entry.pack(side="left", fill="x", expand=True, padx=(0, PAD_X))
        entry.bind("<Return>", self.add_event_id_entry)
        entry.bind("<KP_Enter>", self.add_event_id_entry)
        remove_btn = ctk.CTkButton(
            entry_frame, text="X", width=28, height=28, corner_radius=CORNER_RADIUS,
            command=lambda ef=entry_frame, e=entry: self.remove_event_id_entry(ef, e)
        )
        remove_btn.pack(side="right")
        self.condition_entries.append(entry)
        if event is None and entry.winfo_exists():
            entry.focus_set()

    def remove_event_id_entry(self, entry_frame, entry_widget):
        try:
            if entry_frame.winfo_exists():
                entry_frame.destroy()
            if entry_widget in self.condition_entries:
                self.condition_entries.remove(entry_widget)
            if not self.condition_entries:
                self.add_event_id_entry()
        except Exception as e:
            self.log(f"Error removing Event ID field: {e}")

    def select_save_folder(self):
        folder = filedialog.askdirectory(title="Select Parent Folder for Excel Output")
        if folder:
            self.save_folder_path.set(folder)
            self.log(f"Output folder: {folder}")
        else:
            self.log("Save folder selection cancelled.")

    def update_select_button_text(self):
        mode = self.file_mode.get()
        text = "Select EEG File..." if mode == "Single" else "Select Folder..."
        if hasattr(self, 'select_button') and self.select_button:
            self.select_button_text.set(text)

    def select_data_source(self):
        self.data_paths = []
        file_ext = "*" + self.file_type.get().lower()
        file_type_desc = self.file_type.get().upper()
        try:
            if self.file_mode.get() == "Single":
                ftypes = [(f"{file_type_desc} files", file_ext)]
                other_ext = "*.set" if file_type_desc == ".BDF" else "*.bdf"
                other_desc = ".SET" if file_type_desc == ".BDF" else ".BDF"
                ftypes.append((f"{other_desc} files", other_ext))
                ftypes.append(("All files", "*.*"))
                file_path = filedialog.askopenfilename(title="Select EEG File", filetypes=ftypes)
                if file_path:
                    selected_ext = os.path.splitext(file_path)[1].lower()
                    if selected_ext in ['.bdf', '.set']:
                        self.file_type.set(selected_ext.upper())
                        self.log(f"File type set to {selected_ext.upper()}")
                    self.data_paths = [file_path]
                    self.log(f"Selected file: {os.path.basename(file_path)}")
                else:
                    self.log("No file selected.")
            else:  # Batch mode
                folder = filedialog.askdirectory(title=f"Select Folder with {file_type_desc} Files")
                if folder:
                    search_path = os.path.join(folder, file_ext)
                    found_files = sorted(glob.glob(search_path))
                    if found_files:
                        self.data_paths = found_files
                        self.log(f"Selected folder: {folder}, Found {len(found_files)} '{file_ext}' file(s).")
                    else:
                        self.log(f"No '{file_ext}' files found in {folder}.")
                        messagebox.showwarning("No Files Found", f"No '{file_type_desc}' files found in:\n{folder}")
                else:
                    self.log("No folder selected.")
        except Exception as e:
            self.log(f"Error selecting data: {e}")
            messagebox.showerror("Selection Error", f"Error during selection:\n{e}")
        self._max_progress = len(self.data_paths) if self.data_paths else 1
        self.progress_bar.set(0)

    # --- Logging ---
    def log(self, message):
        if hasattr(self, 'log_text') and self.log_text:
            try:
                ct = threading.current_thread()
                ts = pd.Timestamp.now().strftime('%H:%M:%S.%f')[:-3]
                prefix = "[BG]" if ct != threading.main_thread() else "[GUI]"
                log_msg = f"{ts} {prefix}: {message}\n"

                def update_gui():
                    if hasattr(self, 'log_text') and self.log_text.winfo_exists():
                        self.log_text.configure(state="normal")
                        self.log_text.insert(tk.END, log_msg)
                        self.log_text.see(tk.END)
                        self.log_text.configure(state="disabled")

                if ct != threading.main_thread():
                    if hasattr(self, 'after') and self.winfo_exists():
                        self.after(0, update_gui)
                    print(log_msg, end='')
                else:
                    update_gui()
                    self.update_idletasks()
            except Exception as e:
                print(f"--- GUI Log Error: {e} ---\n{pd.Timestamp.now().strftime('%H:%M:%S.%f')[:-3]} Log Console: {message}")

    # --- Event ID Detection (Background Thread) ---
    def detect_and_populate_event_ids(self):
        self.log("Detect Event IDs button clicked...")
        if self.detection_thread and self.detection_thread.is_alive():
            messagebox.showwarning("Busy", "Event detection is already running.")
            return
        if not self.data_paths:
            messagebox.showerror("No Data Selected", "Please select a data file or folder first.")
            self.log("Detection failed: No data selected.")
            return
        representative_file = self.data_paths[0]
        self.log(f"Starting background detection for: {os.path.basename(representative_file)}")
        self._disable_controls(enable_process_buttons=None)
        try:
            self.detection_thread = threading.Thread(
                target=self._detection_thread_func,
                args=(representative_file, self.gui_queue),
                daemon=True
            )
            self.detection_thread.start()
            self.after(100, self._periodic_detection_queue_check)
        except Exception as start_err:
            self.log(f"Error starting detection thread: {start_err}")
            messagebox.showerror("Thread Error", f"Could not start detection thread:\n{start_err}")
            self._enable_controls(enable_process_buttons=None)

    def _detection_thread_func(self, file_path, gui_queue):
        """ Background task: load file and extract events from annotations. """
        raw = None
        gc.collect()
        try:
            raw = self.load_eeg_file(file_path)
            if raw is None:
                raise ValueError("File loading failed (check log).")
            # Use annotations for event extraction
            events, event_dict = mne.events_from_annotations(raw)
            if events is None or len(events) == 0:
                gui_queue.put({'type': 'log', 'message': "Info: No events found via annotations."})
                detected_labels = []
            else:
                gui_queue.put({'type': 'log', 'message': f"Found {len(events)} events. Event types: {list(event_dict.keys())}"})
                detected_labels = list(event_dict.keys())
            gui_queue.put({'type': 'detection_result', 'ids': detected_labels})
        except Exception as e:
            error_msg = f"Error during event detection: {e}"
            gui_queue.put({'type': 'log', 'message': f"!!! {error_msg}\n{traceback.format_exc()}"})
            gui_queue.put({'type': 'detection_error', 'message': error_msg})
        finally:
            if raw:
                del raw
                gc.collect()
            gui_queue.put({'type': 'detection_done'})

    def _periodic_detection_queue_check(self):
        """ Checks queue for results from the detection thread. """
        detection_finished = False
        try:
            while True:
                message = self.gui_queue.get_nowait()
                msg_type = message.get('type')
                if msg_type == 'log':
                    self.log(message.get('message', ''))
                elif msg_type == 'detection_result':
                    detected_ids = message.get('ids', [])
                    self._clear_and_reset_event_id_fields()
                    if detected_ids:
                        self.log("Populating fields with detected IDs...")
                        for label in detected_ids:
                            self.add_event_id_entry()
                            self.condition_entries[-1].delete(0, tk.END)
                            self.condition_entries[-1].insert(0, label)
                        messagebox.showinfo("Event IDs Detected",
                                            f"Populated list with {len(detected_ids)} unique Event ID(s).\nPlease review.")
                        if self.condition_entries and self.condition_entries[0].winfo_exists():
                            self.condition_entries[0].focus_set()
                    else:
                        messagebox.showinfo("No Events Found", "No event triggers found.\nPlease enter IDs manually.")
                    detection_finished = True
                elif msg_type == 'detection_error':
                    error_msg = message.get('message', 'Unknown error.')
                    if "'Status' channel not found" in error_msg:
                        messagebox.showwarning("Channel Not Found", "'Status' channel not found.\nEnter IDs manually.")
                    else:
                        messagebox.showerror("Detection Error", error_msg)
                    self._clear_and_reset_event_id_fields()
                    detection_finished = True
                elif msg_type == 'detection_done':
                    self.log("Detection thread finished.")
                    detection_finished = True
        except queue.Empty:
            pass
        except Exception as e:
            self.log(f"!!! Queue Check Error: {e}")
            messagebox.showerror("GUI Error", f"Error processing detection results:\n{e}")
            detection_finished = True

        if detection_finished:
            self._enable_controls(enable_process_buttons=None)
            self.detection_thread = None
            gc.collect()
            self.log("Detection process finished.")
        elif self.detection_thread and self.detection_thread.is_alive():
            self.after(100, self._periodic_detection_queue_check)
        else:
            self.log("Warn: Detection thread ended unexpectedly.")
            self._enable_controls(enable_process_buttons=None)
            self.detection_thread = None

    def _clear_and_reset_event_id_fields(self):
        self.log("Clearing Event ID fields...")
        try:
            for widget in self.conditions_scroll_frame.winfo_children():
                widget.destroy()
            self.condition_entries = []
            self.add_event_id_entry()
        except Exception as e:
            self.log(f"Error resetting Event ID fields: {e}")

    # --- Core Processing Control ---
    def start_processing(self):
        if self.processing_thread and self.processing_thread.is_alive():
            messagebox.showwarning("Busy", "Processing is already running.")
            return
        if self.detection_thread and self.detection_thread.is_alive():
            messagebox.showwarning("Busy", "Event detection is running.")
            return
        self.log("=" * 50)
        self.log("START PROCESSING Initiated...")
        if not self._validate_inputs():
            return
        self.preprocessed_data = {}
        self.progress_bar.set(0)
        self._max_progress = len(self.data_paths)
        self._disable_controls(enable_process_buttons=False)
        self.log("Starting background processing thread...")
        thread_args = (list(self.data_paths), self.validated_params.copy(),
                       list(self.current_event_ids_process), self.gui_queue)
        self.processing_thread = threading.Thread(
            target=self._processing_thread_func,
            args=thread_args,
            daemon=True
        )
        self.processing_thread.start()
        self.after(100, self._periodic_queue_check)

    # --- Input Validation ---
    def _validate_inputs(self):
        """ Validates file selection, folder, and parameters. """
        if not self.data_paths:
            self.log("V-Error: No data.")
            messagebox.showerror("Input Error", "No data selected.")
            return False
        save_folder = self.save_folder_path.get()
        if not save_folder:
            self.log("V-Error: No save folder.")
            messagebox.showerror("Input Error", "No output folder selected.")
            return False
        if not os.path.isdir(save_folder):
            try:
                os.makedirs(save_folder, exist_ok=True)
                self.log(f"Created save folder: {save_folder}")
            except Exception as e:
                self.log(f"V-Error: Cannot create folder {save_folder}: {e}")
                messagebox.showerror("Input Error", f"Cannot create folder:\n{save_folder}\n{e}")
                return False

        params = {}
        try:
            def get_float(e):
                s = e.get().strip()
                return float(s) if s else None

            def get_int(e):
                s = e.get().strip()
                return int(s) if s else None

            params['low_pass'] = get_float(self.low_pass_entry)
            assert params['low_pass'] is None or params['low_pass'] >= 0, "Low Pass >= 0"
            params['high_pass'] = get_float(self.high_pass_entry)
            assert params['high_pass'] is None or params['high_pass'] > 0, "High Pass > 0"
            params['downsample_rate'] = get_float(self.downsample_entry)
            assert params['downsample_rate'] is None or params['downsample_rate'] > 0, "Downsample > 0"
            params['epoch_start'] = get_float(self.epoch_start_entry)
            assert params['epoch_start'] is not None, "Epoch Start required"
            params['epoch_end'] = get_float(self.epoch_end_entry)
            assert params['epoch_end'] is not None, "Epoch End required"
            assert params['epoch_start'] < params['epoch_end'], "Epoch Start < End"
            params['reject_thresh'] = get_float(self.reject_thresh_entry)
            assert params['reject_thresh'] is None or params['reject_thresh'] > 0, "Reject Z > 0"
            # Read reference channels as strings
            ref1 = self.ref_channel1_entry.get().strip()
            ref2 = self.ref_channel2_entry.get().strip()
            params['ref_channel1'] = ref1
            params['ref_channel2'] = ref2
            params['max_idx_keep'] = get_int(self.max_idx_keep_entry)
            assert params['max_idx_keep'] is None or params['max_idx_keep'] > 0, "Max Keep Idx > 0"
            if (params['low_pass'] is not None) and (params['high_pass'] is not None):
                assert params['low_pass'] < params['high_pass'], "Low Pass < High Pass"
            params['save_preprocessed'] = self.save_preprocessed.get()
            self.validated_params = params
        except Exception as e:
            self.log(f"V-Error: Invalid parameter: {e}")
            messagebox.showerror("Parameter Error", f"Invalid parameter value:\n{e}")
            return False

        raw_ids = [e.get().strip() for e in self.condition_entries if e.get().strip()]
        if not raw_ids:
            self.log("V-Error: No Event IDs.")
            messagebox.showerror("Event ID Error", "Enter at least one Event ID.")
            return False
        # Use the text strings as condition identifiers
        self.current_event_ids_process = sorted(list(set(raw_ids)))
        self.log("Inputs Validated.")
        self.log(f"Params: {self.validated_params}")
        self.log(f"Event IDs: {self.current_event_ids_process}")
        return True

    # --- Periodic Queue Check (Main Processing) ---
    def _periodic_queue_check(self):
        processing_done = False
        final_success = True
        try:
            while True:
                message = self.gui_queue.get_nowait()
                msg_type = message.get('type')
                if msg_type in ['detection_result', 'detection_error', 'detection_done']:
                    continue
                elif msg_type == 'log':
                    self.log(message.get('message', ''))
                elif msg_type == 'progress':
                    value = message.get('value', 0)
                    self.progress_bar.set(value / self._max_progress if self._max_progress > 0 else 0)
                elif msg_type == 'result':
                    self.preprocessed_data = message.get('data', {})
                    self.log("Preprocessing results received.")
                elif msg_type == 'error':
                    error_msg = message.get('message', 'Err')
                    tb = message.get('traceback', '')
                    self.log(f"!!! THREAD ERROR: {error_msg}")
                    messagebox.showerror("Processing Error", error_msg)
                    processing_done = True
                    final_success = False
                    print(tb)
                elif msg_type == 'done':
                    self.log("BG thread done.")
                    processing_done = True
                else:
                    self.log(f"Warn: Unknown queue msg: {msg_type}")
        except queue.Empty:
            pass
        except Exception as e:
            self.log(f"!!! Queue Check Error: {e}")
            processing_done = True
            final_success = False
            print(traceback.format_exc())

        if processing_done:
            self._finalize_processing(final_success)
        elif self.processing_thread and self.processing_thread.is_alive():
            self.after(100, self._periodic_queue_check)
        else:
            self.log("Warn: Processing thread ended unexpectedly.")
            self._finalize_processing(False)

    # --- Finalize Processing ---
    def _finalize_processing(self, success):
        self.progress_bar.set(1.0 if success else self.progress_bar.get())
        if success and self.preprocessed_data:
            has_data = any(bool(epochs_list) for epochs_list in self.preprocessed_data.values())
            if has_data:
                self.log("\n--- Starting Post-processing Phase ---")
                try:
                    self.post_process(list(self.preprocessed_data.keys()))
                except Exception as post_err:
                    self.log(f"!!! Post-processing Error: {post_err}\n{traceback.format_exc()}")
                    messagebox.showerror("Post-processing Error", f"Error during analysis/saving:\n{post_err}")
            else:
                self.log("--- Skipping Post-processing: No usable epochs found ---")
                messagebox.showwarning("Processing Finished", "Preprocessing OK, but no usable epochs found.\nNo Excel files generated.")
        elif success:
            self.log("--- Skipping Post-processing: Preprocessing OK, but no data retained ---")
            messagebox.showinfo("Processing Finished", "Preprocessing OK, but no data available.\nNo Excel files generated.")
        else:
            self.log("--- Post-processing skipped due to errors ---")
        self._enable_controls(enable_process_buttons=True)
        self.log(f"--- Processing Run Finished at {pd.Timestamp.now()} ---")
        self.processing_thread = None
        gc.collect()

    # --- Disable/Enable Controls ---
    def _disable_controls(self, enable_process_buttons=None):
        widgets = []
        if hasattr(self, 'select_button'):
            widgets.append(self.select_button)
        if hasattr(self, 'detect_button'):
            widgets.append(self.detect_button)
        if hasattr(self, 'add_cond_button'):
            widgets.append(self.add_cond_button)
        if hasattr(self, 'radio_single'):
            widgets.append(self.radio_single)
        if hasattr(self, 'radio_batch'):
            widgets.append(self.radio_batch)
        if hasattr(self, 'options_frame'):
            for w in self.options_frame.winfo_children():
                if isinstance(w, ctk.CTkRadioButton) and w not in [self.radio_single, self.radio_batch]:
                    widgets.append(w)
        p_entries = [getattr(self, n, None) for n in [
            'low_pass_entry', 'high_pass_entry', 'downsample_entry',
            'epoch_start_entry', 'epoch_end_entry', 'reject_thresh_entry',
            'ref_channel1_entry', 'ref_channel2_entry', 'max_idx_keep_entry'
        ]]
        widgets.extend([e for e in p_entries if e])
        if hasattr(self, 'params_frame'):
            for w in self.params_frame.winfo_children():
                if isinstance(w, ctk.CTkCheckBox):
                    widgets.append(w)
        for entry in self.condition_entries:
            if entry and entry.winfo_exists():
                widgets.append(entry)
                p_frame = entry.master
                if p_frame.winfo_exists():
                    for w in p_frame.winfo_children():
                        if isinstance(w, ctk.CTkButton):
                            widgets.append(w)
        if hasattr(self, 'save_frame'):
            for w in self.save_frame.winfo_children():
                if isinstance(w, ctk.CTkButton):
                    widgets.append(w)
        if hasattr(self, 'start_button') and enable_process_buttons is False:
            widgets.append(self.start_button)
        for w in widgets:
            if w and w.winfo_exists():
                try:
                    w.configure(state="disabled")
                except Exception:
                    pass
        self.update_idletasks()

    def _enable_controls(self, enable_process_buttons=None):
        widgets = []
        if hasattr(self, 'select_button'):
            widgets.append(self.select_button)
        if hasattr(self, 'detect_button'):
            widgets.append(self.detect_button)
        if hasattr(self, 'add_cond_button'):
            widgets.append(self.add_cond_button)
        if hasattr(self, 'radio_single'):
            widgets.append(self.radio_single)
        if hasattr(self, 'radio_batch'):
            widgets.append(self.radio_batch)
        if hasattr(self, 'options_frame'):
            for w in self.options_frame.winfo_children():
                if isinstance(w, ctk.CTkRadioButton) and w not in [self.radio_single, self.radio_batch]:
                    widgets.append(w)
        p_entries = [getattr(self, n, None) for n in [
            'low_pass_entry', 'high_pass_entry', 'downsample_entry',
            'epoch_start_entry', 'epoch_end_entry', 'reject_thresh_entry',
            'ref_channel1_entry', 'ref_channel2_entry', 'max_idx_keep_entry'
        ]]
        widgets.extend([e for e in p_entries if e])
        if hasattr(self, 'params_frame'):
            for w in self.params_frame.winfo_children():
                if isinstance(w, ctk.CTkCheckBox):
                    widgets.append(w)
        if hasattr(self, 'conditions_scroll_frame'):
            for entry_frame in self.conditions_scroll_frame.winfo_children():
                if isinstance(entry_frame, ctk.CTkFrame) and entry_frame.winfo_exists():
                    for w in entry_frame.winfo_children():
                        if w and w.winfo_exists():
                            widgets.append(w)
        if hasattr(self, 'save_frame'):
            for w in self.save_frame.winfo_children():
                if isinstance(w, ctk.CTkButton):
                    widgets.append(w)
        if hasattr(self, 'start_button') and enable_process_buttons is True:
            widgets.append(self.start_button)
        for w in widgets:
            if w and w.winfo_exists():
                try:
                    w.configure(state="normal")
                except Exception:
                    pass
        self.update_idletasks()

    # --- Background Processing Thread Function ---
    def _processing_thread_func(self, data_paths, params, conditions_ids_to_process, gui_queue):
        local_data = {cond: [] for cond in conditions_ids_to_process}
        files_w_epochs = 0
        gc.collect()
        try:
            n_files = len(data_paths)
            for i, f_path in enumerate(data_paths):
                f_name = os.path.basename(f_path)
                gui_queue.put({'type': 'log', 'message': f"\nProcessing file {i+1}/{n_files}: {f_name}"})
                raw, raw_proc, evts = None, None, None
                try:
                    raw = self.load_eeg_file(f_path)
                    if raw is None:
                        continue
                    # After pre-processing, extract events using annotations.
                    raw_proc = self.preprocess_raw(raw.copy(), **params)
                    if raw_proc is None:
                        continue
                    events, event_dict = mne.events_from_annotations(raw_proc)
                    if events is None or len(events) == 0:
                        gui_queue.put({'type': 'log', 'message': "Info: No events found via annotations."})
                        evts = None
                    else:
                        gui_queue.put({'type': 'log', 'message': f"Found {len(events)} events. Event types: {list(event_dict.keys())}"})
                        evts = (events, event_dict)
                    file_epochs = False
                    if evts is not None:
                        events, event_dict = evts
                        for cond in conditions_ids_to_process:
                            if cond in event_dict:
                                try:
                                    # Create epochs for the condition using the text label.
                                    epochs = mne.Epochs(raw_proc, events, event_id={cond: event_dict[cond]},
                                                        tmin=params['epoch_start'], tmax=params['epoch_end'],
                                                        preload=False, verbose=False, baseline=None, on_missing='warn')
                                    if len(epochs.events) > 0:
                                        gui_queue.put({'type': 'log', 'message': f"  Found {len(epochs.events)} epochs for condition {cond}."})
                                        local_data[cond].append(epochs)
                                        file_epochs = True
                                except Exception as ep_err:
                                    gui_queue.put({'type': 'log', 'message': f"Epoch error for condition {cond}: {ep_err}\n{traceback.format_exc()}"})
                        if file_epochs:
                            files_w_epochs += 1
                    else:
                        gui_queue.put({'type': 'log', 'message': "Skipping epochs (no events)."})
                    if params['save_preprocessed']:
                        # Save with a name conforming to MNE conventions:
                        p_path = os.path.join(os.path.dirname(f_path), f"{os.path.splitext(f_name)[0]}_preproc_raw.fif")
                        try:
                            raw_proc.save(p_path, overwrite=True, verbose=False)
                        except Exception as s_err:
                            gui_queue.put({'type': 'log', 'message': f"Warn: Save failed: {s_err}"})
                except MemoryError as mem_err:
                    gui_queue.put({'type': 'error', 'message': f"Memory Error {f_name}: {mem_err}",
                                   'traceback': traceback.format_exc()})
                    return
                except Exception as f_err:
                    gui_queue.put({'type': 'log', 'message': f"!!! FILE ERROR {f_name}: {f_err}\n{traceback.format_exc()}"})
                finally:
                    del raw, raw_proc, evts
                    gc.collect()
                    gui_queue.put({'type': 'progress', 'value': i + 1})
            gui_queue.put({'type': 'log', 'message': f"\n--- BG Preprocessing Done ({files_w_epochs} files with epochs) ---"})
            gui_queue.put({'type': 'result', 'data': local_data})
        except MemoryError as mem_err:
            gui_queue.put({'type': 'error', 'message': f"Critical Memory Error: {mem_err}",
                           'traceback': traceback.format_exc()})
        except Exception as e:
            gui_queue.put({'type': 'error', 'message': f"Critical thread error: {e}",
                           'traceback': traceback.format_exc()})
        finally:
            gui_queue.put({'type': 'done'})

    # --- EEG Loading Method ---
    def load_eeg_file(self, filepath):
        """Loads BDF or SET file using MNE."""
        ext = os.path.splitext(filepath)[1].lower()
        raw = None
        base_filename = os.path.basename(filepath)
        self.log(f"Loading: {base_filename}...")
        try:
            load_kwargs = {'preload': True, 'verbose': False}
            if ext == ".bdf":
                try:
                    self.log("Attempting BDF load with stim_channel='Status'.")
                    with mne.utils.use_log_level('WARNING'):
                        raw = mne.io.read_raw_bdf(filepath, stim_channel='Status', **load_kwargs)
                    self.log("BDF loaded successfully with 'Status'.")
                except ValueError as ve:
                    if "could not find stim channel" in str(ve).lower() and "'status'" in str(ve).lower():
                        self.log("Warning: 'Status' channel not found. Attempting load without specifying...")
                        try:
                            with mne.utils.use_log_level('WARNING'):
                                raw = mne.io.read_raw_bdf(filepath, **load_kwargs)
                            self.log("Loaded BDF without 'Status' (event detection may fail).")
                        except Exception as fallback_load_err:
                            self.log(f"Error loading BDF fallback: {fallback_load_err}")
                            messagebox.showerror("Loading Error", f"Could not load BDF file:\n{base_filename}\n\nError: {fallback_load_err}")
                            return None
                    else:
                        raise ve
                except Exception as bdf_err:
                    raise bdf_err
            elif ext == ".set":
                self.log("Attempting SET load.")
                with mne.utils.use_log_level('WARNING'):
                    raw = mne.io.read_raw_eeglab(filepath, **load_kwargs)
                self.log("SET loaded successfully.")
            else:
                self.log(f"Unsupported format '{ext}'.")
                messagebox.showwarning("Unsupported File", f"Format '{ext}' not supported.")
                return None

            if raw is None:
                raise ValueError("MNE load returned None unexpectedly.")
            self.log(f"Load OK: {len(raw.ch_names)} channels @ {raw.info['sfreq']:.1f} Hz.")
            self.log("Applying standard_1020 montage...")
            try:
                montage = mne.channels.make_standard_montage('standard_1020')
                raw.set_montage(montage, on_missing='warn', match_case=False)
                self.log("Montage applied (check warnings for missing channels like EXG).")
            except Exception as m_err:
                self.log(f"Warning: Montage error: {m_err}")
            return raw
        except MemoryError:
            self.log(f"!!! Memory Error loading {base_filename}.")
            messagebox.showerror("Memory Error", f"Memory Error loading {base_filename}.")
            return None
        except ValueError as ve:
            self.log(f"Value Error loading {base_filename}: {ve}")
            messagebox.showerror("Loading Error", f"Could not load:\n{base_filename}\nValue Error: {ve}")
            return None
        except FileNotFoundError:
            self.log(f"!!! File Not Found Error: {filepath}")
            messagebox.showerror("Loading Error", f"File not found:\n{filepath}")
            return None
        except Exception as e:
            self.log(f"!!! General Load Error {base_filename}: {e}\n{traceback.format_exc()}")
            messagebox.showerror("Loading Error", f"Could not load:\n{base_filename}\nError: {e}")
            return None

    # --- Preprocessing Method ---
    def preprocess_raw(self, raw, **params):
        downsample_rate = params.get('downsample_rate')
        low_pass = params.get('low_pass')
        high_pass = params.get('high_pass')
        reject_thresh = params.get('reject_thresh')
        ref_channel1 = params.get('ref_channel1')
        ref_channel2 = params.get('ref_channel2')
        max_idx_keep = params.get('max_idx_keep')
        try:
            ch_names_orig = list(raw.info['ch_names'])
            n_chans_orig = len(ch_names_orig)
            self.log(f"Preprocessing {n_chans_orig} chans...")
            # Apply bipolar reference using string labels.
            if ref_channel1 and ref_channel2:
                if ref_channel1 in ch_names_orig and ref_channel2 in ch_names_orig:
                    new_ch = f"{ref_channel1}-{ref_channel2}"
                    try:
                        self.log(f"Applying bipolar ref: {ref_channel1}-{ref_channel2}...")
                        # Use set_bipolar_reference with string inputs.
                        raw.set_bipolar_reference(ref_channel1, ref_channel2, new_ch,
                                                  drop_refs=False, copy=False, verbose=False)
                        self.log(f"OK. New channel '{new_ch}'.")
                        ch_names_orig = list(raw.info['ch_names'])
                        n_chans_orig = len(ch_names_orig)
                    except Exception as bipol_err:
                        self.log(f"Warn: Bipolar ref failed: {bipol_err}.")
                else:
                    self.log(f"Warn: One or both reference channels ({ref_channel1}, {ref_channel2}) not found. Skipping bipolar ref.")
            else:
                self.log("Skip bipolar ref.")

            # Drop channels while preserving any channel named "Status"
            if max_idx_keep is not None:
                c_names = list(raw.info['ch_names'])
                c_n = len(c_names)
                status_chans = [ch for ch in c_names if ch.lower() == 'status']
                if 0 < max_idx_keep < c_n:
                    # Retain first max_idx_keep channels plus any "Status" channels (without duplication)
                    to_keep = list(dict.fromkeys([c_names[i] for i in range(max_idx_keep)] + status_chans))
                    to_drop = [ch for ch in c_names if ch not in to_keep]
                    if to_drop:
                        self.log(f"Dropping {len(to_drop)} chans (keeping first {max_idx_keep} and 'Status' channel(s))...")
                        try:
                            raw.drop_channels(to_drop)
                            self.log(f"OK. Remaining: {len(raw.ch_names)}")
                        except Exception as drop_err:
                            self.log(f"Warn: Drop failed: {drop_err}")
                elif max_idx_keep >= c_n:
                    self.log("Info: Max Idx Keep >= chans. No drop.")
            else:
                self.log("Skip index drop.")

            # Downsampling
            if downsample_rate:
                c_sf = raw.info['sfreq']
                self.log(f"Downsample check (Tgt: {downsample_rate}Hz). Curr: {c_sf:.1f}Hz.")
                if c_sf > downsample_rate:
                    try:
                        self.log("Downsampling...")
                        raw.resample(downsample_rate, npad="auto", verbose=False)
                        self.log(f"OK. New rate: {raw.info['sfreq']:.1f}Hz.")
                    except Exception as ds_err:
                        self.log(f"!!! ERROR Downsampling: {ds_err}. Stop.")
                        return None
                else:
                    self.log("No downsampling needed.")
            else:
                self.log("Skip downsample.")

            # Filtering
            l = low_pass if low_pass and low_pass > 0 else None
            h = high_pass
            if l or h:
                try:
                    self.log(f"Filtering ({l if l else 'DC'}-{h if h else 'Nyq'}Hz)...")
                    raw.filter(l, h, method='fir', phase='zero-double', fir_window='hamming',
                               fir_design='firwin', pad='edge', verbose=False)
                    self.log("Filter OK.")
                except Exception as f_err:
                    self.log(f"Warn: Filter failed: {f_err}.")
            else:
                self.log("Skip filter.")

            # Kurtosis-based channel rejection & interpolation
            if reject_thresh:
                self.log(f"Kurtosis rejection (Z > {reject_thresh})...")
                orig_bads = list(raw.info['bads'])
                try:
                    picks = mne.pick_types(raw.info, eeg=True, exclude='bads')
                    if len(picks) >= 2:
                        d = raw.get_data(picks)
                        k = kurtosis(d, axis=1, fisher=True, bias=False)
                        del d
                        k = np.nan_to_num(k)
                        m_val = np.mean(k)
                        s_val = np.std(k)
                        bad_k = []
                        if s_val > 1e-9:
                            z = (k - m_val) / s_val
                            names = [raw.info['ch_names'][i] for i in picks]
                            bad_k = [names[i] for i, zv in enumerate(z) if abs(zv) > reject_thresh]
                        else:
                            self.log("Kurtosis std zero.")
                        if bad_k:
                            new_b = [c for c in bad_k if c not in raw.info['bads']]
                            raw.info['bads'].extend(new_b)
                            self.log(f"Bad by Kurt: {bad_k}. Bads: {raw.info['bads']}")
                        else:
                            self.log("No new bads by Kurtosis.")
                        if raw.info['bads']:
                            if raw.get_montage():
                                try:
                                    self.log(f"Interpolating {raw.info['bads']}...")
                                    raw.interpolate_bads(reset_bads=True, mode='accurate', verbose=False)
                                    self.log("Interp OK.")
                                except Exception as int_err:
                                    self.log(f"Warn: Interp failed: {int_err}. Bads remain: {raw.info['bads']}")
                            else:
                                self.log("Warn: No montage for interp.")
                    else:
                        self.log("Skip Kurtosis (<=1 EEG chan).")
                except Exception as kurt_err:
                    self.log(f"Warn: Kurtosis err: {kurt_err}.")
                    raw.info['bads'] = orig_bads
            else:
                self.log("Skip Kurtosis.")

            # Apply average reference
            try:
                self.log("Applying avg ref...")
                raw.set_eeg_reference('average', projection=True)
                self.log("Avg ref OK.")
            except Exception as avg_err:
                self.log(f"Warn: Avg ref failed: {avg_err}")
            self.log(f"Preproc OK. State: {len(raw.ch_names)} chans, {raw.info['sfreq']:.1f} Hz.")
            return raw
        except MemoryError:
            self.log("!!! Memory Error preprocessing.")
            return None
        except Exception as e:
            self.log(f"!!! CRITICAL preproc error: {e}")
            print(traceback.format_exc())
            return None

    # --- Post-processing Method ---
    def post_process(self, conditions_ids_present):
        self.log("--- Post-processing: Calculating Metrics & Saving Excel ---")
        parent_folder = self.save_folder_path.get()
        if not parent_folder or not os.path.isdir(parent_folder):
            self.log(f"Error: Invalid save folder: '{parent_folder}'")
            messagebox.showerror("Save Error", f"Invalid output folder:\n{parent_folder}")
            return

        any_results_saved = False
        for cond in conditions_ids_present:
            epochs_list = self.preprocessed_data.get(cond, [])
            event_name = f"Event_{cond}"
            if not epochs_list:
                continue
            self.log(f"\nPost-processing {event_name} ({len(epochs_list)} file(s))...")
            accum = {'fft': None, 'snr': None, 'z': None, 'bca': None}
            valid_count = 0
            n_ch, ch_names = None, None
            for file_idx, epochs in enumerate(epochs_list):
                self.log(f"  Processing File {file_idx+1}/{len(epochs_list)}...")
                try:
                    if not isinstance(epochs, mne.BaseEpochs) or len(epochs.events) == 0:
                        self.log("    Invalid/empty Epochs. Skipping.")
                        continue
                    self.log(f"    Loading {len(epochs.events)} potential epochs...")
                    epochs.load_data()
                    picks = mne.pick_types(epochs.info, eeg=True, meg=False, stim=False, exclude='bads')
                    if not picks.size:
                        self.log("    No good EEG channels. Skipping.")
                        continue
                    ep_d = epochs.get_data(picks=picks, copy=False)
                    n_ep, n_c, n_t = ep_d.shape
                    sfreq = epochs.info['sfreq']
                    names = [epochs.info['ch_names'][i] for i in picks]
                    dur = n_t / sfreq
                    self.log(f"    Processing {n_ep} epochs, {n_c} channels @ {sfreq:.1f}Hz. Duration: {dur:.2f}s")
                    if n_ch is None:
                        n_ch = n_c
                        ch_names = names
                        self.log(f"    Setting expected chan count={n_ch}.")
                    elif n_ch != n_c or ch_names != names:
                        self.log("    !!! Warning: Channel mismatch! Skipping file.")
                        continue

                    fft_win_sec = 8.0
                    n_fft = int(sfreq * fft_win_sec)
                    if n_t < n_fft:
                        self.log(f"    Warning: Epoch duration ({dur:.1f}s) < FFT window ({fft_win_sec:.1f}s).")
                        n_fft = n_t
                    n_overlap = int(n_fft * 0.50)
                    fmin = epochs.info['lowpass'] if epochs.info.get('lowpass', 0) > 0 else 0.1
                    fmax = epochs.info['highpass']
                    self.log(f"    Calculating PSD (Welch: {fmin:.1f}-{fmax:.1f} Hz, N_FFT={n_fft})...")
                    spec_d = epochs.compute_psd(method='welch', fmin=fmin, fmax=fmax,
                                                  n_fft=n_fft, n_overlap=n_overlap, window='hann', average='mean', verbose=False)
                    pow_d = spec_d.get_data(False)
                    freqs = spec_d.freqs
                    amp_d = np.sqrt(pow_d)
                    self.log("    Calculating metrics...")
                    n_tf = len(TARGET_FREQUENCIES)
                    f_fft = np.zeros((n_c, n_tf))
                    f_snr = np.zeros((n_c, n_tf))
                    f_z = np.zeros((n_c, n_tf))
                    f_bca = np.zeros((n_c, n_tf))
                    noise_r, noise_e = 12, 1
                    for c_idx in range(n_c):
                        for f_idx, t_freq in enumerate(TARGET_FREQUENCIES):
                            t_bin = np.argmin(np.abs(freqs - t_freq))
                            l_b = max(0, t_bin - noise_r)
                            u_b = min(len(freqs), t_bin + noise_r + 1)
                            e_s = max(0, t_bin - noise_e)
                            e_e = min(len(freqs), t_bin + noise_e + 1)
                            n_idx = np.unique(np.concatenate([np.arange(l_b, e_s), np.arange(e_e, u_b)]))
                            n_idx = n_idx[n_idx < len(freqs)]
                            if n_idx.size >= 4:
                                n_m = np.mean(amp_d[c_idx, n_idx])
                                n_s = np.std(amp_d[c_idx, n_idx])
                            else:
                                n_m, n_s = 0, 0
                            f_val = amp_d[c_idx, t_bin]
                            s_val = f_val / n_m if n_m > 1e-12 else 0
                            l_s = max(0, t_bin - 1)
                            l_e = min(len(freqs), t_bin + 2)
                            l_max = np.max(amp_d[c_idx, l_s:l_e]) if l_s < l_e else f_val
                            z_val = (l_max - n_m) / n_s if n_s > 1e-12 else 0
                            b_val = f_val - n_m
                            f_fft[c_idx, f_idx] = f_val
                            f_snr[c_idx, f_idx] = s_val
                            f_z[c_idx, f_idx] = z_val
                            f_bca[c_idx, f_idx] = b_val
                    if accum['fft'] is None:
                        accum = {'fft': f_fft, 'snr': f_snr, 'z': f_z, 'bca': f_bca}
                    else:
                        accum['fft'] += f_fft
                        accum['snr'] += f_snr
                        accum['z'] += f_z
                        accum['bca'] += f_bca
                    valid_count += 1
                    self.log("    Accumulated metrics.")
                except MemoryError:
                    self.log("!!! Memory Error post-processing. Skipping file.")
                    continue
                except Exception as e:
                    self.log(f"!!! Error post-processing: {e}\n{traceback.format_exc()}")
                    continue
                finally:
                    del ep_d, spec_d, pow_d, amp_d
                    gc.collect()
            if valid_count > 0:
                self.log(f"Averaging metrics across {valid_count} files for {event_name}.")
                avg = {k: v / valid_count for k, v in accum.items()}
                cols = [f"{f:.1f}_Hz" for f in TARGET_FREQUENCIES]
                elecs = DEFAULT_ELECTRODE_NAMES_64 if n_ch == 64 else [f"Ch{i+1}" for i in range(n_ch)]
                if n_ch != 64:
                    self.log("Warn: Chan count != 64. Using generic names.")
                dfs = {
                    'FFT_Amplitude': pd.DataFrame(avg['fft'], index=elecs, columns=cols),
                    'SNR': pd.DataFrame(avg['snr'], index=elecs, columns=cols),
                    'Z_Score': pd.DataFrame(avg['z'], index=elecs, columns=cols),
                    'BCA': pd.DataFrame(avg['bca'], index=elecs, columns=cols)
                }
                for df_n, df in dfs.items():
                    df.insert(0, "Electrode", df.index)
                sub_path = os.path.join(parent_folder, event_name)
                excel_path = os.path.join(sub_path, f"{event_name}_FPVS_Results.xlsx")
                try:
                    os.makedirs(sub_path, exist_ok=True)
                except OSError as e:
                    self.log(f"Warn: Subfolder error: {e}. Saving to parent.")
                    excel_path = os.path.join(parent_folder, f"{event_name}_FPVS_Results.xlsx")
                try:
                    self.log(f"Writing Excel: {excel_path}")
                    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                        for name, df in dfs.items():
                            df.to_excel(writer, sheet_name=name.replace('_', ' '), index=False)
                        wb = writer.book
                        cf = wb.add_format({'align': 'center', 'valign': 'vcenter'})
                        for sn in writer.sheets:
                            ws = writer.sheets[sn]
                            c_df = dfs.get(sn.replace(' ', '_'))
                            if c_df is None:
                                continue
                            for c_idx, c_name in enumerate(c_df.columns):
                                w = max(len(str(c_name)), c_df[c_name].astype(str).map(len).max() if not c_df[c_name].empty else 0) + 4
                                ws.set_column(c_idx, c_idx, w, cf)
                    self.log(f"Saved Excel for {event_name}.")
                    any_results_saved = True
                except Exception as ex_err:
                    self.log(f"!!! Excel Error {excel_path}: {ex_err}")
                    messagebox.showerror("Excel Error", f"Failed save for {event_name}.\n{ex_err}")
            else:
                self.log(f"No valid data for {event_name}. No Excel generated.")
        if any_results_saved:
            self.log("Post-processing complete. Results saved.")
        else:
            self.log("Post-processing complete. No results generated.")

# --- Main execution block ---
if __name__ == "__main__":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass  # Not on Windows or DPI awareness setting fails
    app = FPVSApp()
    app.mainloop()
