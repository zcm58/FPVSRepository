#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EEG FPVS Analysis GUI using MNE-Python and CustomTkinter.

Version: 1.20 (April 2025) - Revised to use mne.find_events for event extraction
from a stimulus channel using numerical IDs mapped to labels provided by the user.
Replaces annotation-based extraction from v1.10.

Key functionalities:
- Modern GUI using CustomTkinter.
- Load EEG data (.BDF, .set).
- Process single files or batch folders.
- Preprocessing Steps:
    - Specify Stimulus Channel Name (default: 'Status').
    - Initial Bipolar Reference using user-specified channels.
    - Downsample.
    - Apply standard_1020 montage.
    - Drop channels above a specified index (preserving Stim/Status channel).
    - Bandpass filter.
    - Kurtosis-based channel rejection & interpolation.
    - Average common reference.
- Extract epochs based on numerical triggers found via mne.find_events,
  using a user-provided mapping of Labels to Numerical IDs.
- Post-processing using FFT, SNR, Z-score, BCA computation.
- Saves Excel files with separate sheets per metric, named by condition label.
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
import re # For parsing event map

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
# Fixed parameters for post-processing
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
DEFAULT_STIM_CHANNEL = 'Status' # Default stimulus channel name

# =====================================================
# GUI Configuration
# =====================================================
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")
CORNER_RADIUS = 8
PAD_X = 5
PAD_Y = 5
ENTRY_WIDTH = 100
LABEL_ID_ENTRY_WIDTH = 120


class FPVSApp(ctk.CTk):
    """ Main application class replicating MATLAB FPVS analysis workflow using numerical triggers. """

    def __init__(self):
        super().__init__()
        today_str = pd.Timestamp.now().strftime('%Y-%m-%d')
        # --- Version Update ---
        self.title(f"EEG FPVS Analysis Tool (v1.20 - {today_str})")
        self.geometry("1000x950")  # initial geometry

        # Data structures
        self.preprocessed_data = {}
        # --- Event ID Input Change ---
        # self.condition_entries = [] # Replaced by event_map_entries
        self.event_map_entries = [] # List to hold tuples of (label_entry, id_entry, frame, button)
        self.current_event_map = {} # Will hold {'Label': ID} from GUI during processing
        self.data_paths = []
        self.processing_thread = None
        self.detection_thread = None
        self.gui_queue = queue.Queue()
        self._max_progress = 1
        self.validated_params = {}

        self.save_preprocessed = tk.BooleanVar(value=False)

        self.create_menu()
        self.create_widgets()
        self.log("Welcome to the EEG FPVS Analysis Tool (Numerical Trigger Version)!")
        self.log(f"Appearance Mode: {ctk.get_appearance_mode()}")

        self.update_idletasks()
        req_width = self.winfo_reqwidth()
        req_height = self.winfo_reqheight()
        self.geometry(f"{req_width}x{req_height}")

        # Focus first event label entry if it exists
        if self.event_map_entries:
            self.event_map_entries[0]['label'].focus_set()


    # --- Menu Methods (unchanged) ---
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
         # --- Version Update ---
        messagebox.showinfo("About EEG FPVS Analysis Tool",
                            f"Version: 1.20 ({pd.Timestamp.now().strftime('%B %Y')})\nUsing numerical triggers via mne.find_events.")

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
    def _validate_numeric_input(self, P): # Used for IDs and other params
        if P == "" or P == "-": # Allow empty or negative sign start
            return True
        try:
            float(P) # Check if convertible to float (allows integers too)
            return True
        except ValueError:
            self.bell() # System beep for invalid input
            return False

    def _validate_integer_input(self, P): # Specifically for Event IDs
        if P == "": # Allow empty
             return True
        try:
             int(P) # Must be an integer
             return True
        except ValueError:
             self.bell()
             return False

    # --- GUI Creation ---
    def create_widgets(self):
        """ Builds and arranges GUI components. """
        validate_num_cmd = (self.register(self._validate_numeric_input), '%P')
        # --- New validator for integer Event IDs ---
        validate_int_cmd = (self.register(self._validate_integer_input), '%P')

        main_frame = ctk.CTkFrame(self, corner_radius=0)
        main_frame.pack(fill="both", expand=True, padx=PAD_X*2, pady=PAD_Y*2)
        # --- Adjusted row weights if necessary ---
        main_frame.grid_rowconfigure(3, weight=1) # Row for event map frame
        main_frame.grid_columnconfigure(0, weight=1)

        # Options Frame (Mostly unchanged)
        self.options_frame = ctk.CTkFrame(main_frame)
        self.options_frame.pack(fill="x", padx=PAD_X, pady=PAD_Y)
        # ... (Labels, Radiobuttons for Mode/File Type, Select Button - remain the same) ...
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

        # Row 1: Filter Frequencies
        ctk.CTkLabel(self.params_frame, text="Low Pass (Hz):").grid(row=1, column=0, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.low_pass_entry = ctk.CTkEntry(self.params_frame, width=ENTRY_WIDTH, validate='key', validatecommand=validate_num_cmd, corner_radius=CORNER_RADIUS)
        self.low_pass_entry.insert(0, "0.1")
        self.low_pass_entry.grid(row=1, column=1, padx=PAD_X, pady=PAD_Y)
        ctk.CTkLabel(self.params_frame, text="High Pass (Hz):").grid(row=1, column=2, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.high_pass_entry = ctk.CTkEntry(self.params_frame, width=ENTRY_WIDTH, validate='key', validatecommand=validate_num_cmd, corner_radius=CORNER_RADIUS)
        self.high_pass_entry.insert(0, "50")
        self.high_pass_entry.grid(row=1, column=3, padx=PAD_X, pady=PAD_Y)

        # Row 2: Downsample & Epoch Times
        ctk.CTkLabel(self.params_frame, text="Downsample (Hz):").grid(row=2, column=0, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.downsample_entry = ctk.CTkEntry(self.params_frame, width=ENTRY_WIDTH, validate='key', validatecommand=validate_num_cmd, corner_radius=CORNER_RADIUS)
        self.downsample_entry.insert(0, "256")
        self.downsample_entry.grid(row=2, column=1, padx=PAD_X, pady=PAD_Y)
        ctk.CTkLabel(self.params_frame, text="Epoch Start (s):").grid(row=2, column=2, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.epoch_start_entry = ctk.CTkEntry(self.params_frame, width=ENTRY_WIDTH, validate='key', validatecommand=validate_num_cmd, corner_radius=CORNER_RADIUS)
        self.epoch_start_entry.insert(0, "-1")
        self.epoch_start_entry.grid(row=2, column=3, padx=PAD_X, pady=PAD_Y)
        ctk.CTkLabel(self.params_frame, text="Epoch End (s):").grid(row=2, column=4, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.epoch_end_entry = ctk.CTkEntry(self.params_frame, width=ENTRY_WIDTH, validate='key', validatecommand=validate_num_cmd, corner_radius=CORNER_RADIUS)
        self.epoch_end_entry.insert(0, "125")
        self.epoch_end_entry.grid(row=2, column=5, padx=PAD_X, pady=PAD_Y)

        # Row 3: Rejection & Save Option
        ctk.CTkLabel(self.params_frame, text="Rejection Z-Thresh:").grid(row=3, column=0, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.reject_thresh_entry = ctk.CTkEntry(self.params_frame, width=ENTRY_WIDTH, validate='key', validatecommand=validate_num_cmd, corner_radius=CORNER_RADIUS)
        self.reject_thresh_entry.insert(0, "5")
        self.reject_thresh_entry.grid(row=3, column=1, padx=PAD_X, pady=PAD_Y)
        self.save_preprocessed_checkbox = ctk.CTkCheckBox(self.params_frame, text="Save Preprocessed (.fif)", variable=self.save_preprocessed, onvalue=True, offvalue=False, corner_radius=CORNER_RADIUS)
        self.save_preprocessed_checkbox.grid(row=3, column=2, columnspan=2, padx=PAD_X, pady=PAD_Y, sticky="w")

        # Row 4: Reference Channels
        ctk.CTkLabel(self.params_frame, text="Reference Channel 1:").grid(row=4, column=0, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.ref_channel1_entry = ctk.CTkEntry(self.params_frame, width=ENTRY_WIDTH, corner_radius=CORNER_RADIUS)
        self.ref_channel1_entry.insert(0, "EXG1")
        self.ref_channel1_entry.grid(row=4, column=1, padx=PAD_X, pady=PAD_Y)
        ctk.CTkLabel(self.params_frame, text="Reference Channel 2:").grid(row=4, column=2, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.ref_channel2_entry = ctk.CTkEntry(self.params_frame, width=ENTRY_WIDTH, corner_radius=CORNER_RADIUS)
        self.ref_channel2_entry.insert(0, "EXG2")
        self.ref_channel2_entry.grid(row=4, column=3, padx=PAD_X, pady=PAD_Y)

        # Row 5: Max Channels & Stim Channel
        ctk.CTkLabel(self.params_frame, text="Max EEG Channels Keep:").grid(row=5, column=0, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.max_idx_keep_entry = ctk.CTkEntry(self.params_frame, width=ENTRY_WIDTH, validate='key', validatecommand=validate_int_cmd, corner_radius=CORNER_RADIUS) # Should be integer
        self.max_idx_keep_entry.insert(0, "64")
        self.max_idx_keep_entry.grid(row=5, column=1, padx=PAD_X, pady=PAD_Y)
        # --- New Stimulus Channel Entry ---
        ctk.CTkLabel(self.params_frame, text="Stimulus Channel Name:").grid(row=5, column=2, sticky="w", padx=PAD_X, pady=PAD_Y)
        self.stim_channel_entry = ctk.CTkEntry(self.params_frame, width=ENTRY_WIDTH, corner_radius=CORNER_RADIUS)
        self.stim_channel_entry.insert(0, DEFAULT_STIM_CHANNEL) # Default 'Status'
        self.stim_channel_entry.grid(row=5, column=3, padx=PAD_X, pady=PAD_Y)


        # --- Event ID Frame - Changed for Label:ID Mapping ---
        event_map_frame_outer = ctk.CTkFrame(main_frame)
        event_map_frame_outer.pack(fill="both", expand=True, padx=PAD_X, pady=PAD_Y)
        ctk.CTkLabel(
            event_map_frame_outer, text="Event Label to Numerical ID Mapping",
            font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w", padx=PAD_X, pady=(PAD_Y, 0))

        # Add headers for the columns
        header_frame = ctk.CTkFrame(event_map_frame_outer, fg_color="transparent")
        header_frame.pack(fill="x", padx=PAD_X, pady=(2, 0))
        ctk.CTkLabel(header_frame, text="Condition Label (for Output)", width=LABEL_ID_ENTRY_WIDTH*2, anchor="w").pack(side="left", padx=(0, PAD_X))
        ctk.CTkLabel(header_frame, text="Numerical ID (Trigger Code)", width=LABEL_ID_ENTRY_WIDTH, anchor="w").pack(side="left", padx=(PAD_X, PAD_X))
        # Placeholder on the right to align with remove button
        ctk.CTkLabel(header_frame, text="", width=28).pack(side="right", padx=(PAD_X, 0))


        self.event_map_scroll_frame = ctk.CTkScrollableFrame(event_map_frame_outer, label_text="")
        self.event_map_scroll_frame.pack(fill="both", expand=True, padx=PAD_X, pady=(0, PAD_Y))
        self.event_map_scroll_frame.grid_columnconfigure(0, weight=1) # Make label entry expand
        self.event_map_scroll_frame.grid_columnconfigure(1, weight=0) # ID entry fixed width
        self.event_map_scroll_frame.grid_columnconfigure(2, weight=0) # Button fixed width

        self.event_map_entries = [] # Reset list
        self.add_event_map_entry()  # Pre-populate one row

        # Buttons below the scrollable frame
        event_map_button_frame = ctk.CTkFrame(event_map_frame_outer, fg_color="transparent")
        event_map_button_frame.pack(fill="x", pady=(0, PAD_Y), padx=PAD_X)
        # --- Changed Detect Button Text ---
        self.detect_button = ctk.CTkButton(
            event_map_button_frame, text="Detect Numerical IDs",
            command=self.detect_and_show_event_ids, corner_radius=CORNER_RADIUS
        )
        self.detect_button.pack(side="left", padx=(0, PAD_X))
        self.add_map_button = ctk.CTkButton(
            event_map_button_frame, text="Add Label:ID Row",
            command=self.add_event_map_entry, corner_radius=CORNER_RADIUS
        )
        self.add_map_button.pack(side="left")
        # --- End Event ID Frame Change ---


        # Save Location Frame (Unchanged)
        self.save_frame = ctk.CTkFrame(main_frame)
        self.save_frame.pack(fill="x", padx=PAD_X, pady=PAD_Y)
        # ... (Label, Button, Entry - remain the same) ...
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


        # Bottom Controls Frame (Unchanged layout)
        bottom_controls_frame = ctk.CTkFrame(main_frame)
        bottom_controls_frame.pack(fill="both", expand=True, side="bottom", padx=PAD_X, pady=(PAD_Y, 0))
        # ... (Log Frame, Progress Bar, Start Button - remain the same) ...
        log_frame_outer = ctk.CTkFrame(bottom_controls_frame)
        log_frame_outer.pack(fill="both", expand=True, padx=0, pady=(0, PAD_Y))
        ctk.CTkLabel(log_frame_outer, text="Log", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=PAD_X, pady=(PAD_Y, 0))
        self.log_text = ctk.CTkTextbox(log_frame_outer, height=150, wrap="word", state="disabled", corner_radius=CORNER_RADIUS)
        self.log_text.pack(fill="both", expand=True, padx=PAD_X, pady=PAD_Y)

        progress_start_frame = ctk.CTkFrame(bottom_controls_frame, fg_color="transparent")
        progress_start_frame.pack(fill="x", side="bottom", padx=0, pady=PAD_Y)
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

    # --- Methods to add/remove Label:ID rows ---
    def add_event_map_entry(self, event=None):
        """Adds a new row with Condition Label and Numerical ID entry fields."""
        # Validator for integer IDs
        validate_int_cmd = (self.register(self._validate_integer_input), '%P')

        # Container frame for one row
        entry_frame = ctk.CTkFrame(self.event_map_scroll_frame, fg_color="transparent")
        entry_frame.pack(fill="x", pady=1, padx=1)

        # Condition Label entry
        label_entry = ctk.CTkEntry(
            entry_frame,
            placeholder_text="Condition Label",
            width=LABEL_ID_ENTRY_WIDTH*2,
            corner_radius=CORNER_RADIUS
        )
        label_entry.pack(side="left", fill="x", expand=True, padx=(0, PAD_X))
        # Bind Enter on the internal tk.Entry
        label_entry._entry.bind("<Return>",   self._add_row_and_focus_label)
        label_entry._entry.bind("<KP_Enter>", self._add_row_and_focus_label)

        # Numerical ID entry
        id_entry = ctk.CTkEntry(
            entry_frame,
            placeholder_text="Numerical ID",
            width=LABEL_ID_ENTRY_WIDTH,
            validate='key',
            validatecommand=validate_int_cmd,
            corner_radius=CORNER_RADIUS
        )
        id_entry.pack(side="left", padx=(0, PAD_X))
        # Bind Enter on the internal tk.Entry
        id_entry._entry.bind("<Return>",   self._add_row_and_focus_label)
        id_entry._entry.bind("<KP_Enter>", self._add_row_and_focus_label)

        # Remove‑row button
        remove_btn = ctk.CTkButton(
            entry_frame,
            text="X",
            width=28,
            height=28,
            corner_radius=CORNER_RADIUS,
            command=lambda ef=entry_frame: self.remove_event_map_entry(ef)
        )
        remove_btn.pack(side="right")

        # Track this row’s widgets
        self.event_map_entries.append({
            'frame': entry_frame,
            'label': label_entry,
            'id':    id_entry,
            'button': remove_btn
        })

        # If this was user‑initiated (not the initial call), focus the new label
        if event is None and label_entry.winfo_exists():
            label_entry.focus_set()


    def remove_event_map_entry(self, entry_frame_to_remove):
        """Removes the specified Label:ID row."""
        try:
            # Find the corresponding entry in the list
            entry_to_remove = None
            for entry_data in self.event_map_entries:
                if entry_data['frame'] == entry_frame_to_remove:
                    entry_to_remove = entry_data
                    break

            if entry_to_remove:
                 # Check if widgets exist before destroying
                 if entry_to_remove['frame'].winfo_exists():
                     entry_to_remove['frame'].destroy()
                 self.event_map_entries.remove(entry_to_remove)

                 # Ensure there's always at least one entry row
                 if not self.event_map_entries:
                     self.add_event_map_entry()
                 # Optional: set focus to the last label entry if available
                 elif self.event_map_entries and self.event_map_entries[-1]['label'].winfo_exists():
                      self.event_map_entries[-1]['label'].focus_set()
            else:
                 self.log("Warning: Could not find the specified event map row to remove.")

        except Exception as e:
            self.log(f"Error removing Event Map row: {e}")
    # --- End Label:ID row methods ---


    def select_save_folder(self):
        # (Unchanged)
        folder = filedialog.askdirectory(title="Select Parent Folder for Excel Output")
        if folder:
            self.save_folder_path.set(folder)
            self.log(f"Output folder: {folder}")
        else:
            self.log("Save folder selection cancelled.")

    def update_select_button_text(self):
        # (Unchanged)
        mode = self.file_mode.get()
        text = "Select EEG File..." if mode == "Single" else "Select Folder..."
        if hasattr(self, 'select_button') and self.select_button:
             self.select_button_text.set(text)

    def select_data_source(self):
        # (Unchanged - still selects BDF/SET files)
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
            else: # Batch mode
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


    # --- Logging (Unchanged) ---
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


    # --- Event ID Detection (Changed for find_events) ---
    def detect_and_show_event_ids(self):
        # --- Method Renamed and Logic Changed ---
        self.log("Detect Numerical IDs button clicked...")
        if self.detection_thread and self.detection_thread.is_alive():
            messagebox.showwarning("Busy", "Event detection is already running.")
            return

        if not self.data_paths:
            messagebox.showerror("No Data Selected", "Please select a data file or folder first.")
            self.log("Detection failed: No data selected.")
            return

        # Get stim channel name from GUI (use default if empty)
        stim_channel_name = self.stim_channel_entry.get().strip()
        if not stim_channel_name:
             stim_channel_name = DEFAULT_STIM_CHANNEL
             self.log(f"Stim channel entry empty, using default: {DEFAULT_STIM_CHANNEL}")
        else:
             self.log(f"Using stim channel from entry: {stim_channel_name}")


        representative_file = self.data_paths[0]
        self.log(f"Starting background detection using mne.find_events for: {os.path.basename(representative_file)}")
        self._disable_controls(enable_process_buttons=None)

        try:
            self.detection_thread = threading.Thread(
                target=self._detection_thread_func,
                # --- Pass stim channel name ---
                args=(representative_file, stim_channel_name, self.gui_queue),
                daemon=True
            )
            self.detection_thread.start()
            self.after(100, self._periodic_detection_queue_check)
        except Exception as start_err:
             self.log(f"Error starting detection thread: {start_err}")
             messagebox.showerror("Thread Error", f"Could not start detection thread:\n{start_err}")
             self._enable_controls(enable_process_buttons=None)


    def _detection_thread_func(self, file_path, stim_channel_name, gui_queue):
        # --- Method Changed to use find_events ---
        """ Background task: load file and find numerical event IDs. """
        raw = None
        gc.collect()
        try:
            # Load the file (don't necessarily need stim_channel here, find_events takes it)
            raw = self.load_eeg_file(file_path)
            if raw is None:
                raise ValueError("File loading failed (check log).")

            gui_queue.put({'type': 'log', 'message': f"Searching for numerical triggers on channel '{stim_channel_name}'..."})
            # --- Use mne.find_events ---
            try:
                 # consecutive=True helps merge stepped triggers if needed
                 events = mne.find_events(raw, stim_channel=stim_channel_name, consecutive=True, verbose=False)
            except ValueError as find_err:
                 if "not found" in str(find_err):
                      gui_queue.put({'type': 'log', f'message': f"Error: Stim channel '{stim_channel_name}' not found in {os.path.basename(file_path)}."})
                      gui_queue.put({'type': 'detection_error', 'message': f"Stim channel '{stim_channel_name}' not found."})
                      return # Exit thread
                 else:
                      raise find_err # Re-raise other find_events errors

            if events is None or len(events) == 0:
                gui_queue.put({'type': 'log', 'message': f"Info: No events found using mne.find_events on channel '{stim_channel_name}'."})
                detected_ids = []
            else:
                # Extract the unique numerical IDs from the third column of the events array
                unique_numeric_ids = sorted(list(np.unique(events[:, 2])))
                gui_queue.put({'type': 'log', 'message': f"Found {len(events)} event triggers. Unique numerical IDs: {unique_numeric_ids}"})
                detected_ids = unique_numeric_ids

            # Send numerical IDs back
            gui_queue.put({'type': 'detection_result', 'ids': detected_ids})

        except Exception as e:
            error_msg = f"Error during event ID detection: {e}"
            gui_queue.put({'type': 'log', 'message': f"!!! {error_msg}\n{traceback.format_exc()}"})
            gui_queue.put({'type': 'detection_error', 'message': error_msg})
        finally:
            if raw:
                del raw
                gc.collect()
            gui_queue.put({'type': 'detection_done'})


    def _periodic_detection_queue_check(self):
        # --- Method Changed: Only displays detected IDs ---
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
                    if detected_ids:
                        # --- Just show the IDs found ---
                        id_string = ", ".join(map(str, detected_ids))
                        self.log(f"Detected numerical event IDs: {id_string}")
                        messagebox.showinfo("Numerical IDs Detected",
                                            f"Found the following unique numerical event IDs in the file:\n\n{id_string}\n\nPlease enter the desired Label:ID pairs manually below.")
                    else:
                        messagebox.showinfo("No Numerical IDs Found", "No numerical event triggers found using mne.find_events.\nPlease check the Stimulus Channel Name or the file content.")
                    detection_finished = True
                elif msg_type == 'detection_error':
                    error_msg = message.get('message', 'Unknown error.')
                    messagebox.showerror("Detection Error", error_msg)
                    detection_finished = True
                elif msg_type == 'detection_done':
                    self.log("Numerical ID detection thread finished.")
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
            self.log("Numerical ID detection process finished.")
        elif self.detection_thread and self.detection_thread.is_alive():
            self.after(100, self._periodic_detection_queue_check)
        else:
             self.log("Warn: Detection thread ended unexpectedly.")
             self._enable_controls(enable_process_buttons=None)
             self.detection_thread = None

    # --- Removed _clear_and_reset_event_id_fields ---
    # Not needed as we don't auto-populate the Label:ID map anymore.


    # --- Core Processing Control ---
    def start_processing(self):
        # (Checks remain the same)
        if self.processing_thread and self.processing_thread.is_alive():
            messagebox.showwarning("Busy", "Processing is already running.")
            return
        if self.detection_thread and self.detection_thread.is_alive():
            messagebox.showwarning("Busy", "Event detection is running. Please wait.")
            return

        self.log("=" * 50)
        self.log("START PROCESSING Initiated...")

        # --- Validation updated for event map ---
        if not self._validate_inputs():
            return

        self.preprocessed_data = {}
        self.progress_bar.set(0)
        self._max_progress = len(self.data_paths)

        self._disable_controls(enable_process_buttons=False)

        self.log("Starting background processing thread...")
        # --- Pass validated params including event_id_map ---
        thread_args = (list(self.data_paths), self.validated_params.copy(), self.gui_queue)
        self.processing_thread = threading.Thread(
            target=self._processing_thread_func,
            args=thread_args,
            daemon=True
        )
        self.processing_thread.start()
        self.after(100, self._periodic_queue_check)


    # --- Input Validation (Updated for Event Map & Stim Channel) ---
    def _validate_inputs(self):
        """ Validates file selection, folder, parameters, and event map. """
        # 1. Check Data Paths (Unchanged)
        if not self.data_paths:
            self.log("V-Error: No data.")
            messagebox.showerror("Input Error", "No data selected.")
            return False

        # 2. Check Save Folder (Unchanged)
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

        # 3. Validate Parameters
        params = {}
        try:
            def get_float(e): return float(e.get().strip()) if e.get().strip() else None
            def get_int(e): return int(e.get().strip()) if e.get().strip() else None

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
            params['ref_channel1'] = self.ref_channel1_entry.get().strip()
            params['ref_channel2'] = self.ref_channel2_entry.get().strip()
            params['max_idx_keep'] = get_int(self.max_idx_keep_entry)
            assert params['max_idx_keep'] is None or params['max_idx_keep'] > 0, "Max Keep Idx > 0"
            if (params['low_pass'] is not None) and (params['high_pass'] is not None):
                 assert params['low_pass'] < params['high_pass'], "Low Pass < High Pass"

            # --- Get Stim Channel Name ---
            stim_ch = self.stim_channel_entry.get().strip()
            if not stim_ch:
                self.log("Stim Channel Name empty, using default: " + DEFAULT_STIM_CHANNEL)
                params['stim_channel'] = DEFAULT_STIM_CHANNEL
            else:
                params['stim_channel'] = stim_ch

            params['save_preprocessed'] = self.save_preprocessed.get()

        except AssertionError as e:
             self.log(f"V-Error: Invalid parameter: {e}")
             messagebox.showerror("Parameter Error", f"Invalid parameter value:\n{e}")
             return False
        except ValueError as e:
             self.log(f"V-Error: Non-numeric parameter: {e}")
             messagebox.showerror("Parameter Error", f"Invalid numeric value entered.\nPlease check parameters.")
             return False
        except Exception as e:
             self.log(f"V-Error: Unexpected param validation: {e}")
             messagebox.showerror("Parameter Error", f"Unexpected error validating parameters:\n{e}")
             return False

        # --- 4. Validate Event Label:ID Map ---
        event_map = {}
        try:
            unique_labels = set()
            unique_ids = set()
            for entry_data in self.event_map_entries:
                 label = entry_data['label'].get().strip()
                 id_str = entry_data['id'].get().strip()

                 # Skip empty rows silently
                 if not label and not id_str:
                      continue

                 # Validate row completeness
                 if not label:
                      messagebox.showerror("Event Map Error", "Found a row with a Numerical ID but no Condition Label.")
                      entry_data['label'].focus_set()
                      return False
                 if not id_str:
                      messagebox.showerror("Event Map Error", f"Condition Label '{label}' has no Numerical ID specified.")
                      entry_data['id'].focus_set()
                      return False

                 # Validate Label uniqueness
                 if label in unique_labels:
                      messagebox.showerror("Event Map Error", f"Duplicate Condition Label found: '{label}'. Labels must be unique.")
                      entry_data['label'].focus_set()
                      return False
                 unique_labels.add(label)

                 # Validate ID is integer and unique (optional check for unique IDs, depends on paradigm)
                 try:
                      num_id = int(id_str)
                      # Example: check if ID is already used - uncomment if IDs must be unique
                      # if num_id in unique_ids:
                      #     messagebox.showerror("Event Map Error", f"Duplicate Numerical ID found: {num_id}. IDs might need to be unique depending on analysis.")
                      #     entry_data['id'].focus_set()
                      #     return False
                      unique_ids.add(num_id)
                 except ValueError:
                      # This should be caught by validate_int_cmd, but double-check
                      messagebox.showerror("Event Map Error", f"Invalid Numerical ID for label '{label}': '{id_str}'. Must be an integer.")
                      entry_data['id'].focus_set()
                      return False

                 event_map[label] = num_id

            if not event_map:
                self.log("V-Error: No Event Map entries.")
                messagebox.showerror("Event Map Error", "Please enter at least one Condition Label and its corresponding Numerical ID.")
                # Focus the first empty field if available
                if self.event_map_entries:
                     if not self.event_map_entries[0]['label'].get().strip():
                          self.event_map_entries[0]['label'].focus_set()
                     elif not self.event_map_entries[0]['id'].get().strip():
                          self.event_map_entries[0]['id'].focus_set()
                return False

            # Store the validated map
            params['event_id_map'] = event_map
            # Store validated parameters including the map
            self.validated_params = params

        except Exception as e:
            self.log(f"V-Error: Unexpected error validating Event Map: {e}")
            messagebox.showerror("Event Map Error", f"An unexpected error occurred during Event Map validation:\n{e}")
            return False

        # --- End Event Map Validation ---

        self.log("Inputs Validated Successfully.")
        self.log(f"Parameters: {self.validated_params}") # Includes event_id_map now
        # self.log(f"Event Map: {self.validated_params['event_id_map']}") # Logged within params
        return True


    # --- Periodic Queue Check (Main Processing - Unchanged) ---
    def _periodic_queue_check(self):
        processing_done = False
        final_success = True
        try:
            while True:
                message = self.gui_queue.get_nowait()
                msg_type = message.get('type')

                # ignore detection‐thread messages
                if msg_type in ['detection_result', 'detection_error', 'detection_done']:
                    continue

                if msg_type == 'log':
                    self.log(message.get('message', ''))
                elif msg_type == 'progress':
                    value = message.get('value', 0)
                    progress_fraction = (value / self._max_progress) if self._max_progress > 0 else 0
                    self.progress_bar.set(progress_fraction)
                    # ── FORCE IMMEDIATE REDRAW ──
                    self.update_idletasks()
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
                    if tb: print(tb)
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

    # --- Finalize Processing (Unchanged) ---
    def _finalize_processing(self, success):
        self.progress_bar.set(1.0 if success else self.progress_bar.get())
        if success and self.preprocessed_data:
            has_data = any(bool(epochs_list) for epochs_list in self.preprocessed_data.values())
            if has_data:
                self.log("\n--- Starting Post-processing Phase ---")
                try:
                    # Pass the condition labels (keys of the dict) to post_process
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


    # --- Disable/Enable Controls (Updated for new widgets) ---
    def _disable_controls(self, enable_process_buttons=None):
        """Disables GUI elements during processing."""
        widgets = []
        # Options Frame
        if hasattr(self, 'select_button'): widgets.append(self.select_button)
        if hasattr(self, 'radio_single'): widgets.append(self.radio_single)
        if hasattr(self, 'radio_batch'): widgets.append(self.radio_batch)
        if hasattr(self, 'options_frame'):
            for w in self.options_frame.winfo_children():
                 if isinstance(w, ctk.CTkRadioButton) and w not in [self.radio_single, self.radio_batch]:
                     widgets.append(w)

        # Params Frame (incl. stim channel)
        param_widgets = [getattr(self, n, None) for n in [
            'low_pass_entry', 'high_pass_entry', 'downsample_entry',
            'epoch_start_entry', 'epoch_end_entry', 'reject_thresh_entry',
            'ref_channel1_entry', 'ref_channel2_entry', 'max_idx_keep_entry',
            'stim_channel_entry', # Added
            'save_preprocessed_checkbox'
        ]]
        widgets.extend([w for w in param_widgets if w])

        # Event Map Frame
        if hasattr(self, 'detect_button'): widgets.append(self.detect_button)
        if hasattr(self, 'add_map_button'): widgets.append(self.add_map_button)
        for entry_data in self.event_map_entries:
             if entry_data['label'] and entry_data['label'].winfo_exists(): widgets.append(entry_data['label'])
             if entry_data['id'] and entry_data['id'].winfo_exists(): widgets.append(entry_data['id'])
             if entry_data['button'] and entry_data['button'].winfo_exists(): widgets.append(entry_data['button'])

        # Save Frame
        if hasattr(self, 'save_frame'):
            for w in self.save_frame.winfo_children():
                if isinstance(w, ctk.CTkButton): widgets.append(w)

        # Start Button
        if hasattr(self, 'start_button') and enable_process_buttons is False:
            widgets.append(self.start_button)

        # Disable all
        for w in widgets:
             if w and w.winfo_exists():
                 try: w.configure(state="disabled")
                 except Exception: pass
        self.update_idletasks()


    def _enable_controls(self, enable_process_buttons=None):
        """Enables GUI elements after processing."""
        widgets = []
        # Options Frame
        if hasattr(self, 'select_button'): widgets.append(self.select_button)
        if hasattr(self, 'radio_single'): widgets.append(self.radio_single)
        if hasattr(self, 'radio_batch'): widgets.append(self.radio_batch)
        if hasattr(self, 'options_frame'):
             for w in self.options_frame.winfo_children():
                  if isinstance(w, ctk.CTkRadioButton) and w not in [self.radio_single, self.radio_batch]:
                      widgets.append(w)

        # Params Frame (incl. stim channel)
        param_widgets = [getattr(self, n, None) for n in [
            'low_pass_entry', 'high_pass_entry', 'downsample_entry',
            'epoch_start_entry', 'epoch_end_entry', 'reject_thresh_entry',
            'ref_channel1_entry', 'ref_channel2_entry', 'max_idx_keep_entry',
            'stim_channel_entry', # Added
            'save_preprocessed_checkbox'
        ]]
        widgets.extend([w for w in param_widgets if w])

        # Event Map Frame
        if hasattr(self, 'detect_button'): widgets.append(self.detect_button)
        if hasattr(self, 'add_map_button'): widgets.append(self.add_map_button)
        for entry_data in self.event_map_entries:
              if entry_data['label'] and entry_data['label'].winfo_exists(): widgets.append(entry_data['label'])
              if entry_data['id'] and entry_data['id'].winfo_exists(): widgets.append(entry_data['id'])
              if entry_data['button'] and entry_data['button'].winfo_exists(): widgets.append(entry_data['button'])

        # Save Frame
        if hasattr(self, 'save_frame'):
            for w in self.save_frame.winfo_children():
                 if isinstance(w, ctk.CTkButton): widgets.append(w)

        # Start Button
        if hasattr(self, 'start_button') and enable_process_buttons is True:
            widgets.append(self.start_button)

        # Enable all
        for w in widgets:
             if w and w.winfo_exists():
                 try: w.configure(state="normal")
                 except Exception: pass
        self.update_idletasks()


    # --- Background Processing Thread Function (Updated for find_events and map) ---
    def _processing_thread_func(self, data_paths, params, gui_queue):
        # --- No conditions_ids_to_process argument needed, map is in params ---
        event_id_map = params['event_id_map'] # e.g., {'Cond1': 32, 'Cond2': 64}
        condition_labels = list(event_id_map.keys()) # Get labels ['Cond1', 'Cond2']
        stim_channel_name = params['stim_channel'] # Get stim channel from params

        # Initialize local data storage using labels as keys
        local_data = {label: [] for label in condition_labels}
        files_w_epochs = 0
        gc.collect()

        try:
            n_files = len(data_paths)
            for i, f_path in enumerate(data_paths):
                f_name = os.path.basename(f_path)
                gui_queue.put({'type': 'log', 'message': f"\nProcessing file {i+1}/{n_files}: {f_name}"})
                raw, raw_proc = None, None # Removed evts variable
                gc.collect()

                try:
                    # 1. Load Raw Data (Unchanged)
                    raw = self.load_eeg_file(f_path)
                    if raw is None: continue

                    # 2. Preprocess Raw Data (Unchanged, but uses stim_channel param implicitly if needed by funcs)
                    # Pass the raw copy and other parameters from the validated dict
                    raw_proc = self.preprocess_raw(raw.copy(), **params)
                    if raw_proc is None:
                        del raw; gc.collect()
                        continue
                    del raw; gc.collect()

                    # 3. Find Numerical Events
                    gui_queue.put({'type': 'log', 'message': f"Searching for numerical triggers on channel '{stim_channel_name}'..."})
                    try:
                         # Use find_events on the *preprocessed* data
                         events = mne.find_events(raw_proc, stim_channel=stim_channel_name, consecutive=True, verbose=False)
                         gui_queue.put({'type': 'log', 'message': f"Found {len(events)} event triggers."})
                    except ValueError as find_err:
                         # Handle case where stim channel isn't found *after* preprocessing (e.g., if dropped)
                         gui_queue.put({'type': 'log', 'message': f"Warning: Could not find stim channel '{stim_channel_name}' after preprocessing in {f_name}. Skipping epoching for this file. Error: {find_err}"})
                         events = None # Ensure events is None if find_events fails
                    except Exception as find_err:
                         gui_queue.put({'type': 'log', 'message': f"Warning: Error during mne.find_events for {f_name}. Skipping epoching. Error: {find_err}"})
                         events = None

                    # 4. Create Epochs for desired conditions using the map
                    file_produced_epochs = False
                    if events is not None and len(events) > 0:
                         # Iterate through the Label:ID map provided by the user
                         for label, number in event_id_map.items():
                             try:
                                 # Create epochs for *only* the current label/number pair
                                 current_event_id = {label: number}
                                 epochs = mne.Epochs(raw_proc, events,
                                                     event_id=current_event_id,
                                                     tmin=params['epoch_start'],
                                                     tmax=params['epoch_end'],
                                                     preload=False, # Load later
                                                     verbose=False,
                                                     baseline=None,
                                                     on_missing='warn') # Warn if ID not in events

                                 if len(epochs.events) > 0:
                                     gui_queue.put({'type': 'log', 'message': f"  -> Created {len(epochs.events)} epochs for '{label}' (ID: {number})."})
                                     # Append to the list associated with the label
                                     local_data[label].append(epochs)
                                     file_produced_epochs = True
                                 # else: MNE automatically handles if the event ID isn't found in 'events' with on_missing='warn'

                             except Exception as ep_err:
                                 gui_queue.put({'type': 'log', 'message': f"!!! Epoch creation error for '{label}' (ID: {number}): {ep_err}\n{traceback.format_exc()}"})
                         if file_produced_epochs:
                             files_w_epochs += 1
                    else:
                        gui_queue.put({'type': 'log', 'message': "Skipping epoch creation (no events found or error finding events)."})

                    # 5. Optional: Save Preprocessed Data (Unchanged)
                    if params['save_preprocessed']:
                         p_path = os.path.join(os.path.dirname(f_path), f"{os.path.splitext(f_name)[0]}_preproc_raw.fif")
                         try:
                             gui_queue.put({'type': 'log', 'message': f"Saving preprocessed to: {p_path}"})
                             raw_proc.save(p_path, overwrite=True, verbose=False)
                         except Exception as s_err:
                             gui_queue.put({'type': 'log', 'message': f"Warn: Save failed: {s_err}"})

                except MemoryError as mem_err:
                     gui_queue.put({'type': 'error',
                                     'message': f"Memory Error {f_name}: {mem_err}",
                                     'traceback': traceback.format_exc()})
                     del raw, raw_proc; gc.collect()
                     return # Stop thread
                except Exception as f_err:
                    gui_queue.put({'type': 'log', 'message': f"!!! FILE ERROR {f_name}: {f_err}\n{traceback.format_exc()}"})
                finally:
                    del raw_proc # Ensure cleanup
                    gc.collect()
                    gui_queue.put({'type': 'progress', 'value': i + 1})

            # --- Loop Finished ---
            gui_queue.put({'type': 'log', 'message': f"\n--- BG Preprocessing Done ---"})
            gui_queue.put({'type': 'log', 'message': f"Found epochs for {files_w_epochs}/{n_files} files matching the specified Label:ID map."})
            gui_queue.put({'type': 'result', 'data': local_data})

        except MemoryError as mem_err:
             gui_queue.put({'type': 'error', 'message': f"Critical Memory Error: {mem_err}", 'traceback': traceback.format_exc()})
        except Exception as e:
            gui_queue.put({'type': 'error', 'message': f"Critical thread error: {e}", 'traceback': traceback.format_exc()})
        finally:
            gui_queue.put({'type': 'done'})


    # --- EEG Loading Method (Mostly unchanged, stim_channel handled later) ---
    def load_eeg_file(self, filepath):
        """Loads BDF or SET file using MNE-Python."""
        ext = os.path.splitext(filepath)[1].lower()
        raw = None
        base_filename = os.path.basename(filepath)
        self.log(f"Loading: {base_filename}...")

        try:
            load_kwargs = {'preload': True, 'verbose': False}

            if ext == ".bdf":
                 try:
                     self.log(f"Attempting BDF load, will find events later using specified stim channel ('{self.validated_params.get('stim_channel', DEFAULT_STIM_CHANNEL)}').")
                     # Load without forcing 'Status', let find_events handle channel later
                     with mne.utils.use_log_level('WARNING'):
                          raw = mne.io.read_raw_bdf(filepath, **load_kwargs)
                     self.log("BDF loaded successfully.")
                 except Exception as bdf_err:
                      raise bdf_err # Let outer handler catch load errors

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
                 raw.set_montage(montage, on_missing='warn', match_case=False, verbose=False) # Match case False is important
                 self.log("Montage applied (check warnings for missing channels like EXG or Stim).")
            except Exception as m_err:
                 self.log(f"Warning: Montage error: {m_err}")

            return raw
        except MemoryError as me:
             self.log(f"!!! Memory Error loading {base_filename}: {me}")
             messagebox.showerror("Memory Error", f"Memory Error loading {base_filename}.")
             return None
        # ... (other error handling unchanged) ...
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


    # --- Preprocessing Method (Updated to preserve stim channel) ---
    def preprocess_raw(self, raw, **params):
        """Applies the preprocessing steps to a raw MNE object."""
        downsample_rate = params.get('downsample_rate')
        low_pass = params.get('low_pass')
        high_pass = params.get('high_pass')
        reject_thresh = params.get('reject_thresh')
        ref_channel1 = params.get('ref_channel1')
        ref_channel2 = params.get('ref_channel2')
        max_idx_keep = params.get('max_idx_keep')
        # --- Get stim channel name ---
        stim_channel_name = params.get('stim_channel', DEFAULT_STIM_CHANNEL)

        try:
            ch_names_orig = list(raw.info['ch_names'])
            n_chans_orig = len(ch_names_orig)
            self.log(f"Preprocessing {n_chans_orig} chans...")

            # 1. Initial Bipolar Reference (Unchanged)
            if ref_channel1 and ref_channel2:
                 # ... (bipolar logic unchanged) ...
                 if ref_channel1 in ch_names_orig and ref_channel2 in ch_names_orig:
                    new_channel_name = f"{ref_channel1}-{ref_channel2}"
                    try:
                        self.log(f"Applying bipolar ref: {ref_channel1}-{ref_channel2}...")
                        raw.set_bipolar_reference(ref_channel1, ref_channel2, new_channel_name, drop_refs=False, copy=False, verbose=False)
                        self.log(f"OK. New channel '{new_channel_name}'.")
                        ch_names_orig = list(raw.info['ch_names']) # Update names
                        n_chans_orig = len(ch_names_orig)
                    except Exception as bipol_err:
                        self.log(f"Warn: Bipolar ref failed: {bipol_err}.")
                 else:
                    self.log(f"Warn: One or both ref channels ({ref_channel1}, {ref_channel2}) not found. Skipping bipolar ref.")
            else:
                self.log("Skip bipolar ref.")


            # 2. Drop channels (preserve stim channel)
            # --- Modified to preserve stim channel ---
            if max_idx_keep is not None:
                c_names = list(raw.info['ch_names'])
                c_n = len(c_names)
                # Find stim channel(s) matching the name (case sensitive)
                stim_chans_found = [ch for ch in c_names if ch == stim_channel_name]

                if 0 < max_idx_keep < c_n:
                    # Keep first max_idx_keep channels plus any matching stim channels
                    to_keep = list(dict.fromkeys(
                        [c_names[i] for i in range(max_idx_keep)] + stim_chans_found
                    ))
                    to_drop = [ch for ch in c_names if ch not in to_keep]
                    if to_drop:
                        self.log(f"Dropping {len(to_drop)} chans (keeping first {max_idx_keep} and stim channel '{stim_channel_name}')...")
                        try:
                            raw.drop_channels(to_drop)
                            self.log(f"OK. Remaining: {len(raw.ch_names)}")
                        except Exception as drop_err:
                            self.log(f"Warn: Drop failed: {drop_err}")
                    else:
                         self.log("No channels to drop based on Max Idx Keep & Stim channel criteria.")
                elif max_idx_keep >= c_n:
                     self.log(f"Info: Max Idx Keep ({max_idx_keep}) >= chans ({c_n}). No drop.")
            else:
                 self.log("Skip index drop.")

            # 3. Downsampling (Unchanged)
            if downsample_rate:
                 # ... (downsample logic unchanged) ...
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


            # 4. Filtering (Unchanged)
            l = low_pass if low_pass and low_pass > 0 else None
            h = high_pass
            if l or h:
                # ... (filter logic unchanged) ...
                 try:
                    self.log(f"Filtering ({l if l else 'DC'}-{h if h else 'Nyq'}Hz)...")
                    raw.filter(l, h, method='fir', phase='zero-double', fir_window='hamming',
                               fir_design='firwin', pad='edge', verbose=False)
                    self.log("Filter OK.")
                 except Exception as f_err:
                    self.log(f"Warn: Filter failed: {f_err}.")
            else:
                 self.log("Skip filter.")


            # 5. Kurtosis-based rejection & interpolation (Unchanged)
            if reject_thresh:
                 # ... (kurtosis logic unchanged, uses pick_types(eeg=True) so ignores stim channel) ...
                 self.log(f"Kurtosis rejection (Z > {reject_thresh})...")
                 orig_bads = list(raw.info['bads'])
                 try:
                    picks = mne.pick_types(raw.info, eeg=True, exclude='bads')
                    if len(picks) >= 2:
                        d = raw.get_data(picks)
                        k = kurtosis(d, axis=1, fisher=True, bias=False)
                        del d; gc.collect()
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
                            if new_b:
                                raw.info['bads'].extend(new_b)
                                self.log(f"Bad by Kurt: {new_b}. Total bads: {raw.info['bads']}")
                            else:
                                self.log("No new bads by Kurtosis.")
                        else:
                            self.log("No channels exceeded kurtosis threshold.")

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
                     raw.info['bads'] = orig_bads # Restore bads on error
            else:
                 self.log("Skip Kurtosis.")


            # 6. Apply average reference (Unchanged)
            try:
                 # ... (avg ref logic unchanged, uses set_eeg_reference so ignores stim channel) ...
                 self.log("Applying avg ref...")
                 raw.set_eeg_reference('average', projection=True, verbose=False)
                 raw.apply_proj(verbose=False)
                 self.log("Avg ref OK.")
            except Exception as avg_err:
                 self.log(f"Warn: Avg ref failed: {avg_err}")

            self.log(f"Preproc OK. State: {len(raw.ch_names)} chans, {raw.info['sfreq']:.1f} Hz.")
            return raw
        except MemoryError as me:
            self.log(f"!!! Memory Error preprocessing: {me}")
            return None
        except Exception as e:
            self.log(f"!!! CRITICAL preproc error: {e}")
            print(traceback.format_exc())
            return None

    def _focus_next_id_entry(self, event):
        """
        When Enter/Return is pressed in a Numerical ID CTkEntry,
        move focus to the next ID field (and create it if it doesn’t exist).
        """
        widget = event.widget
        target_idx = None

        # find which row this widget belongs to
        for idx, entry in enumerate(self.event_map_entries):
            ctkw = entry['id']  # your CTkEntry
            internal = getattr(ctkw, "entry", None) or getattr(ctkw, "_entry", None)
            # match either the CTkEntry or its internal tk.Entry
            if ctkw is widget or internal is widget:
                target_idx = idx
                break

    def _add_row_and_focus_label(self, event):
        """
        When Enter/Return is pressed in either the Label or ID entry,
        create a new row and focus its Label field.
        """
        # Add the new row
        self.add_event_map_entry()
        # Focus the label entry of the newly created row
        self.event_map_entries[-1]['label'].focus_set()
        # Prevent default behavior
        return "break"

    # --- Post-processing Method (Unchanged - uses labels from dict keys) ---


def post_process(self, condition_labels_present):
    """Calculates metrics (FFT, SNR, Z-score, BCA) using time-domain epoch averaging and saves to Excel."""
    self.log("--- Post-processing: Calculating Metrics & Saving Excel ---")
    parent_folder = self.save_folder_path.get()
    if not parent_folder or not os.path.isdir(parent_folder):
        self.log(f"Error: Invalid save folder: '{parent_folder}'")
        messagebox.showerror("Save Error", f"Invalid output folder:\n{parent_folder}")
        return

    # Extract participant ID from first filename (e.g. "SC_P14.bdf" → "P14")
    first_file = os.path.basename(self.data_paths[0])
    pid = os.path.splitext(first_file)[0]
    if "_" in pid:
        pid = pid.split("_")[-1]

    any_results_saved = False

    for cond_label in condition_labels_present:
        epochs_list = self.preprocessed_data.get(cond_label, [])
        if not epochs_list:
            self.log(f"\nSkipping post-processing for '{cond_label}': No epoch data.")
            continue

        self.log(f"\nPost-processing '{cond_label}' ({len(epochs_list)} file(s))...")
        accum = {'fft': None, 'snr': None, 'z': None, 'bca': None}
        valid_count = 0
        electrode_names = None

        for file_idx, epochs in enumerate(epochs_list):
            self.log(f"  File {file_idx + 1}/{len(epochs_list)} for '{cond_label}'...")
            gc.collect()
            try:
                if not isinstance(epochs, mne.BaseEpochs) or len(epochs.events) == 0:
                    self.log("    Invalid or empty epochs. Skipping.")
                    continue

                epochs.load_data()
                picks = mne.pick_types(epochs.info, eeg=True, exclude='bads')
                if picks.size == 0:
                    self.log("    No good EEG channels. Skipping.")
                    continue

                # Time-domain averaging over epochs
                ep_data = epochs.get_data(picks=picks)  # [n_epochs, n_ch, n_t]
                avg_data = np.mean(ep_data.astype(np.float64), axis=0)  # [n_ch, n_t]
                n_ch, n_t = avg_data.shape
                sfreq = epochs.info['sfreq']

                # Build MATLAB-style freq vector
                num_bins = n_t // 2 + 1
                freqs = np.linspace(0, sfreq / 2.0, num=num_bins, endpoint=True)

                # Single FFT on averaged waveform
                fft_full = np.fft.fft(avg_data, axis=1)
                fft_vals = np.abs(fft_full[:, :num_bins]) / n_t * 2  # [n_ch, num_bins]

                # Determine electrode names once
                if electrode_names is None:
                    ch_names = [epochs.info['ch_names'][i] for i in picks]
                    if n_ch == len(DEFAULT_ELECTRODE_NAMES_64) and ch_names == DEFAULT_ELECTRODE_NAMES_64:
                        electrode_names = DEFAULT_ELECTRODE_NAMES_64
                    elif n_ch == len(DEFAULT_ELECTRODE_NAMES_64):
                        electrode_names = DEFAULT_ELECTRODE_NAMES_64
                    else:
                        electrode_names = [f"Ch{i + 1}" for i in range(n_ch)]

                # Allocate arrays for metrics
                n_tf = len(TARGET_FREQUENCIES)
                f_fft = np.zeros((n_ch, n_tf))
                f_snr = np.zeros((n_ch, n_tf))
                f_z = np.zeros((n_ch, n_tf))
                f_bca = np.zeros((n_ch, n_tf))

                # Compute metrics
                for c_idx in range(n_ch):
                    for f_idx, t_freq in enumerate(TARGET_FREQUENCIES):
                        if not (freqs[0] <= t_freq <= freqs[-1]):
                            continue
                        t_bin = np.argmin(np.abs(freqs - t_freq))

                        # MATLAB-style noise window: 25 bins, exclude offsets -2,-1,0
                        low, high = t_bin - 12, t_bin + 13
                        exclude = {t_bin - 2, t_bin - 1, t_bin}
                        noise_idx = [
                            i for i in range(low, high)
                            if 0 <= i < len(freqs) and i not in exclude
                        ]

                        if len(noise_idx) >= 4:
                            noise_mean = fft_vals[c_idx, noise_idx].mean()
                            noise_std = fft_vals[c_idx, noise_idx].std()
                        else:
                            noise_mean = noise_std = 0

                        amp_val = fft_vals[c_idx, t_bin]
                        snr_val = amp_val / noise_mean if noise_mean > 1e-12 else 0
                        peak = fft_vals[c_idx, max(0, t_bin - 1):min(len(freqs), t_bin + 2)].max()
                        z_val = (peak - noise_mean) / noise_std if noise_std > 1e-12 else 0
                        bca_val = amp_val - noise_mean

                        f_fft[c_idx, f_idx] = amp_val
                        f_snr[c_idx, f_idx] = snr_val
                        f_z[c_idx, f_idx] = z_val
                        f_bca[c_idx, f_idx] = bca_val

                # Accumulate
                if accum['fft'] is None:
                    accum = {'fft': f_fft, 'snr': f_snr, 'z': f_z, 'bca': f_bca}
                else:
                    accum['fft'] += f_fft
                    accum['snr'] += f_snr
                    accum['z'] += f_z
                    accum['bca'] += f_bca
                valid_count += 1

            except Exception as e:
                self.log(f"!!! Error in post-processing file: {e}")
            finally:
                gc.collect()

        # Average and save if data present
        if valid_count > 0 and electrode_names:
            avg = {k: v / valid_count for k, v in accum.items()}
            cols = [f"{f:.1f}_Hz" for f in TARGET_FREQUENCIES]

            dfs = {
                'FFT Amplitude': pd.DataFrame(avg['fft'], index=electrode_names, columns=cols),
                'SNR': pd.DataFrame(avg['snr'], index=electrode_names, columns=cols),
                'Z Score': pd.DataFrame(avg['z'], index=electrode_names, columns=cols),
                'BCA': pd.DataFrame(avg['bca'], index=electrode_names, columns=cols)
            }
            for df in dfs.values():
                df.insert(0, 'Electrode', df.index)

            # Folder named as condition label
            folder_name = cond_label.replace('/', '-').strip()
            sub_path = os.path.join(parent_folder, folder_name)
            os.makedirs(sub_path, exist_ok=True)

            # File named "PID condition_label.xlsx"
            safe_label = folder_name
            excel_filename = f"{pid} {safe_label}.xlsx"
            excel_path = os.path.join(sub_path, excel_filename)

            self.log(f"Writing Excel: {excel_path}")
            with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                workbook = writer.book
                center_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
                for sheet_name, df in dfs.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    worksheet = writer.sheets[sheet_name]
                    for col_idx, col_name in enumerate(df.columns):
                        header_len = len(str(col_name))
                        try:
                            mx = df[col_name].astype(str).map(len).max()
                            if pd.isna(mx):
                                mx = header_len
                        except:
                            mx = header_len
                        width = max(header_len, int(mx)) + 4
                        worksheet.set_column(col_idx, col_idx, width, center_fmt)

            self.log(f"Saved Excel for '{cond_label}'.")
            any_results_saved = True
        else:
            self.log(f"No valid data for '{cond_label}'. No Excel generated.")

    # Final message
    if any_results_saved:
        self.log("Post-processing complete. Results saved.")
        messagebox.showinfo("Processing Complete", "Analysis finished and Excel files saved successfully.")
    else:
        self.log("Post-processing complete. No results generated.")


# --- Main execution block (Unchanged) ---
if __name__ == "__main__":
    try:
        # Improve DPI awareness on Windows if possible
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass  # Not on Windows or DPI awareness setting fails
    app = FPVSApp()
    app.mainloop()