#!/usr/bin/env python3
"""
EEG FPVS Analysis GUI using MNE-Python and Tkinter.

This script provides a graphical user interface (GUI) to load, preprocess,
and analyze Frequency-Following Potential (FPVS) EEG data.

Version: 1.4 (April 2025) - Numerical Event ID Processing

Key functionalities:
- Load EEG data in .BDF or .set format.
- Process single files or batch process folders.
- Button to automatically detect numerical Event IDs from the first selected file
  using mne.find_events on 'Status' channel.
- Apply preprocessing steps (configurable). Runs in background thread.
- Extract epochs based on user-selected numerical Event IDs.
  Uses preload=False for memory efficiency.
- Perform post-processing using FFT (FFT Amp, SNR, Z-score, BCA).
- Save results to formatted Excel files, named using Event IDs.
- Provides responsive logging and progress feedback via threading & queue.
- Includes tooltips, menu bar, and Enter key binding for conditions fields.
"""

# === Dependencies ===
# Standard Libraries:
import os
import glob
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sys
import traceback
import threading
import queue

# Third-Party Libraries:
import numpy as np
import pandas as pd
from scipy.stats import kurtosis
try:
    import mne
except ImportError: messagebox.showerror("Dependency Error", "MNE-Python required. pip install mne"); sys.exit(1)
try:
    import xlsxwriter
except ImportError: messagebox.showerror("Dependency Error", "XlsxWriter required. pip install xlsxwriter"); sys.exit(1)
# === End Dependencies ===


# =====================================================
# Fixed parameters for post-processing
# =====================================================
TARGET_FREQUENCIES = np.arange(1.2, 16.8 + 1.2, 1.2)
ELECTRODE_NAMES = [
    'Fp1','AF7','AF3','F1','F3','F5','F7','FT7','FC5','FC3',
    'FC1','C1','C3','C5','T7','TP7','CP5','CP3','CP1','P1',
    'P3','P5','P7','P9','PO7','PO3','O1','Iz','Oz','POz','Pz',
    'CPz','Fpz','Fp2','AF8','AF4','AFz','Fz','F2','F4','F6',
    'F8','FT8','FC6','FC4','FC2','FCz','Cz','C2','C4','C6',
    'T8','TP8','CP6','CP4','CP2','P2','P4','P6','P8','P10',
    'PO8','PO4','O2'
]

class FPVSApp(tk.Tk):
    """
    Main application class for the EEG FPVS Analysis GUI.
    Handles GUI, user input, background processing thread, condition/event ID detection,
    and post-processing based on numerical Event IDs.
    """
    def __init__(self):
        """ Initialize the application. """
        super().__init__()
        self.title(f"EEG FPVS Analysis Tool (v1.4 - Event IDs)")
        self.geometry("950x900")

        # Data structures
        self.preprocessed_data = {}    # Keyed by integer Event ID
        self.condition_entries = []    # Holds GUI tk.Entry widgets for Event IDs
        self.current_event_ids_process = [] # List of integer IDs for current run
        self.data_paths = []
        self.tooltip_map = {}
        self.processing_thread = None
        self.gui_queue = queue.Queue()

        self.create_menu()
        self.create_widgets()

    # --- Menu Methods ---
    def create_menu(self):
        """Creates the main menu bar."""
        # (Code Unchanged)
        menubar = tk.Menu(self); self.config(menu=menubar)
        file_menu = tk.Menu(menubar, tearoff=0); menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Select Parent Save Folder...", command=self.select_save_folder)
        file_menu.add_separator(); file_menu.add_command(label="Exit", command=self.quit)
        help_menu = tk.Menu(menubar, tearoff=0); menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About...", command=self.show_about_dialog)

    def show_about_dialog(self):
        """Displays 'About' dialog."""
        # (Code Unchanged)
        messagebox.showinfo(
            "About EEG FPVS Analysis Tool",
            f"Version: 1.4 ({pd.Timestamp.now().strftime('%B %Y')})\n\n"
            "Processes EEG data using MNE-Python.\n"
            "Detects and processes numerical Event IDs.\n"
            "Features background processing."
        )

    def quit(self):
        """Exits the application cleanly."""
        # (Code Unchanged)
        if self.processing_thread and self.processing_thread.is_alive():
            if messagebox.askyesno("Exit Confirmation", "Processing ongoing. Stop and exit?"): self.destroy()
            else: return
        else: self.destroy()

    # --- Validation Methods ---
    # (Code Unchanged)
    def _validate_numeric_input(self, P):
        if P == "" or P == "-": return True
        try: float(P); return True
        except ValueError: self.bell(); return False
    def _validate_integer_input(self, P): # Used for indices AND Event IDs now
        if P == "": return True
        try:
            # Event IDs can theoretically be negative, but usually positive
            # Let's allow any integer for now for flexibility
            int(P); return True
        except ValueError: self.bell(); return False

    # --- Tooltip Methods ---
    # (Code Unchanged)
    def _show_tooltip(self, event):
        widget = event.widget
        try:
            if widget.winfo_exists() and hasattr(self, 'status_bar_label'):
                 tooltip_text = self.tooltip_map.get(widget, "")
                 self.status_bar_label.config(text=tooltip_text)
        except tk.TclError: pass
    def _clear_tooltip(self, event):
        if hasattr(self, 'status_bar_label'): self.status_bar_label.config(text="")
    def _bind_tooltip(self, widget, text):
        self.tooltip_map[widget] = text
        widget.bind("<Enter>", self._show_tooltip, add='+'); widget.bind("<Leave>", self._clear_tooltip, add='+')

    # --- GUI Creation ---
    def create_widgets(self):
        """ Builds and arranges all the GUI components. """
        # Use integer validation for Event ID fields
        validate_event_id_cmd = (self.register(self._validate_integer_input), '%P')
        # Keep separate numeric validation for other fields
        validate_num_cmd = (self.register(self._validate_numeric_input), '%P')
        validate_int_idx_cmd = (self.register(self._validate_integer_input), '%P') # For indices (non-negative)

        main_frame = ttk.Frame(self, padding="10"); main_frame.pack(fill="both", expand=True)
        self.status_bar_label = ttk.Label(self, text="", relief="sunken", anchor="w"); self.status_bar_label.pack(side="bottom", fill="x")

        # Top Options Frame (Unchanged)
        options_frame = ttk.LabelFrame(main_frame, text="Processing Options"); options_frame.pack(fill="x", padx=5, pady=5)
        self.file_mode=tk.StringVar(value="Single"); ttk.Label(options_frame, text="Mode:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.radio_single=ttk.Radiobutton(options_frame, text="Single File", variable=self.file_mode, value="Single", command=self.update_select_button_text); self.radio_single.grid(row=0, column=1, padx=5); self._bind_tooltip(self.radio_single, "Process one selected EEG file.")
        self.radio_batch=ttk.Radiobutton(options_frame, text="Batch Folder", variable=self.file_mode, value="Batch", command=self.update_select_button_text); self.radio_batch.grid(row=0, column=2, padx=5); self._bind_tooltip(self.radio_batch, "Process all matching EEG files in a selected folder.")
        self.file_type=tk.StringVar(value=".BDF"); ttk.Label(options_frame, text="File Type:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        rb_bdf=ttk.Radiobutton(options_frame, text=".BDF", variable=self.file_type, value=".BDF"); rb_bdf.grid(row=1, column=1, padx=5); self._bind_tooltip(rb_bdf, "Select if processing BioSemi Data Format files.")
        rb_set=ttk.Radiobutton(options_frame, text=".set", variable=self.file_type, value=".set"); rb_set.grid(row=1, column=2, padx=5); self._bind_tooltip(rb_set, "Select if processing EEGLAB (.set) files.")

        # Preprocessing Parameters Frame (Unchanged)
        params_frame = ttk.LabelFrame(main_frame, text="Preprocessing Parameters"); params_frame.pack(fill="x", padx=5, pady=5)
        ttk.Label(params_frame, text="Low Pass (Hz):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.low_pass_entry=ttk.Entry(params_frame, width=8, validate='key', validatecommand=validate_num_cmd); self.low_pass_entry.insert(0, "0.1"); self.low_pass_entry.grid(row=0, column=1, padx=5); self._bind_tooltip(self.low_pass_entry, "Low cutoff frequency (Hz).")
        ttk.Label(params_frame, text="High Pass (Hz):").grid(row=0, column=2, sticky="w", padx=5, pady=2)
        self.high_pass_entry=ttk.Entry(params_frame, width=8, validate='key', validatecommand=validate_num_cmd); self.high_pass_entry.insert(0, "50"); self.high_pass_entry.grid(row=0, column=3, padx=5); self._bind_tooltip(self.high_pass_entry, "High cutoff frequency (Hz).")
        ttk.Label(params_frame, text="Downsample (Hz):").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.downsample_entry=ttk.Entry(params_frame, width=8, validate='key', validatecommand=validate_num_cmd); self.downsample_entry.insert(0, "256"); self.downsample_entry.grid(row=1, column=1, padx=5); self._bind_tooltip(self.downsample_entry, "Target sampling rate (Hz).")
        ttk.Label(params_frame, text="Epoch Start (s):").grid(row=1, column=2, sticky="w", padx=5, pady=2)
        self.epoch_start_entry=ttk.Entry(params_frame, width=8, validate='key', validatecommand=validate_num_cmd); self.epoch_start_entry.insert(0, "-1"); self.epoch_start_entry.grid(row=1, column=3, padx=5); self._bind_tooltip(self.epoch_start_entry, "Epoch start time relative to event (s).")
        ttk.Label(params_frame, text="Epoch End (s):").grid(row=1, column=4, sticky="w", padx=5, pady=2)
        self.epoch_end_entry=ttk.Entry(params_frame, width=8, validate='key', validatecommand=validate_num_cmd); self.epoch_end_entry.insert(0, "5"); self.epoch_end_entry.grid(row=1, column=5, padx=5); self._bind_tooltip(self.epoch_end_entry, "Epoch end time relative to event (s).")
        ttk.Label(params_frame, text="Rejection Z-Thresh:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.reject_thresh_entry=ttk.Entry(params_frame, width=8, validate='key', validatecommand=validate_num_cmd); self.reject_thresh_entry.insert(0, "5"); self.reject_thresh_entry.grid(row=2, column=1, padx=5); self._bind_tooltip(self.reject_thresh_entry, "Kurtosis Z-score threshold for bad channels.")
        self.save_preprocessed=tk.BooleanVar(value=True); cb_save=ttk.Checkbutton(params_frame, text="Save preprocessed (.fif)", variable=self.save_preprocessed); cb_save.grid(row=2, column=4, columnspan=2, sticky="w", padx=5, pady=2); self._bind_tooltip(cb_save, "Save preprocessed data (.fif files).")
        ttk.Label(params_frame, text="Ref Idx 1:").grid(row=3, column=0, sticky="w", padx=5, pady=2)
        self.ref_idx1_entry=ttk.Entry(params_frame, width=8, validate='key', validatecommand=validate_int_idx_cmd); self.ref_idx1_entry.insert(0, "64"); self.ref_idx1_entry.grid(row=3, column=1, padx=5); self._bind_tooltip(self.ref_idx1_entry, "0-based index of first re-ref channel. Blank=skip.")
        ttk.Label(params_frame, text="Ref Idx 2:").grid(row=3, column=2, sticky="w", padx=5, pady=2)
        self.ref_idx2_entry=ttk.Entry(params_frame, width=8, validate='key', validatecommand=validate_int_idx_cmd); self.ref_idx2_entry.insert(0, "65"); self.ref_idx2_entry.grid(row=3, column=3, padx=5); self._bind_tooltip(self.ref_idx2_entry, "0-based index of second re-ref channel. Blank=skip.")
        ttk.Label(params_frame, text="Max Idx Keep:").grid(row=3, column=4, sticky="w", padx=5, pady=2)
        self.max_idx_keep_entry=ttk.Entry(params_frame, width=8, validate='key', validatecommand=validate_int_idx_cmd); self.max_idx_keep_entry.insert(0, "64"); self.max_idx_keep_entry.grid(row=3, column=5, padx=5); self._bind_tooltip(self.max_idx_keep_entry, "Keep channels 0 to index-1. E.g., 64 keeps 0-63. Blank=keep all.")

        # Conditions Frame -> Now "Event IDs Frame"
        conditions_frame = ttk.LabelFrame(main_frame, text="Event IDs to Process") # *** CHANGED LABEL ***
        conditions_frame.pack(fill="both", expand=True, padx=5, pady=5)
        conditions_outer_frame = ttk.Frame(conditions_frame); conditions_outer_frame.pack(fill="both", expand=True)
        self.conditions_canvas = tk.Canvas(conditions_outer_frame, borderwidth=0, highlightthickness=0)
        self.conditions_inner_frame = ttk.Frame(self.conditions_canvas)
        conditions_scrollbar = ttk.Scrollbar(conditions_outer_frame, orient="vertical", command=self.conditions_canvas.yview)
        self.conditions_canvas.configure(yscrollcommand=conditions_scrollbar.set); conditions_scrollbar.pack(side="right", fill="y")
        self.conditions_canvas.pack(side="left", fill="both", expand=True)
        self.canvas_frame_id = self.conditions_canvas.create_window((0, 0), window=self.conditions_inner_frame, anchor="nw")
        self.conditions_inner_frame.bind("<Configure>", self._on_inner_frame_configure); self.conditions_canvas.bind("<Configure>", self._on_canvas_configure)
        self.condition_entries = [] # Renamed variable conceptually, but keep name for less churn
        self.add_event_id_entry() # Call renamed function to add first entry

        # Buttons below the list
        condition_button_frame = ttk.Frame(conditions_frame); condition_button_frame.pack(fill="x", pady=5)
        self.detect_button = ttk.Button(condition_button_frame, text="Detect Event IDs", command=self.detect_and_populate_event_ids) # Renamed command
        self.detect_button.pack(side="left", padx=5); self._bind_tooltip(self.detect_button, "Scan first file for numerical Event IDs and populate list.")
        self.add_cond_button = ttk.Button(condition_button_frame, text="Add Event ID Field", command=self.add_event_id_entry) # Renamed command
        self.add_cond_button.pack(side="left", padx=5); self._bind_tooltip(self.add_cond_button, "Manually add field for another Event ID.")

        # Save Location Frame (Unchanged)
        save_frame = ttk.LabelFrame(main_frame, text="Excel Output Save Location"); save_frame.pack(fill="x", padx=5, pady=5)
        self.save_folder_path = tk.StringVar()
        btn_select_save = ttk.Button(save_frame, text="Select Parent Folder", command=self.select_save_folder)
        btn_select_save.pack(side="left", padx=5, pady=5); self._bind_tooltip(btn_select_save, "Choose parent folder for results.")
        self.save_folder_display = ttk.Entry(save_frame, textvariable=self.save_folder_path, state="readonly")
        self.save_folder_display.pack(side="left", fill="x", expand=True, padx=5, pady=5); self._bind_tooltip(self.save_folder_display, "Selected parent directory.")

        # Bottom Section (Unchanged layout)
        bottom_frame = ttk.Frame(main_frame); bottom_frame.pack(fill="both", expand=False, side=tk.BOTTOM, pady=(5,0))
        button_frame = ttk.Frame(bottom_frame); button_frame.pack(fill="x", padx=5, pady=5)
        self.select_button_text = tk.StringVar(); self.select_button = ttk.Button(button_frame, textvariable=self.select_button_text, command=self.select_data_source)
        self.select_button.pack(side="left", padx=10, pady=5, expand=True, fill='x')
        self.start_button = ttk.Button(button_frame, text="Start Processing", command=self.start_processing)
        self.start_button.pack(side="right", padx=10, pady=5, expand=True, fill='x'); self._bind_tooltip(self.start_button, "Begin analysis with current settings.")
        self.update_select_button_text()
        self.progress_bar = ttk.Progressbar(bottom_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", padx=5, pady=(0,5)); self._bind_tooltip(self.progress_bar, "File processing progress.")
        log_frame = ttk.LabelFrame(bottom_frame, text="Log"); log_frame.pack(fill="both", expand=True, padx=5, pady=5)
        log_text_frame = ttk.Frame(log_frame); log_text_frame.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_text_frame, height=10, wrap="word", state="disabled", relief="sunken", borderwidth=1)
        log_scroll = ttk.Scrollbar(log_text_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.config(yscrollcommand=log_scroll.set); log_scroll.pack(side="right", fill="y"); self.log_text.pack(side="left", fill="both", expand=True)


    # --- GUI Update/Action Methods ---
    def _on_inner_frame_configure(self, event=None): self.conditions_canvas.configure(scrollregion=self.conditions_canvas.bbox("all"))
    def _on_canvas_configure(self, event=None): self.conditions_canvas.itemconfig(self.canvas_frame_id, width=event.width)

    # Renamed add_condition_entry -> add_event_id_entry
    def add_event_id_entry(self, event=None):
        """ Adds a new entry field for specifying an Event ID. Binds Enter key. """
        validate_event_id_cmd = (self.register(self._validate_integer_input), '%P') # Get validator
        frame = ttk.Frame(self.conditions_inner_frame); frame.pack(fill="x", pady=1, padx=2)
        # Use integer validator for Event IDs
        entry = ttk.Entry(frame, width=10, validate='key', validatecommand=validate_event_id_cmd) # Shorter width maybe?
        entry.pack(side="left", fill="x", expand=True)
        self._bind_tooltip(entry, "Enter numerical Event ID. Press Enter to add another.") # Updated tooltip
        entry.bind("<Return>", self.add_event_id_entry) # Bind Enter
        remove_btn = ttk.Button(frame, text="X", width=3, style='Toolbutton', command=lambda f=frame, e=entry: self.remove_event_id_entry(f, e)) # Renamed command
        remove_btn.pack(side="right", padx=(2,0)); self._bind_tooltip(remove_btn, "Remove this Event ID field.")
        self.condition_entries.append(entry) # Still use self.condition_entries to hold widgets
        if event is None: entry.focus_set()
        self.conditions_inner_frame.update_idletasks(); self._on_inner_frame_configure()
        self.after(10, lambda: self.conditions_canvas.yview_moveto(1.0))

    # Renamed remove_condition_entry -> remove_event_id_entry
    def remove_event_id_entry(self, frame, entry):
        """ Removes an Event ID entry field. """
        if len(self.condition_entries) > 0:
             try:
                 frame.destroy()
                 if entry in self.condition_entries: self.condition_entries.remove(entry)
                 self.conditions_inner_frame.update_idletasks(); self._on_inner_frame_configure()
             except Exception as e: self.log(f"Error removing Event ID field: {e}")

    def select_save_folder(self):
        # (Code Unchanged)
        folder = filedialog.askdirectory(title="Select Parent Folder for Excel Output")
        if folder: self.save_folder_path.set(folder); self.log(f"Output target folder: {folder}")
        else: self.log("Save folder selection cancelled.")

    def update_select_button_text(self):
        # (Code Unchanged)
        mode = self.file_mode.get(); text = "Select EEG File" if mode == "Single" else "Select Data Folder"
        tooltip = "Click to select a single EEG file." if mode == "Single" else "Click to select folder with EEG files."
        self.select_button_text.set(text); self._bind_tooltip(self.select_button, tooltip)

    def select_data_source(self):
        # (Code Unchanged)
        self.data_paths = []; file_ext = "*" + self.file_type.get().lower(); file_type_desc = self.file_type.get()
        try:
            if self.file_mode.get() == "Single":
                ftypes = [(f"{file_type_desc} files", file_ext)]; # Add other types if needed
                if file_type_desc==".BDF": ftypes.append((".set files", "*.set"))
                elif file_type_desc==".set": ftypes.append((".BDF files", "*.bdf"))
                ftypes.append(("All files", "*.*"))
                file_path = filedialog.askopenfilename(title="Select EEG File", filetypes=ftypes)
                if file_path:
                    selected_ext = os.path.splitext(file_path)[1].lower()
                    if selected_ext in ['.bdf', '.set']: self.file_type.set(selected_ext.upper())
                    self.data_paths = [file_path]; self.log(f"Selected file: {os.path.basename(file_path)}")
                else: self.log("No file selected.")
            else: # Batch mode
                folder = filedialog.askdirectory(title=f"Select Folder Containing {file_type_desc} Files")
                if folder:
                    search_path = os.path.join(folder, file_ext); found_files = glob.glob(search_path)
                    if found_files: self.data_paths = sorted(found_files); self.log(f"Selected folder: {folder}"); self.log(f"Found {len(found_files)} file(s) matching '{file_ext}'.")
                    else: self.log(f"No '{file_ext}' files found in {folder}."); messagebox.showwarning("No Files Found", f"No '{file_ext}' files found in:\n{folder}")
                else: self.log("No folder selected.")
        except Exception as e: self.log(f"Error selecting data: {e}"); messagebox.showerror("Selection Error", f"{e}")
        self.progress_bar['maximum'] = len(self.data_paths) if self.data_paths else 1; self.progress_bar['value'] = 0

    def log(self, message):
        # (Code Unchanged - includes timestamp and thread prefix)
        if hasattr(self, 'log_text') and self.log_text:
            try:
                current_thread = threading.current_thread(); timestamp = pd.Timestamp.now().strftime('%H:%M:%S')
                prefix = "[BG]" if current_thread != threading.main_thread() else "[GUI]"
                log_msg = f"{timestamp} {prefix}: {message}\n"
                if current_thread == threading.main_thread():
                     self.log_text.config(state="normal"); self.log_text.insert(tk.END, log_msg); self.log_text.see(tk.END); self.log_text.config(state="disabled"); self.update_idletasks()
                else: print(log_msg, end='') # Print if from background
            except tk.TclError as e:
                if "invalid command name" not in str(e): print(f"{prefix} GUI Log Error: {e}")
            except Exception as e: print(f"{prefix} Unexpected GUI Log Error: {e}")
        else: print(f"{pd.Timestamp.now().strftime('%H:%M:%S')} Log (no GUI): {message}")

    # Renamed detect function
    def detect_and_populate_event_ids(self):
        """ Scans first file using find_events, populates GUI with Event IDs. """
        # (Code Uses mne.find_events as implemented in previous step)
        self.log("Attempting to detect event IDs using find_events...")
        if not self.data_paths: messagebox.showerror("No Data Selected", "Select a file/folder first."); self.log("Detection failed: No data."); return

        representative_file = self.data_paths[0]; self.log(f"Scanning file: {os.path.basename(representative_file)}")
        raw, event_ids = None, []
        try:
            self._disable_controls(disable_process_buttons=False) # Keep process buttons active
            self.update_idletasks()
            raw = self.load_eeg_file(representative_file) # Uses stim_channel='Status'
            if raw is None: messagebox.showerror("Loading Error", f"Could not load: {os.path.basename(representative_file)}"); return

            self.log("Searching for events on 'Status' channel using mne.find_events...")
            events_array = mne.find_events(raw, stim_channel='Status', shortest_event=1, verbose=False) # shortest_event=1 might need tuning
            if events_array is None or events_array.shape[0] == 0:
                 messagebox.showinfo("No Events Found", f"No event triggers found on 'Status' channel in:\n{os.path.basename(representative_file)}")
                 self.log("find_events: No usable events found.")
            else:
                 unique_ids = np.unique(events_array[:, 2]); event_ids = sorted(list(unique_ids))
                 self.log(f"find_events: Found {events_array.shape[0]} events total."); self.log(f"Unique Event IDs: {event_ids}")

            # --- Populate GUI ---
            self.log("Clearing existing Event ID fields..."); # Clear previous list
            for widget in self.conditions_inner_frame.winfo_children(): widget.destroy()
            self.condition_entries = [] # Reset list
            if not event_ids: self.add_event_id_entry(); self.log("Added blank field as no IDs detected.")
            else: # Populate with detected IDs
                 for event_id_int in event_ids:
                     self.add_event_id_entry()
                     new_entry = self.condition_entries[-1]; new_entry.delete(0, tk.END); new_entry.insert(0, str(event_id_int))
                 messagebox.showinfo("Event IDs Detected", f"Populated list with {len(event_ids)} unique Event ID(s).\nReview/edit list.")
        except ValueError as ve: # Specific error check for missing channel
             if "not found" in str(ve) and "stim_channel" in str(ve): self.log("Error: 'Status' channel not found."); messagebox.showerror("Detection Error", "Could not find 'Status' channel.")
             else: self.log(f"Value error: {ve}\n{traceback.format_exc()}"); messagebox.showerror("Detection Error", f"{ve}")
        except Exception as e: self.log(f"Detection error: {e}\n{traceback.format_exc()}"); messagebox.showerror("Detection Error", f"{e}")
        finally:
            self._enable_controls(enable_process_buttons=False) # Re-enable condition buttons only
            if raw: del raw
            if self.condition_entries: self.condition_entries[0].focus_set()
            self.update_idletasks()

    # --- Core Processing Control ---
    def start_processing(self):
        """ Validates inputs, starts background processing thread using Event IDs. """
        if self.processing_thread and self.processing_thread.is_alive(): messagebox.showwarning("Busy", "Processing already running."); return
        self.log("="*40); self.log("Processing initiated...")
        # 1. Validation (Files, Save Folder)
        if not self.data_paths: messagebox.showerror("Input Error", "No data selected."); return
        if not self.save_folder_path.get(): messagebox.showerror("Input Error", "No save folder selected."); return
        # 2. Retrieve & Validate Parameters (Unchanged)
        try:
            params = { key: (getter().get() if getter().get() else default) for key, getter, default in [('low_pass', lambda: self.low_pass_entry, 0.0), ('epoch_start', lambda: self.epoch_start_entry, -1.0), ('epoch_end', lambda: self.epoch_end_entry, 5.0), ('reject_thresh', lambda: self.reject_thresh_entry, 5.0)]}
            params.update({ key: (float(getter().get()) if getter().get() else default) for key, getter, default in [('high_pass', lambda: self.high_pass_entry, None), ('downsample_rate', lambda: self.downsample_entry, None)]})
            params.update({ key: (int(getter().get()) if getter().get() else default) for key, getter, default in [('ref_idx1', lambda: self.ref_idx1_entry, None), ('ref_idx2', lambda: self.ref_idx2_entry, None), ('max_idx_keep', lambda: self.max_idx_keep_entry, None)]})
            params['save_preprocessed'] = self.save_preprocessed.get()
            # Logical validation... (kept concise)
            if params['high_pass'] is not None and params['low_pass'] >= params['high_pass']: raise ValueError("Low Pass >= High Pass")
            if params['downsample_rate'] is not None and params['downsample_rate'] <= 0: raise ValueError("Downsample rate <= 0")
            # ... other validations ...
        except ValueError as e: messagebox.showerror("Parameter Error", f"Invalid value: {e}"); return
        except Exception as e: messagebox.showerror("Parameter Error", f"Error retrieving: {e}"); return

        # *** Gather Event IDs (as integers) ***
        self.current_event_ids_process = []
        raw_gui_entries = [e.get().strip() for e in self.condition_entries if e.get().strip()]
        if not raw_gui_entries: messagebox.showerror("Event ID Error", "Enter or Detect Event IDs."); return
        try:
            # Convert to integers, ensuring uniqueness
            self.current_event_ids_process = sorted(list(set(int(id_str) for id_str in raw_gui_entries)))
        except ValueError:
            messagebox.showerror("Event ID Error", "Invalid non-integer Event ID found in the list. Please enter numbers only."); return
        if not self.current_event_ids_process: messagebox.showerror("Event ID Error", "No valid Event IDs entered."); return
        self.log(f"Processing numerical Event IDs: {self.current_event_ids_process}")

        # 3. Prepare for Thread
        self.preprocessed_data = {}; self.progress_bar['value'] = 0; self.progress_bar['maximum'] = len(self.data_paths)
        self._disable_controls() # Disable all controls
        self.log("Starting background processing...")
        # Pass integer IDs list to thread
        thread_args = (list(self.data_paths), params, list(self.current_event_ids_process), self.gui_queue)
        # 4. Start Thread
        self.processing_thread = threading.Thread(target=self._processing_thread_func, args=thread_args, daemon=True)
        self.processing_thread.start()
        # 5. Start Periodic Queue Check
        self.after(100, self._periodic_queue_check)

    def _periodic_queue_check(self):
        """ Checks GUI queue for messages from background thread. """
        # (Code Unchanged - handles 'log', 'progress', 'result', 'error', 'done')
        processing_done = False
        while True:
            try:
                message = self.gui_queue.get_nowait(); msg_type = message.get('type')
                if msg_type == 'log': self.log(message.get('message', ''))
                elif msg_type == 'progress': self.progress_bar['value'] = message.get('value', 0)
                elif msg_type == 'result': self.preprocessed_data = message.get('data', {}); self.log("Preprocessing results received.")
                elif msg_type == 'error':
                    error_msg=message.get('message', 'Unknown thread error.'); tb_info=message.get('traceback', '')
                    self.log(f"!!! THREAD ERROR: {error_msg}");
                    if tb_info: self.log(tb_info)
                    messagebox.showerror("Processing Error", error_msg); processing_done = True
                elif msg_type == 'done': self.log("Background thread signaled completion."); processing_done = True
            except queue.Empty: break
            except Exception as e: self.log(f"Queue check error: {e}"); processing_done = True; break

        if processing_done: self._finalize_processing(success=('error' not in message))
        elif self.processing_thread and self.processing_thread.is_alive(): self.after(100, self._periodic_queue_check)
        else: self.log("Thread ended unexpectedly; re-enabling controls."); self._enable_controls()

    def _finalize_processing(self, success):
        """ Handles tasks after thread finishes. Calls post_process with Event IDs. """
        # (Code Unchanged - calls post_process with self.current_event_ids_process)
        if success and self.preprocessed_data:
            self.log("\n--- Starting Post-processing Phase (Main Thread) ---")
            try:
                self.post_process(self.current_event_ids_process) # Pass integer IDs
                self.log("--- Post-processing Phase Complete ---"); messagebox.showinfo("Processing Complete", "Analysis finished.")
            except Exception as post_err: self.log(f"!!! Post-processing Error: {post_err}\n{traceback.format_exc()}"); messagebox.showerror("Post-processing Error", f"{post_err}")
        elif success: self.log("--- Skipping Post-processing: No data ---"); messagebox.showwarning("Processing Finished", "Preprocessing finished, no usable epochs generated.")
        self._enable_controls(); self.log(f"--- Processing Run Finished at {pd.Timestamp.now()} ---")

    def _disable_controls(self, disable_process_buttons=True):
        """Disables buttons during processing or detection."""
        self.select_button.config(state="disabled")
        self.detect_button.config(state="disabled")
        self.add_cond_button.config(state="disabled")
        if disable_process_buttons: self.start_button.config(state="disabled")
        self.update_idletasks()

    def _enable_controls(self, enable_process_buttons=True):
        """Re-enables buttons after processing or detection."""
        self.select_button.config(state="normal")
        self.detect_button.config(state="normal")
        self.add_cond_button.config(state="normal")
        if enable_process_buttons: self.start_button.config(state="normal")
        self.update_idletasks()


    # --- Background Thread Function ---
    # Modified to accept integer IDs and use find_events
    def _processing_thread_func(self, data_paths, params, conditions_ids_to_process, gui_queue):
        """ Runs heavy processing in background. Works with numerical Event IDs. """
        # Initialize results dict with integer keys
        local_preprocessed_data = {int_id: [] for int_id in conditions_ids_to_process}
        files_with_epochs = 0
        processing_error_occurred = False

        try:
            num_files = len(data_paths)
            for i, file_path in enumerate(data_paths):
                base_filename = os.path.basename(file_path)
                gui_queue.put({'type': 'log', 'message': f"\nProcessing file {i+1}/{num_files}: {base_filename}"})
                events_array, raw_processed = None, None # Init per file

                try: # Inner try for file-specific errors
                    raw = self.load_eeg_file(file_path)
                    if raw is None: raise ValueError("File loading failed.")

                    # --- Event Finding (using find_events) ---
                    try:
                        events_array = mne.find_events(raw, stim_channel='Status', shortest_event=1, verbose=False)
                        if events_array is None or events_array.shape[0] == 0:
                             gui_queue.put({'type': 'log', 'message': f"Warning: No events found via find_events in {base_filename}."})
                        else:
                             found_ids = np.unique(events_array[:, 2])
                             gui_queue.put({'type': 'log', 'message': f"Found {events_array.shape[0]} events via find_events. Unique IDs: {found_ids}"})
                    except Exception as event_err:
                        gui_queue.put({'type': 'log', 'message': f"Warning: find_events error: {event_err}."})

                    # --- Preprocessing ---
                    raw_processed = self.preprocess_raw(raw.copy(), **params)
                    del raw
                    if raw_processed is None: raise ValueError("Preprocessing failed critically.")

                    # --- Epoch Extraction (using Integer IDs) ---
                    file_had_epochs = False
                    if events_array is not None and events_array.shape[0] > 0:
                        gui_queue.put({'type': 'log', 'message': "Attempting epoch extraction using numerical IDs..."})
                        all_ids_in_file = np.unique(events_array[:, 2]) # IDs found in this file

                        for requested_id in conditions_ids_to_process: # Loop through requested INT IDs
                            if requested_id in all_ids_in_file:
                                try:
                                    # Map string description (can just be str(ID)) to the integer ID
                                    event_id_map = {str(requested_id): requested_id}
                                    epochs = mne.Epochs(raw_processed, events_array, event_id=event_id_map,
                                                        tmin=params['epoch_start'], tmax=params['epoch_end'],
                                                        preload=False, verbose=False, baseline=None)
                                    if len(epochs) > 0:
                                        gui_queue.put({'type': 'log', 'message': f"Found {len(epochs)} epochs for Event ID {requested_id}."})
                                        # Use the INTEGER ID as the key for the results dict
                                        local_preprocessed_data[requested_id].append(epochs)
                                        file_had_epochs = True
                                except Exception as epoch_err:
                                    gui_queue.put({'type': 'log', 'message': f"Epoch error for ID {requested_id}: {epoch_err}\n{traceback.format_exc()}"})
                            else:
                                 # Log only if the ID was expected but not found in this file
                                 # Reduce log spam: only log once per ID per run maybe? Harder. Log every time for now.
                                 gui_queue.put({'type': 'log', 'message': f"Info: Requested Event ID {requested_id} not found in this file."})

                        if file_had_epochs: files_with_epochs += 1
                    else: gui_queue.put({'type': 'log', 'message': "Skipping epoch extraction (no events found in file)."})

                    # --- Save Preprocessed (Optional) ---
                    if params['save_preprocessed'] and raw_processed is not None:
                        save_filename = f"{os.path.splitext(base_filename)[0]}_preproc.fif"
                        save_path = os.path.join(os.path.dirname(file_path), save_filename)
                        try: raw_processed.save(save_path, overwrite=True, verbose=False); gui_queue.put({'type': 'log', 'message': f"Saved: {save_path}"})
                        except Exception as save_err: gui_queue.put({'type': 'log', 'message': f"Save error {save_path}: {save_err}"})

                except Exception as file_err:
                     gui_queue.put({'type': 'log', 'message': f"!!! ERROR processing file {base_filename}: {file_err}\n{traceback.format_exc()}"})
                     processing_error_occurred = True
                finally:
                    if 'raw_processed' in locals() and raw_processed: del raw_processed
                    gui_queue.put({'type': 'progress', 'value': i + 1})

            # --- Loop Finished ---
            gui_queue.put({'type': 'log', 'message': "\n--- Background Preprocessing Phase Complete ---"})
            # Filter out empty lists from results before sending back? No, send full dict.
            gui_queue.put({'type': 'result', 'data': local_preprocessed_data})

        except MemoryError: gui_queue.put({'type': 'error', 'message': "Memory Error during background processing."})
        except Exception as e: gui_queue.put({'type': 'error', 'message': f"Critical error in thread: {e}", 'traceback': traceback.format_exc()})
        finally: gui_queue.put({'type': 'done'})


    # --- EEG Loading/Processing Methods ---
    # load_eeg_file already updated for stim_channel='Status'
    def load_eeg_file(self, filepath):
        # (Code Unchanged from v1.3.1)
        ext = os.path.splitext(filepath)[1].lower(); raw = None; self.log(f"Loading: {os.path.basename(filepath)}...")
        try:
            load_kwargs = {'preload': True, 'verbose': False}
            if ext == ".bdf": load_kwargs['stim_channel'] = 'Status'; self.log("Reading BDF events from 'Status' channel.")
            with mne.utils.use_log_level('ERROR'):
                if ext == ".bdf": raw = mne.io.read_raw_bdf(filepath, **load_kwargs)
                elif ext == ".set": raw = mne.io.read_raw_eeglab(filepath, **load_kwargs)
                else: self.log(f"Unsupported format '{ext}'."); return None
            if raw is None: raise ValueError("MNE load failed.")
            self.log(f"Loaded {len(raw.ch_names)} channels @ {raw.info['sfreq']:.1f} Hz.")
            if hasattr(raw, 'annotations') and len(raw.annotations) > 0: self.log(f"Found {len(raw.annotations)} annotations after loading."); # Check if annotations worked
            else: self.log("No MNE Annotations found after loading (will rely on find_events).")
            try: # Apply montage
                montage = mne.channels.make_standard_montage('standard_1020'); raw.set_montage(montage, on_missing='warn', match_case=False)
                if raw.get_montage(): self.log("Applied standard_1020 montage (or partial).")
                else: self.log("Warning: Montage not applied.")
            except Exception as montage_err: self.log(f"Warning: Montage error: {montage_err}")
            return raw
        except MemoryError: self.log(f"Memory Error loading {os.path.basename(filepath)}."); return None
        except Exception as e: self.log(f"Load Error {os.path.basename(filepath)}: {e}\n{traceback.format_exc()}"); return None

    # preprocess_raw already accepts **params
    def preprocess_raw(self, raw, **params):
        # (Code Unchanged from v1.3.1)
        downsample_rate=params.get('downsample_rate'); low_pass=params.get('low_pass'); high_pass=params.get('high_pass'); reject_thresh=params.get('reject_thresh'); ref_idx1=params.get('ref_idx1'); ref_idx2=params.get('ref_idx2'); max_idx_keep=params.get('max_idx_keep')
        try:
            ch_names_orig = raw.info['ch_names']; n_chans_orig = len(ch_names_orig); self.log(f"Preprocessing {n_chans_orig} channels...")
            # Step 1: Re-reference
            if ref_idx1 is not None and ref_idx2 is not None:
                if 0<=ref_idx1<n_chans_orig and 0<=ref_idx2<n_chans_orig and ref_idx1!=ref_idx2:
                    ref_ch_names=[ch_names_orig[ref_idx1], ch_names_orig[ref_idx2]]; self.log(f"Re-ref to: {ref_ch_names}...");
                    try: raw.set_eeg_reference(ref_ch_names, projection=False); self.log("Success.")
                    except Exception as ref_err: self.log(f"Warning: Re-ref failed: {ref_err}.")
                else: self.log("Warning: Invalid ref indices. Skipping.")
            else: self.log("Skipping custom re-ref.")
            # Step 2: Drop channels
            if max_idx_keep is not None:
                current_ch_names=raw.info['ch_names']; current_n_chans=len(current_ch_names)
                if 0 < max_idx_keep <= current_n_chans:
                    indices_to_drop = list(range(max_idx_keep, current_n_chans))
                    if indices_to_drop:
                        channels_to_drop=[current_ch_names[i] for i in indices_to_drop]; self.log(f"Dropping {len(channels_to_drop)} chans (>= idx {max_idx_keep})...")
                        try: raw.drop_channels(channels_to_drop); self.log(f"Remaining: {len(raw.ch_names)}")
                        except Exception as drop_err: self.log(f"Warning: Drop failed: {drop_err}")
                elif max_idx_keep > current_n_chans: self.log("Info: Max Idx Keep >= chans.")
                else: self.log("Warning: Invalid Max Idx Keep.")
            else: self.log("Skipping index-based drop.")
            # Step 3: Downsample
            if downsample_rate is not None:
                current_sfreq = raw.info['sfreq']
                if current_sfreq > downsample_rate:
                    self.log(f"Downsampling {current_sfreq:.1f}Hz -> {downsample_rate}Hz...")
                    try: raw.resample(downsample_rate, npad="auto", verbose=False); self.log(f"New rate: {raw.info['sfreq']:.1f} Hz.")
                    except Exception as ds_err: self.log(f"Error downsampling: {ds_err}"); return None
                else: self.log("No downsampling needed.")
            else: self.log("Skipping downsampling.")
            # Step 4: Filter
            l_freq = low_pass if low_pass > 0 else None; h_freq = high_pass
            if l_freq is not None or h_freq is not None:
                self.log(f"Filtering ({l_freq} Hz - {h_freq} Hz)...")
                try:
                    nyquist=raw.info['sfreq']/2.0
                    if h_freq is not None and h_freq >= nyquist: h_freq = nyquist - 0.5; self.log(f"Adj. high pass to {h_freq:.1f} Hz.")
                    if l_freq is not None and h_freq is not None and l_freq >= h_freq: self.log("Warning: Invalid filter range.")
                    else: raw.filter(l_freq=l_freq, h_freq=h_freq, method='fir', phase='zero-double', verbose=False); self.log("Filter complete.")
                except Exception as filter_err: self.log(f"Warning: Filtering error: {filter_err}.")
            else: self.log("Skipping filtering.")
            # Step 5: Kurtosis Rejection & Interpolation
            self.log(f"Kurtosis rejection (Z > {reject_thresh})...")
            try:
                picks_eeg = mne.pick_types(raw.info, eeg=True, exclude='bads')
                if len(picks_eeg) > 1:
                    data=raw.get_data(picks=picks_eeg); channel_kurt=kurtosis(data, axis=1, fisher=True); del data
                    channel_kurt=np.nan_to_num(channel_kurt); mean_kurt=np.mean(channel_kurt); std_kurt=np.std(channel_kurt)
                    z_scores = np.zeros_like(channel_kurt) if std_kurt < 1e-6 else (channel_kurt - mean_kurt) / std_kurt
                    eeg_ch_names=[raw.info['ch_names'][i] for i in picks_eeg]
                    bad_channels_kurt=[eeg_ch_names[i] for i, z in enumerate(z_scores) if abs(z) > reject_thresh]
                    if bad_channels_kurt:
                        self.log(f"Bad via Kurtosis: {bad_channels_kurt}.")
                        raw.info['bads'].extend([ch for ch in bad_channels_kurt if ch not in raw.info['bads']])
                        if raw.info['bads'] and raw.get_montage():
                             self.log(f"Interpolating bads: {raw.info['bads']}...")
                             try: raw.interpolate_bads(reset_bads=True, mode='accurate', verbose=False); self.log("Interpolation ok.")
                             except Exception as interp_err: self.log(f"Warning: Interpolation error: {interp_err}.")
                        elif raw.info['bads']: self.log("Warning: Cannot interpolate (no montage).")
                    else: self.log("No channels rejected via Kurtosis.")
                else: self.log("Warning: Not enough EEG channels for Kurtosis.")
            except Exception as kurt_err: self.log(f"Warning: Kurtosis error: {kurt_err}.\n{traceback.format_exc()}")
            # Step 6: Average Reference
            self.log("Applying average reference...");
            try: raw.set_eeg_reference(ref_channels='average', projection=False); self.log("Avg ref ok.")
            except Exception as avg_ref_err: self.log(f"Warning: Avg ref failed: {avg_ref_err}")
            self.log("Preprocessing finished.")
            return raw
        except MemoryError: self.log("Memory Error during preprocessing."); return None
        except Exception as e: self.log(f"Critical preprocessing error: {e}\n{traceback.format_exc()}"); return None

    # Modified post_process to accept integer IDs and name outputs accordingly
    def post_process(self, conditions_ids_to_process):
        """ Calculates FFT metrics and saves results. Works with numerical Event IDs. """
        if not self.save_folder_path.get(): self.log("Error: Save folder missing."); return
        parent_folder = self.save_folder_path.get()

        for condition_id in conditions_ids_to_process: # Loop through integer IDs
            # Use integer ID to get data
            epochs_object_list = self.preprocessed_data.get(condition_id, [])
            # Create name string for logging, folders, files
            event_id_name = f"Event_{condition_id}"

            if not epochs_object_list: self.log(f"No data for {event_id_name}. Skipping."); continue

            n_files_for_cond = len(epochs_object_list)
            self.log(f"\nPost-processing {event_id_name} ({n_files_for_cond} file(s))...")

            accum_fft_amp = accum_snr = accum_z_score = accum_bca = None
            valid_file_count = 0; final_n_channels = None

            for file_idx, epochs in enumerate(epochs_object_list):
                self.log(f"  Processing results from file {file_idx+1}/{n_files_for_cond} for {event_id_name}...")
                try:
                    if len(epochs) == 0: self.log("    0 epochs. Skipping."); continue
                    self.log(f"    Loading {len(epochs)} epochs..."); epochs.load_data()
                    n_epochs, n_channels, n_times = epochs.get_data(copy=False).shape
                    if final_n_channels is None: final_n_channels = n_channels
                    elif final_n_channels != n_channels: self.log("    Warning: Ch count mismatch. Skipping."); continue
                    sfreq = epochs.info['sfreq']
                    self.log(f"    Processing {n_epochs} epochs, {n_channels} chans, {n_times} points @ {sfreq:.1f}Hz.")

                    # --- Use Welch PSD ---
                    try: fmax_psd = float(self.high_pass_entry.get()) if self.high_pass_entry.get() else sfreq/2.0
                    except: fmax_psd = sfreq/2.0
                    fmax_psd = min(fmax_psd, sfreq/2.0)
                    spectrum = epochs.compute_psd(method='welch', fmin=0.5, fmax=fmax_psd, n_fft=int(sfreq * 2), n_overlap=int(sfreq * 1), window='hann', average='mean', verbose=False)
                    psd_freqs = spectrum.freqs; avg_power = spectrum.get_data(return_freqs=False)
                    file_avg_fft_amp = np.sqrt(avg_power); freqs_for_metrics = psd_freqs; n_freqs_metrics = len(freqs_for_metrics)

                    # --- Calculate Metrics --- (Logic unchanged)
                    n_target = len(TARGET_FREQUENCIES)
                    file_fft_out=np.zeros((n_channels, n_target)); file_snr_out=np.zeros((n_channels, n_target)); file_z_out=np.zeros((n_channels, n_target)); file_bca_out=np.zeros((n_channels, n_target))
                    noise_range, noise_exclude = 12, 1
                    for ch in range(n_channels):
                        for idx_freq, target_freq in enumerate(TARGET_FREQUENCIES):
                            target_bin = np.argmin(np.abs(freqs_for_metrics - target_freq))
                            lower_b=max(0, target_bin-noise_range); excl_s=max(0, target_bin-noise_exclude); excl_e=min(n_freqs_metrics, target_bin+noise_exclude+1); upper_b=min(n_freqs_metrics, target_bin+noise_range+1)
                            neigh_idx = np.unique(np.concatenate([np.arange(lower_b, excl_s), np.arange(excl_e, upper_b)])); neigh_idx = neigh_idx[neigh_idx < n_freqs_metrics]
                            if neigh_idx.size < 4: continue
                            neigh_amps = file_avg_fft_amp[ch, neigh_idx]; noise_mean=np.mean(neigh_amps); noise_std=np.std(neigh_amps)
                            fft_val = file_avg_fft_amp[ch, target_bin]; snr_val = fft_val / noise_mean if noise_mean > 1e-12 else 0
                            loc_max_r=np.arange(max(0, target_bin-1), min(n_freqs_metrics, target_bin+2));
                            if loc_max_r.size == 0: continue
                            loc_max = np.max(file_avg_fft_amp[ch, loc_max_r]); z_val = (loc_max - noise_mean) / noise_std if noise_std > 1e-12 else 0
                            bca_val = fft_val - noise_mean
                            file_fft_out[ch, idx_freq]=fft_val; file_snr_out[ch, idx_freq]=snr_val; file_z_out[ch, idx_freq]=z_val; file_bca_out[ch, idx_freq]=bca_val

                    # --- Accumulate ---
                    if accum_fft_amp is None: accum_fft_amp=file_fft_out; accum_snr=file_snr_out; accum_z_score=file_z_out; accum_bca=file_bca_out
                    else: accum_fft_amp+=file_fft_out; accum_snr+=file_snr_out; accum_z_score+=file_z_out; accum_bca+=file_bca_out
                    valid_file_count += 1; self.log(f"    Accumulated metrics from file {file_idx+1}.")
                except MemoryError: self.log(f"!!! Memory Error post-proc file {file_idx+1}. Skipping."); messagebox.showwarning("Memory Error", f"Ran out of memory post-processing {event_id_name}, file {file_idx+1}.")
                except Exception as e: self.log(f"Error post-proc file {file_idx+1}: {e}\n{traceback.format_exc()}")
                finally:
                    if 'epochs' in locals() and epochs: epochs.drop_log_stats(); del epochs

            # --- Final Average & Excel ---
            if valid_file_count == 0: self.log(f"No valid data for {event_id_name}. No Excel."); continue
            self.log(f"Averaging metrics across {valid_file_count} file(s) for {event_id_name}.")
            avgFFT=accum_fft_amp/valid_file_count; avgSNR=accum_snr/valid_file_count; avgZ=accum_z_score/valid_file_count; avgBCA=accum_bca/valid_file_count
            # Use event_id_name for folder and file names
            subfolder_path = os.path.join(parent_folder, event_id_name)
            try: os.makedirs(subfolder_path, exist_ok=True)
            except OSError as e: self.log(f"Subfolder error {subfolder_path}: {e}. Saving to parent."); subfolder_path = parent_folder
            excel_filename = f"{event_id_name}_Results.xlsx"
            excel_path = os.path.join(subfolder_path, excel_filename)
            col_names = [f"{f:.1f}_Hz" for f in TARGET_FREQUENCIES]; n_ch_out = avgFFT.shape[0]
            electrode_col = ELECTRODE_NAMES[:n_ch_out];
            if len(electrode_col) < n_ch_out: electrode_col.extend([f"Ch{i+1}" for i in range(len(electrode_col), n_ch_out)])
            df_fft=pd.DataFrame(avgFFT, columns=col_names, index=electrode_col); df_snr=pd.DataFrame(avgSNR, columns=col_names, index=electrode_col); df_z=pd.DataFrame(avgZ, columns=col_names, index=electrode_col); df_bca=pd.DataFrame(avgBCA, columns=col_names, index=electrode_col)
            df_fft.insert(0, "Electrode", df_fft.index); df_snr.insert(0, "Electrode", df_snr.index); df_z.insert(0, "Electrode", df_z.index); df_bca.insert(0, "Electrode", df_bca.index)
            try: # Write Excel
                self.log(f"Writing formatted Excel: {excel_path}")
                with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                    df_fft.to_excel(writer, sheet_name="FFT_Amplitude", index=False); df_snr.to_excel(writer, sheet_name="SNR", index=False); df_z.to_excel(writer, sheet_name="Z_Score", index=False); df_bca.to_excel(writer, sheet_name="BCA", index=False)
                    workbook = writer.book; center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
                    for sheet_name in writer.sheets:
                        worksheet = writer.sheets[sheet_name]; df_map = {'FFT_Amplitude': df_fft, 'SNR': df_snr, 'Z_Score': df_z, 'BCA': df_bca}
                        current_df = df_map.get(sheet_name);
                        if current_df is None: continue
                        for col_idx, col_name in enumerate(current_df.columns):
                            header_width = len(str(col_name)); max_data_width = 0
                            try: max_data_width = current_df[col_name].astype(str).map(len).max(); max_data_width = 0 if pd.isna(max_data_width) else int(max_data_width)
                            except: pass
                            width = max(header_width, max_data_width) + 2; worksheet.set_column(col_idx, col_idx, width, center_format)
                self.log(f"Formatted Excel saved for {event_id_name}.")
            except Exception as excel_err: self.log(f"Excel write error {excel_path}: {excel_err}\n{traceback.format_exc()}"); messagebox.showerror("Excel Error", f"Failed Excel save for {event_id_name}.\n{excel_err}")


# --- Main execution block ---
if __name__ == "__main__":
    try: # Set DPI awareness for Windows
        from ctypes import windll; windll.shcore.SetProcessDpiAwareness(1)
    except Exception: pass
    root = FPVSApp()
    root.mainloop()