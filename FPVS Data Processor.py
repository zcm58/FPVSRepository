#!/usr/bin/env python3
"""
EEG FPVS Analysis GUI using MNE-Python and Tkinter.

This script provides a graphical user interface (GUI) to load, preprocess,
and analyze Frequency-Following Potential (FPVS) EEG data.

Version: 1.3 (April 2025) - Added Condition Detection

Key functionalities:
- Load EEG data in .BDF or .set format.
- Process single files or batch process folders.
- Button to automatically detect event conditions from the first selected file.
- Apply preprocessing steps (configurable indices, downsampling, filtering,
  Kurtosis rejection, referencing). Runs in a background thread.
- Extract epochs based on user-defined/detected condition names (Case Sensitive).
  Uses preload=False for memory efficiency.
- Perform post-processing using FFT (FFT Amp, SNR, Z-score, BCA), averaging
  in the frequency domain. Runs after preprocessing is complete.
- Save results to formatted Excel files (using xlsxwriter).
- Provides responsive logging and progress feedback via threading & queue.
- Includes tooltips, menu bar, and Enter key binding for conditions.
"""

# === Dependencies ===
# Standard Libraries:
import os
import glob
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sys
import traceback
import threading # For running processing in background
import queue     # For communicating between threads

# Third-Party Libraries:
import numpy as np
import pandas as pd
from scipy.stats import kurtosis
try:
    import mne # For EEG data processing
except ImportError:
    messagebox.showerror("Dependency Error", "MNE-Python is required. pip install mne")
    sys.exit(1)

try:
    import xlsxwriter
except ImportError:
    messagebox.showerror("Dependency Error", "XlsxWriter is required. pip install xlsxwriter")
    sys.exit(1)
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

    Handles GUI, user input, background processing thread, thread communication,
    condition detection, and post-processing.
    """
    def __init__(self):
        """
        Initialize the application window, variables, widgets, menu, and GUI queue.
        """
        super().__init__()
        self.title(f"EEG FPVS Analysis Tool (v1.3 - {pd.Timestamp.now().year})")
        self.geometry("950x900")

        # Data structures
        self.preprocessed_data = {}    # Populated from thread results
        self.condition_names_gui = []  # Stores final list used for processing (original case)
        self.current_conditions_process = [] # List used during processing run (original case)
        self.data_paths = []
        self.tooltip_map = {}
        self.processing_thread = None

        # Queue for thread communication
        self.gui_queue = queue.Queue()

        # Create Menu Bar
        self.create_menu()

        # Create Widgets (Status bar needs to exist before tooltips are bound)
        self.create_widgets()

    # --- Menu Methods ---
    def create_menu(self):
        """Creates the main menu bar for the application."""
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Select Parent Save Folder...", command=self.select_save_folder)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About...", command=self.show_about_dialog)

    def show_about_dialog(self):
        """Displays a simple 'About' dialog box."""
        messagebox.showinfo(
            "About EEG FPVS Analysis Tool",
            f"Version: 1.3 ({pd.Timestamp.now().strftime('%B %Y')})\n\n"
            "Processes and analyzes FPVS EEG data using MNE-Python.\n"
            "Features condition detection and background processing."
        )

    def quit(self):
        """Exits the application cleanly, checking for running thread."""
        if self.processing_thread and self.processing_thread.is_alive():
            if messagebox.askyesno("Exit Confirmation", "Processing is ongoing. Stop and exit?"):
                # Ideally, signal thread to stop here if implemented.
                # For now, just exit (daemon thread will be terminated)
                self.destroy()
            else: return # Don't exit if user cancels
        else:
             self.destroy()

    # --- Validation Methods ---
    def _validate_numeric_input(self, P):
        """Validation function for numeric (float) entry fields."""
        if P == "" or P == "-": return True
        try: float(P); return True
        except ValueError: self.bell(); return False

    def _validate_integer_input(self, P):
        """Validation function for non-negative integer entry fields."""
        if P == "": return True
        try:
            if int(P) >= 0: return True
            else: self.bell(); return False
        except ValueError: self.bell(); return False

    # --- Tooltip Methods ---
    def _show_tooltip(self, event):
        """Callback to display tooltip text in the status bar."""
        widget = event.widget
        try:
            if widget.winfo_exists() and hasattr(self, 'status_bar_label'):
                 tooltip_text = self.tooltip_map.get(widget, "")
                 self.status_bar_label.config(text=tooltip_text)
        except tk.TclError: pass

    def _clear_tooltip(self, event):
        """Callback to clear tooltip text from the status bar."""
        if hasattr(self, 'status_bar_label'):
            self.status_bar_label.config(text="")

    def _bind_tooltip(self, widget, text):
        """Associates a widget with its tooltip text and binds hover events."""
        self.tooltip_map[widget] = text
        widget.bind("<Enter>", self._show_tooltip, add='+')
        widget.bind("<Leave>", self._clear_tooltip, add='+')

    # --- GUI Creation ---
    def create_widgets(self):
        """
        Builds and arranges all the GUI components within the main window.
        """
        validate_num_cmd = (self.register(self._validate_numeric_input), '%P')
        validate_int_cmd = (self.register(self._validate_integer_input), '%P')

        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill="both", expand=True)

        # Create Status Bar EARLY for tooltips
        self.status_bar_label = ttk.Label(self, text="", relief="sunken", anchor="w")
        self.status_bar_label.pack(side="bottom", fill="x")

        # Top Options Frame
        options_frame = ttk.LabelFrame(main_frame, text="Processing Options")
        options_frame.pack(fill="x", padx=5, pady=5)
        self.file_mode = tk.StringVar(value="Single")
        ttk.Label(options_frame, text="Mode:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.radio_single = ttk.Radiobutton(options_frame, text="Single File", variable=self.file_mode, value="Single", command=self.update_select_button_text)
        self.radio_single.grid(row=0, column=1, padx=5); self._bind_tooltip(self.radio_single, "Process one selected EEG file.")
        self.radio_batch = ttk.Radiobutton(options_frame, text="Batch Folder", variable=self.file_mode, value="Batch", command=self.update_select_button_text)
        self.radio_batch.grid(row=0, column=2, padx=5); self._bind_tooltip(self.radio_batch, "Process all matching EEG files in a selected folder.")
        self.file_type = tk.StringVar(value=".BDF")
        ttk.Label(options_frame, text="File Type:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        rb_bdf = ttk.Radiobutton(options_frame, text=".BDF", variable=self.file_type, value=".BDF")
        rb_bdf.grid(row=1, column=1, padx=5); self._bind_tooltip(rb_bdf, "Select if processing BioSemi Data Format files.")
        rb_set = ttk.Radiobutton(options_frame, text=".set", variable=self.file_type, value=".set")
        rb_set.grid(row=1, column=2, padx=5); self._bind_tooltip(rb_set, "Select if processing EEGLAB (.set) files.")

        # Preprocessing Parameters Frame
        params_frame = ttk.LabelFrame(main_frame, text="Preprocessing Parameters")
        params_frame.pack(fill="x", padx=5, pady=5)
        # (Entries created and tooltips bound as before)
        ttk.Label(params_frame, text="Low Pass (Hz):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.low_pass_entry = ttk.Entry(params_frame, width=8, validate='key', validatecommand=validate_num_cmd)
        self.low_pass_entry.insert(0, "0.1"); self.low_pass_entry.grid(row=0, column=1, padx=5)
        self._bind_tooltip(self.low_pass_entry, "Low cutoff frequency for bandpass filter (Hz).")
        ttk.Label(params_frame, text="High Pass (Hz):").grid(row=0, column=2, sticky="w", padx=5, pady=2)
        self.high_pass_entry = ttk.Entry(params_frame, width=8, validate='key', validatecommand=validate_num_cmd)
        self.high_pass_entry.insert(0, "50"); self.high_pass_entry.grid(row=0, column=3, padx=5)
        self._bind_tooltip(self.high_pass_entry, "High cutoff frequency for bandpass filter (Hz).")
        ttk.Label(params_frame, text="Downsample (Hz):").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.downsample_entry = ttk.Entry(params_frame, width=8, validate='key', validatecommand=validate_num_cmd)
        self.downsample_entry.insert(0, "256"); self.downsample_entry.grid(row=1, column=1, padx=5)
        self._bind_tooltip(self.downsample_entry, "Target sampling rate (Hz). Applied if original rate is higher.")
        ttk.Label(params_frame, text="Epoch Start (s):").grid(row=1, column=2, sticky="w", padx=5, pady=2)
        self.epoch_start_entry = ttk.Entry(params_frame, width=8, validate='key', validatecommand=validate_num_cmd)
        self.epoch_start_entry.insert(0, "-1"); self.epoch_start_entry.grid(row=1, column=3, padx=5)
        self._bind_tooltip(self.epoch_start_entry, "Start time of epoch relative to event onset (seconds).")
        ttk.Label(params_frame, text="Epoch End (s):").grid(row=1, column=4, sticky="w", padx=5, pady=2)
        self.epoch_end_entry = ttk.Entry(params_frame, width=8, validate='key', validatecommand=validate_num_cmd)
        self.epoch_end_entry.insert(0, "5"); self.epoch_end_entry.grid(row=1, column=5, padx=5)
        self._bind_tooltip(self.epoch_end_entry, "End time of epoch relative to event onset (seconds).")
        ttk.Label(params_frame, text="Rejection Z-Thresh:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.reject_thresh_entry = ttk.Entry(params_frame, width=8, validate='key', validatecommand=validate_num_cmd)
        self.reject_thresh_entry.insert(0, "5"); self.reject_thresh_entry.grid(row=2, column=1, padx=5)
        self._bind_tooltip(self.reject_thresh_entry, "Z-score threshold for Kurtosis bad channel rejection.")
        self.save_preprocessed = tk.BooleanVar(value=True)
        cb_save = ttk.Checkbutton(params_frame, text="Save preprocessed (.fif)", variable=self.save_preprocessed)
        cb_save.grid(row=2, column=4, columnspan=2, sticky="w", padx=5, pady=2)
        self._bind_tooltip(cb_save, "Save preprocessed data (.fif files) next to originals.")
        ttk.Label(params_frame, text="Ref Idx 1:").grid(row=3, column=0, sticky="w", padx=5, pady=2)
        self.ref_idx1_entry = ttk.Entry(params_frame, width=8, validate='key', validatecommand=validate_int_cmd)
        self.ref_idx1_entry.insert(0, "64"); self.ref_idx1_entry.grid(row=3, column=1, padx=5)
        self._bind_tooltip(self.ref_idx1_entry, "0-based index of first channel for initial re-ref. Blank=skip.")
        ttk.Label(params_frame, text="Ref Idx 2:").grid(row=3, column=2, sticky="w", padx=5, pady=2)
        self.ref_idx2_entry = ttk.Entry(params_frame, width=8, validate='key', validatecommand=validate_int_cmd)
        self.ref_idx2_entry.insert(0, "65"); self.ref_idx2_entry.grid(row=3, column=3, padx=5)
        self._bind_tooltip(self.ref_idx2_entry, "0-based index of second channel for initial re-ref. Blank=skip.")
        ttk.Label(params_frame, text="Max Idx Keep:").grid(row=3, column=4, sticky="w", padx=5, pady=2)
        self.max_idx_keep_entry = ttk.Entry(params_frame, width=8, validate='key', validatecommand=validate_int_cmd)
        self.max_idx_keep_entry.insert(0, "64"); self.max_idx_keep_entry.grid(row=3, column=5, padx=5)
        self._bind_tooltip(self.max_idx_keep_entry, "Keep channels 0 to this index (exclusive). E.g., 64 keeps 0-63. Blank=keep all.")

        # Conditions Frame
        conditions_frame = ttk.LabelFrame(main_frame, text="Condition Names (Event Markers) - Case Sensitive") # Updated Label
        conditions_frame.pack(fill="both", expand=True, padx=5, pady=5)
        conditions_outer_frame = ttk.Frame(conditions_frame)
        conditions_outer_frame.pack(fill="both", expand=True)
        self.conditions_canvas = tk.Canvas(conditions_outer_frame, borderwidth=0, highlightthickness=0) # No border/highlight
        self.conditions_inner_frame = ttk.Frame(self.conditions_canvas) # Holds the entries
        conditions_scrollbar = ttk.Scrollbar(conditions_outer_frame, orient="vertical", command=self.conditions_canvas.yview)
        self.conditions_canvas.configure(yscrollcommand=conditions_scrollbar.set)
        conditions_scrollbar.pack(side="right", fill="y")
        self.conditions_canvas.pack(side="left", fill="both", expand=True)
        self.canvas_frame_id = self.conditions_canvas.create_window((0, 0), window=self.conditions_inner_frame, anchor="nw")
        self.conditions_inner_frame.bind("<Configure>", self._on_inner_frame_configure)
        self.conditions_canvas.bind("<Configure>", self._on_canvas_configure)
        self.condition_entries = []
        self.add_condition_entry() # Start with one blank entry

        # Buttons below the condition list
        condition_button_frame = ttk.Frame(conditions_frame)
        condition_button_frame.pack(fill="x", pady=5)

        # *** NEW: Detect Conditions Button ***
        self.detect_button = ttk.Button(condition_button_frame, text="Detect Conditions", command=self.detect_and_populate_conditions)
        self.detect_button.pack(side="left", padx=5)
        self._bind_tooltip(self.detect_button, "Scan first selected file for event markers and populate list below.")

        self.add_cond_button = ttk.Button(condition_button_frame, text="Add Condition Field", command=self.add_condition_entry)
        self.add_cond_button.pack(side="left", padx=5)
        self._bind_tooltip(self.add_cond_button, "Manually add another field for a condition name.")

        # Save Location Frame
        save_frame = ttk.LabelFrame(main_frame, text="Excel Output Save Location")
        save_frame.pack(fill="x", padx=5, pady=5)
        self.save_folder_path = tk.StringVar()
        btn_select_save = ttk.Button(save_frame, text="Select Parent Folder", command=self.select_save_folder)
        btn_select_save.pack(side="left", padx=5, pady=5); self._bind_tooltip(btn_select_save, "Choose parent folder for results.")
        self.save_folder_display = ttk.Entry(save_frame, textvariable=self.save_folder_path, state="readonly")
        self.save_folder_display.pack(side="left", fill="x", expand=True, padx=5, pady=5)
        self._bind_tooltip(self.save_folder_display, "Selected parent directory for saving results.")

        # Bottom Section (Log, Progress, Buttons)
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill="both", expand=False, side=tk.BOTTOM, pady=(5,0))

        # Control Buttons
        button_frame = ttk.Frame(bottom_frame)
        button_frame.pack(fill="x", padx=5, pady=5)
        self.select_button_text = tk.StringVar(); self.select_button = ttk.Button(button_frame, textvariable=self.select_button_text, command=self.select_data_source)
        self.select_button.pack(side="left", padx=10, pady=5, expand=True, fill='x')
        self.start_button = ttk.Button(button_frame, text="Start Processing", command=self.start_processing)
        self.start_button.pack(side="right", padx=10, pady=5, expand=True, fill='x'); self._bind_tooltip(self.start_button, "Begin analysis with current settings.")
        self.update_select_button_text()

        # Progress Bar
        self.progress_bar = ttk.Progressbar(bottom_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", padx=5, pady=(0,5)); self._bind_tooltip(self.progress_bar, "File processing progress.")

        # Logging Box
        log_frame = ttk.LabelFrame(bottom_frame, text="Log"); log_frame.pack(fill="both", expand=True, padx=5, pady=5)
        log_text_frame = ttk.Frame(log_frame); log_text_frame.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_text_frame, height=10, wrap="word", state="disabled", relief="sunken", borderwidth=1)
        log_scroll = ttk.Scrollbar(log_text_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.config(yscrollcommand=log_scroll.set); log_scroll.pack(side="right", fill="y"); self.log_text.pack(side="left", fill="both", expand=True)

    # --- GUI Update/Action Methods ---

    def _on_inner_frame_configure(self, event=None):
        """Updates canvas scroll region for conditions."""
        self.conditions_canvas.configure(scrollregion=self.conditions_canvas.bbox("all"))

    def _on_canvas_configure(self, event=None):
        """Updates the inner frame width for conditions."""
        self.conditions_canvas.itemconfig(self.canvas_frame_id, width=event.width)

    def add_condition_entry(self, event=None):
        """Adds a new condition entry field, binding Enter key."""
        frame = ttk.Frame(self.conditions_inner_frame)
        frame.pack(fill="x", pady=1, padx=2) # Reduced padding
        entry = ttk.Entry(frame, width=40)
        entry.pack(side="left", fill="x", expand=True)
        self._bind_tooltip(entry, "Enter condition name EXACTLY as in data (Case Sensitive). Press Enter to add another.")
        entry.bind("<Return>", self.add_condition_entry)
        remove_btn = ttk.Button(frame, text="X", width=3, style='Toolbutton', command=lambda f=frame, e=entry: self.remove_condition_entry(f, e)) # Smaller button?
        remove_btn.pack(side="right", padx=(2,0))
        self._bind_tooltip(remove_btn, "Remove this condition field.")
        self.condition_entries.append(entry)
        if event is None: entry.focus_set() # Only focus if added manually, not during populate
        self.conditions_inner_frame.update_idletasks()
        self._on_inner_frame_configure()
        # Scroll to bottom if needed
        self.after(10, lambda: self.conditions_canvas.yview_moveto(1.0))


    def remove_condition_entry(self, frame, entry):
        """Removes a condition entry field."""
        if len(self.condition_entries) > 0: # Allow removing the last one now
             try:
                 frame.destroy()
                 if entry in self.condition_entries: # Check before removing
                      self.condition_entries.remove(entry)
                 self.conditions_inner_frame.update_idletasks()
                 self._on_inner_frame_configure()
                 # If list becomes empty, maybe add one back? Or allow empty state. Let's allow empty.
             except Exception as e: self.log(f"Error removing condition frame: {e}")
        else: pass # No entries to remove

    def select_save_folder(self):
        """Opens dialog to select parent folder for Excel results."""
        folder = filedialog.askdirectory(title="Select Parent Folder for Excel Output")
        if folder: self.save_folder_path.set(folder); self.log(f"Output target folder: {folder}")
        else: self.log("Save folder selection cancelled.")

    def update_select_button_text(self):
        """Updates text and tooltip of the data selection button."""
        mode = self.file_mode.get()
        text = "Select EEG File" if mode == "Single" else "Select Data Folder"
        tooltip = "Click to select a single EEG file." if mode == "Single" else "Click to select folder with EEG files."
        self.select_button_text.set(text); self._bind_tooltip(self.select_button, tooltip)

    def select_data_source(self):
        """Prompts user for file/folder, updates data_paths and progress bar max."""
        # (Code Unchanged)
        self.data_paths = []
        file_ext = "*" + self.file_type.get().lower(); file_type_desc = self.file_type.get()
        try:
            if self.file_mode.get() == "Single":
                ftypes = [(f"{file_type_desc} files", file_ext)]
                if file_type_desc == ".BDF": ftypes.append((".set files", "*.set"))
                elif file_type_desc == ".set": ftypes.append((".BDF files", "*.bdf"))
                ftypes.append(("All files", "*.*"))
                file_path = filedialog.askopenfilename(title="Select EEG File", filetypes=ftypes)
                if file_path:
                    selected_ext = os.path.splitext(file_path)[1].lower()
                    if selected_ext == '.bdf': self.file_type.set(".BDF")
                    elif selected_ext == '.set': self.file_type.set(".set")
                    self.data_paths = [file_path]; self.log(f"Selected file: {os.path.basename(file_path)}")
                else: self.log("No file selected.")
            else: # Batch mode
                folder = filedialog.askdirectory(title=f"Select Folder Containing {file_type_desc} Files")
                if folder:
                    search_path = os.path.join(folder, file_ext); found_files = glob.glob(search_path)
                    if found_files:
                        self.data_paths = sorted(found_files); self.log(f"Selected folder: {folder}")
                        self.log(f"Found {len(found_files)} file(s) matching '{file_ext}'.")
                    else: self.log(f"No files matching '{file_ext}' found in {folder}."); messagebox.showwarning("No Files Found", f"No files with '{file_ext}' found in:\n{folder}")
                else: self.log("No folder selected.")
        except Exception as e: self.log(f"Error during selection: {e}"); messagebox.showerror("Selection Error", f"An error occurred: {e}")
        self.progress_bar['maximum'] = len(self.data_paths) if self.data_paths else 1
        self.progress_bar['value'] = 0

    def log(self, message):
        """Appends a timestamped message to the logging text box safely."""
        if hasattr(self, 'log_text') and self.log_text:
            try:
                # Simple check if called from main thread (crude but often works)
                # A more robust way involves checking threading.current_thread() == threading.main_thread()
                # but let's rely on the queue mechanism for background logging.
                # If called directly from main thread (e.g., GUI actions, post-processing):
                self.log_text.config(state="normal")
                self.log_text.insert(tk.END, f"{pd.Timestamp.now().strftime('%H:%M:%S')}: {message}\n")
                self.log_text.see(tk.END)
                self.log_text.config(state="disabled")
                self.update_idletasks() # Force update if called from main thread actions
            except tk.TclError as e:
                if "invalid command name" not in str(e): print(f"GUI Log Error: {e}")
            except Exception as e: print(f"Unexpected GUI Log Error: {e}")
        else:
             # Fallback for early calls or calls from non-main thread without queue
             print(f"{pd.Timestamp.now().strftime('%H:%M:%S')} Log: {message}")

    def detect_and_populate_conditions(self):
        """
        Scans the first selected file for event markers and populates the
        condition list in the GUI. Runs on the main thread.
        """
        self.log("Attempting to detect conditions...")
        if not self.data_paths:
            messagebox.showerror("No Data Selected", "Please select a file or folder first.")
            self.log("Detection failed: No data selected.")
            return

        representative_file = self.data_paths[0]
        self.log(f"Scanning file for conditions: {os.path.basename(representative_file)}")

        raw = None
        event_id_dict = None
        try:
            # Temporarily disable buttons during detection
            self.detect_button.config(state="disabled")
            self.add_cond_button.config(state="disabled")
            self.update_idletasks()

            raw = self.load_eeg_file(representative_file)
            if raw is None:
                messagebox.showerror("Loading Error", f"Could not load file to detect conditions:\n{os.path.basename(representative_file)}")
                self.log("Condition detection failed: File loading error.")
                return

            events, event_id_dict = mne.events_from_annotations(raw, verbose=False)
            if not event_id_dict: # Check if dictionary is empty
                 messagebox.showinfo("No Events Found", f"No event markers / annotations found in:\n{os.path.basename(representative_file)}")
                 self.log("Condition detection: No events found in the file.")
                 return # Keep existing fields if none found? Or clear? Let's clear.

            # Clear existing entries robustly
            for widget in self.conditions_inner_frame.winfo_children():
                widget.destroy()
            self.condition_entries = [] # Reset the list

            # Populate with detected conditions (original case)
            detected_conditions = sorted(event_id_dict.keys())
            self.log(f"Detected {len(detected_conditions)} unique conditions: {detected_conditions}")
            for event_name in detected_conditions:
                 # Ensure event_name is a string before inserting
                 event_name_str = str(event_name)
                 self.add_condition_entry() # Add empty field first
                 new_entry = self.condition_entries[-1]
                 new_entry.delete(0, tk.END) # Clear potential default text if any
                 new_entry.insert(0, event_name_str) # Insert detected name

            messagebox.showinfo("Conditions Detected",
                                f"Found {len(detected_conditions)} condition(s) in the first file.\n"
                                "Please review the list and remove any unwanted entries before processing.")

        except Exception as e:
            self.log(f"Error during condition detection: {e}")
            self.log(traceback.format_exc())
            messagebox.showerror("Detection Error", f"An error occurred while detecting conditions:\n{e}")
        finally:
            # Re-enable buttons
            self.detect_button.config(state="normal")
            self.add_cond_button.config(state="normal")
            # Cleanup loaded raw object
            if raw is not None:
                del raw
            # Ensure focus is reasonable after operation
            if self.condition_entries:
                 self.condition_entries[0].focus_set()
            self.update_idletasks()


    # --- Core Processing Control ---

    def start_processing(self):
        """
        Validates inputs, retrieves parameters, and starts the background
        processing thread. Condition matching is Case Sensitive.
        """
        if self.processing_thread and self.processing_thread.is_alive():
            messagebox.showwarning("Busy", "Processing is already in progress."); return

        self.log("="*40); self.log("Processing initiated...")

        # 1. Validation (Files, Save Folder)
        if not self.data_paths: messagebox.showerror("Input Error", "No data file(s) selected."); return
        if not self.save_folder_path.get(): messagebox.showerror("Input Error", "No save folder selected."); return

        # 2. Retrieve & Validate Parameters
        try: # Simplified retrieval into params dict
            params = { key: (getter().get() if getter().get() else default)
                       for key, getter, default in [
                           ('low_pass', lambda: self.low_pass_entry, 0.0),
                           ('epoch_start', lambda: self.epoch_start_entry, -1.0),
                           ('epoch_end', lambda: self.epoch_end_entry, 5.0),
                           ('reject_thresh', lambda: self.reject_thresh_entry, 5.0)]}
            params.update({ key: (float(getter().get()) if getter().get() else default)
                            for key, getter, default in [
                            ('high_pass', lambda: self.high_pass_entry, None),
                            ('downsample_rate', lambda: self.downsample_entry, None)]})
            params.update({ key: (int(getter().get()) if getter().get() else default)
                           for key, getter, default in [
                            ('ref_idx1', lambda: self.ref_idx1_entry, None),
                            ('ref_idx2', lambda: self.ref_idx2_entry, None),
                            ('max_idx_keep', lambda: self.max_idx_keep_entry, None)]})
            params['save_preprocessed'] = self.save_preprocessed.get()

            # Logical validation (moved inside try block)
            if params['high_pass'] is not None and params['low_pass'] >= params['high_pass']: raise ValueError("Low Pass >= High Pass")
            if params['downsample_rate'] is not None and params['downsample_rate'] <= 0: raise ValueError("Downsample rate <= 0")
            if params['epoch_start'] >= params['epoch_end']: raise ValueError("Epoch Start >= Epoch End")
            if params['reject_thresh'] < 0: raise ValueError("Reject Z-Thresh < 0")
            if any(p is not None and p < 0 for p in [params['ref_idx1'], params['ref_idx2']]): raise ValueError("Ref indices < 0")
            if params['max_idx_keep'] is not None and params['max_idx_keep'] <= 0: raise ValueError("Max Idx Keep <= 0")
            if params['ref_idx1'] is not None and params['ref_idx2'] is not None and params['ref_idx1'] == params['ref_idx2']: raise ValueError("Ref indices same")
            if (params['ref_idx1'] is None) != (params['ref_idx2'] is None): raise ValueError("Specify both ref indices or neither")
        except ValueError as e: messagebox.showerror("Parameter Error", f"Invalid value: {e}"); return
        except Exception as e: messagebox.showerror("Parameter Error", f"Error retrieving: {e}"); return

        # Gather condition names (Case Sensitive)
        self.condition_names_gui = [e.get().strip() for e in self.condition_entries if e.get().strip()]
        if not self.condition_names_gui: messagebox.showerror("Condition Error", "Enter or Detect condition names."); return
        self.current_conditions_process = list(self.condition_names_gui) # Keep original case
        self.log(f"Processing conditions (Case Sensitive): {self.current_conditions_process}")

        # 3. Prepare for Thread
        self.preprocessed_data = {} # Reset results dict
        self.progress_bar['value'] = 0; self.progress_bar['maximum'] = len(self.data_paths) # Ensure max is set
        self.start_button.config(state="disabled"); self.select_button.config(state="disabled")
        self.detect_button.config(state="disabled"); self.add_cond_button.config(state="disabled") # Disable condition buttons too
        self.log("Starting background processing...")
        thread_args = (list(self.data_paths), params, list(self.current_conditions_process), self.gui_queue)

        # 4. Start Thread
        self.processing_thread = threading.Thread(target=self._processing_thread_func, args=thread_args, daemon=True)
        self.processing_thread.start()

        # 5. Start Periodic Queue Check
        self.after(100, self._periodic_queue_check) # Check queue every 100ms


    def _periodic_queue_check(self):
        """ Checks GUI queue for messages from background thread. """
        while True:
            try:
                message = self.gui_queue.get_nowait()
                msg_type = message.get('type')

                if msg_type == 'log': self.log(message.get('message', ''))
                elif msg_type == 'progress': self.progress_bar['value'] = message.get('value', 0)
                elif msg_type == 'result': self.preprocessed_data = message.get('data', {}); self.log("Preprocessing results received.")
                elif msg_type == 'error':
                    error_msg = message.get('message', 'Unknown thread error.'); tb_info = message.get('traceback', '')
                    self.log(f"!!! THREAD ERROR: {error_msg}");
                    if tb_info: self.log(tb_info)
                    messagebox.showerror("Processing Error", error_msg)
                    self._finalize_processing(success=False) # Call finalize
                    return # Stop loop
                elif msg_type == 'done':
                    self.log("Background thread signaled completion.")
                    self._finalize_processing(success=True) # Call finalize
                    return # Stop loop
            except queue.Empty: break # No more messages
            except Exception as e: self.log(f"Queue check error: {e}"); break # Safety break

        # Reschedule only if thread is alive and processing seems ongoing
        if self.processing_thread and self.processing_thread.is_alive():
            self.after(100, self._periodic_queue_check)
        elif not self.start_button['state'] == 'normal': # Check if finalize wasn't called
             self.log("Thread ended; re-enabling controls (fallback).")
             self._enable_controls() # Fallback re-enable

    def _finalize_processing(self, success):
        """Handles tasks after thread finishes (successfully or via error)."""
        if success and self.preprocessed_data:
            self.log("\n--- Starting Post-processing Phase (Main Thread) ---")
            try:
                self.post_process(self.current_conditions_process) # Use original case list
                self.log("--- Post-processing Phase Complete ---")
                messagebox.showinfo("Processing Complete", "Analysis finished successfully.")
            except Exception as post_err:
                self.log(f"!!! Post-processing Error: {post_err}"); self.log(traceback.format_exc())
                messagebox.showerror("Post-processing Error", f"Error during post-processing:\n{post_err}")
        elif success: # Preprocessing succeeded but yielded no data
            self.log("--- Skipping Post-processing: No preprocessed data generated ---")
            messagebox.showwarning("Processing Finished", "Preprocessing finished, but no usable epochs were generated for the specified conditions.")

        self._enable_controls()
        self.log(f"--- Processing Run Finished at {pd.Timestamp.now()} ---")

    def _enable_controls(self):
        """Re-enables buttons after processing finishes or stops."""
        self.start_button.config(state="normal")
        self.select_button.config(state="normal")
        self.detect_button.config(state="normal")
        self.add_cond_button.config(state="normal")
        self.update_idletasks()


    # --- Background Thread Function ---

    def _processing_thread_func(self, data_paths, params, conditions_to_process, gui_queue):
        """ Runs heavy processing in background thread. """
        local_preprocessed_data = {cond: [] for cond in conditions_to_process}
        files_with_epochs = 0
        processing_error_occurred = False # Flag to track errors

        try:
            num_files = len(data_paths)
            for i, file_path in enumerate(data_paths):
                base_filename = os.path.basename(file_path)
                gui_queue.put({'type': 'log', 'message': f"\nProcessing file {i+1}/{num_files}: {base_filename}"})
                events, event_id_dict, raw_processed = None, None, None # Init per file

                try: # Inner try for file-specific errors to allow continuing batch
                    raw = self.load_eeg_file(file_path)
                    if raw is None: raise ValueError("File loading failed.")

                    # --- Event Extraction ---
                    try:
                        events, event_id_dict = mne.events_from_annotations(raw, verbose=False)
                        if not events.size: gui_queue.put({'type': 'log', 'message': f"Warning: No events found."})
                        else: gui_queue.put({'type': 'log', 'message': f"Extracted {len(events)} events. IDs: {event_id_dict}"})
                    except Exception as event_err: gui_queue.put({'type': 'log', 'message': f"Warning: Event extraction error: {event_err}."})

                    # --- Preprocessing ---
                    raw_processed = self.preprocess_raw(raw.copy(), **params) # Pass params dict
                    del raw # Free memory
                    if raw_processed is None: raise ValueError("Preprocessing failed critically.")

                    # --- Epoch Extraction (Case Sensitive) ---
                    file_had_epochs = False
                    if events is not None and event_id_dict is not None and events.size > 0:
                        gui_queue.put({'type': 'log', 'message': "Attempting case-sensitive epoch extraction..."})
                        for cond_process in conditions_to_process:
                            if cond_process in event_id_dict: # Case-sensitive check
                                target_event_id = {cond_process: event_id_dict[cond_process]}
                                try:
                                    epochs = mne.Epochs(raw_processed, events, event_id=target_event_id,
                                                        tmin=params['epoch_start'], tmax=params['epoch_end'],
                                                        preload=False, verbose=False, baseline=None)
                                    if len(epochs) > 0:
                                        gui_queue.put({'type': 'log', 'message': f"Found {len(epochs)} epochs for '{cond_process}'."})
                                        local_preprocessed_data.setdefault(cond_process, []).append(epochs)
                                        file_had_epochs = True
                                except Exception as epoch_err: gui_queue.put({'type': 'log', 'message': f"Epoch creation error for '{cond_process}': {epoch_err}\n{traceback.format_exc()}"})
                        if file_had_epochs: files_with_epochs += 1
                    else: gui_queue.put({'type': 'log', 'message': f"Skipping epoch extraction (no initial events)."})

                    # --- Save Preprocessed (Optional) ---
                    if params['save_preprocessed'] and raw_processed is not None:
                        save_filename = f"{os.path.splitext(base_filename)[0]}_preproc.fif"
                        save_path = os.path.join(os.path.dirname(file_path), save_filename)
                        try: raw_processed.save(save_path, overwrite=True, verbose=False); gui_queue.put({'type': 'log', 'message': f"Saved: {save_path}"})
                        except Exception as save_err: gui_queue.put({'type': 'log', 'message': f"Save error {save_path}: {save_err}"})

                except Exception as file_err:
                     # Log error for this specific file but continue batch
                     gui_queue.put({'type': 'log', 'message': f"!!! ERROR processing file {base_filename}: {file_err}"})
                     gui_queue.put({'type': 'log', 'message': traceback.format_exc()})
                     processing_error_occurred = True # Mark that at least one file failed

                finally:
                    # Cleanup per file and update progress
                    if 'raw_processed' in locals() and raw_processed: del raw_processed
                    gui_queue.put({'type': 'progress', 'value': i + 1})

            # --- Loop Finished ---
            gui_queue.put({'type': 'log', 'message': "\n--- Background Preprocessing Phase Complete ---"})
            gui_queue.put({'type': 'result', 'data': local_preprocessed_data})

        except MemoryError: gui_queue.put({'type': 'error', 'message': "Memory Error during background processing."})
        except Exception as e: gui_queue.put({'type': 'error', 'message': f"Critical error in thread: {e}", 'traceback': traceback.format_exc()})
        finally: gui_queue.put({'type': 'done'}) # Signal completion (queue checker handles state)


    # --- EEG Loading/Processing Methods ---
    # (load_eeg_file, preprocess_raw, post_process) - Mostly unchanged, logging is handled

    def load_eeg_file(self, filepath):
        """ Loads EEG file, attempts montage. Logs via self.log (console fallback). """
        ext = os.path.splitext(filepath)[1].lower(); raw = None
        self.log(f"Loading: {os.path.basename(filepath)}...")
        try:
            load_kwargs = {'preload': True, 'verbose': False}
            if ext == ".bdf":
                 with mne.utils.use_log_level('ERROR'): raw = mne.io.read_raw_bdf(filepath, **load_kwargs)
            elif ext == ".set":
                 with mne.utils.use_log_level('ERROR'): raw = mne.io.read_raw_eeglab(filepath, **load_kwargs)
            else: self.log(f"Unsupported format '{ext}'."); return None
            self.log(f"Loaded {len(raw.ch_names)} channels @ {raw.info['sfreq']} Hz.")
            try:
                montage = mne.channels.make_standard_montage('standard_1020')
                raw.set_montage(montage, on_missing='warn', match_case=False)
                if raw.get_montage(): self.log("Applied standard_1020 montage.")
                else: self.log("Warning: Montage not applied.")
            except Exception as montage_err: self.log(f"Warning: Montage error: {montage_err}")
            return raw
        except MemoryError: self.log(f"Memory Error loading {os.path.basename(filepath)}."); return None
        except Exception as e: self.log(f"Load Error {os.path.basename(filepath)}: {e}\n{traceback.format_exc()}"); return None

    def preprocess_raw(self, raw, downsample_rate, low_pass, high_pass, reject_thresh, ref_idx1, ref_idx2, max_idx_keep, **kwargs): # Added **kwargs to absorb unused params
        """ Applies preprocessing steps. Logs via self.log (console fallback). """
        # (Code largely unchanged from previous threaded version)
        try:
            ch_names_orig = raw.info['ch_names']; n_chans_orig = len(ch_names_orig)
            self.log(f"Preprocessing {n_chans_orig} channels...")
            # Step 1: Re-reference
            if ref_idx1 is not None and ref_idx2 is not None:
                if 0 <= ref_idx1 < n_chans_orig and 0 <= ref_idx2 < n_chans_orig and ref_idx1 != ref_idx2:
                    ref_ch_names = [ch_names_orig[ref_idx1], ch_names_orig[ref_idx2]]
                    self.log(f"Re-referencing to: {ref_ch_names}..."); # Shortened log
                    try: raw.set_eeg_reference(ref_ch_names, projection=False); self.log("Success.")
                    except Exception as ref_err: self.log(f"Warning: Re-ref failed: {ref_err}.")
                else: self.log(f"Warning: Invalid ref indices. Skipping.")
            else: self.log("Skipping custom re-ref.")
            # Step 2: Drop channels
            if max_idx_keep is not None:
                current_ch_names = raw.info['ch_names']; current_n_chans = len(current_ch_names)
                if 0 < max_idx_keep <= current_n_chans:
                    indices_to_drop = list(range(max_idx_keep, current_n_chans))
                    if indices_to_drop:
                        channels_to_drop = [current_ch_names[i] for i in indices_to_drop]
                        self.log(f"Dropping {len(channels_to_drop)} channels (>= idx {max_idx_keep})...");
                        try: raw.drop_channels(channels_to_drop); self.log(f"Remaining: {len(raw.ch_names)}")
                        except Exception as drop_err: self.log(f"Warning: Drop failed: {drop_err}")
                elif max_idx_keep > current_n_chans: self.log(f"Info: Max Idx Keep >= channels. None dropped.")
                else: self.log(f"Warning: Invalid Max Idx Keep. None dropped.")
            else: self.log("Skipping index-based channel drop.")
            # Step 3: Downsample
            if downsample_rate is not None:
                current_sfreq = raw.info['sfreq']
                if current_sfreq > downsample_rate:
                    self.log(f"Downsampling {current_sfreq:.1f}Hz -> {downsample_rate}Hz...")
                    try: raw.resample(downsample_rate, npad="auto", verbose=False); self.log(f"New rate: {raw.info['sfreq']:.1f} Hz.")
                    except Exception as ds_err: self.log(f"Error downsampling: {ds_err}"); return None
                else: self.log(f"No downsampling needed.")
            else: self.log("Skipping downsampling.")
            # Step 4: Filter
            l_freq = low_pass if low_pass > 0 else None; h_freq = high_pass
            if l_freq is not None or h_freq is not None:
                self.log(f"Filtering ({l_freq} Hz - {h_freq} Hz)...")
                try:
                    nyquist = raw.info['sfreq'] / 2.0
                    if h_freq is not None and h_freq >= nyquist: h_freq = nyquist - 0.5; self.log(f"Adjusted high pass to {h_freq:.1f} Hz.")
                    if l_freq is not None and h_freq is not None and l_freq >= h_freq: self.log(f"Warning: Invalid filter range. Skipping.")
                    else: raw.filter(l_freq=l_freq, h_freq=h_freq, method='fir', phase='zero-double', verbose=False); self.log("Filtering complete.") # fir_design='firwin' default
                except Exception as filter_err: self.log(f"Warning: Filtering error: {filter_err}.")
            else: self.log("Skipping filtering.")
            # Step 5: Kurtosis Rejection & Interpolation
            self.log(f"Kurtosis rejection (Z > {reject_thresh})...")
            try: # Simplified Kurtosis block
                picks_eeg = mne.pick_types(raw.info, eeg=True, exclude='bads')
                if len(picks_eeg) > 1:
                    data = raw.get_data(picks=picks_eeg); channel_kurt = kurtosis(data, axis=1, fisher=True); del data
                    channel_kurt = np.nan_to_num(channel_kurt); mean_kurt = np.mean(channel_kurt); std_kurt = np.std(channel_kurt)
                    z_scores = np.zeros_like(channel_kurt) if std_kurt < 1e-6 else (channel_kurt - mean_kurt) / std_kurt
                    eeg_ch_names = [raw.info['ch_names'][i] for i in picks_eeg]
                    bad_channels_kurt = [eeg_ch_names[i] for i, z in enumerate(z_scores) if abs(z) > reject_thresh]
                    if bad_channels_kurt:
                        self.log(f"Found bad via Kurtosis: {bad_channels_kurt}.")
                        raw.info['bads'].extend([ch for ch in bad_channels_kurt if ch not in raw.info['bads']])
                        if raw.info['bads'] and raw.get_montage():
                             self.log(f"Interpolating bads: {raw.info['bads']}...")
                             try: raw.interpolate_bads(reset_bads=True, mode='accurate', verbose=False); self.log("Interpolation complete.")
                             except Exception as interp_err: self.log(f"Warning: Interpolation error: {interp_err}.")
                        elif raw.info['bads']: self.log("Warning: Cannot interpolate (no montage).")
                    else: self.log("No channels rejected via Kurtosis.")
                else: self.log("Warning: Not enough EEG channels for Kurtosis.")
            except Exception as kurt_err: self.log(f"Warning: Kurtosis error: {kurt_err}.\n{traceback.format_exc()}")
            # Step 6: Average Reference
            self.log("Applying average reference...");
            try: raw.set_eeg_reference(ref_channels='average', projection=False); self.log("Avg ref complete.")
            except Exception as avg_ref_err: self.log(f"Warning: Avg ref failed: {avg_ref_err}")
            self.log("Preprocessing steps finished.")
            return raw
        except MemoryError: self.log("Memory Error during preprocessing."); return None
        except Exception as e: self.log(f"Critical preprocessing error: {e}\n{traceback.format_exc()}"); return None

    def post_process(self, conditions_to_process):
        """ Calculates FFT metrics and saves results. Runs in Main Thread. """
        # (Code Unchanged from previous version - uses self.log, processes self.preprocessed_data)
        if not self.save_folder_path.get(): self.log("Error: Save folder missing."); return
        parent_folder = self.save_folder_path.get()

        for cond_name in conditions_to_process: # Now case-sensitive
            epochs_object_list = self.preprocessed_data.get(cond_name, [])
            if not epochs_object_list: self.log(f"No data for '{cond_name}'. Skipping."); continue

            n_files_for_cond = len(epochs_object_list)
            self.log(f"\nPost-processing '{cond_name}' ({n_files_for_cond} file(s))...")

            accum_fft_amp = accum_snr = accum_z_score = accum_bca = None
            valid_file_count = 0; final_n_channels = None

            for file_idx, epochs in enumerate(epochs_object_list):
                self.log(f"  Processing results from file {file_idx+1}/{n_files_for_cond}...")
                try:
                    if len(epochs) == 0: self.log("    0 epochs. Skipping."); continue
                    self.log(f"    Loading {len(epochs)} epochs..."); epochs.load_data()
                    n_epochs, n_channels, n_times = epochs.get_data(copy=False).shape
                    if final_n_channels is None: final_n_channels = n_channels
                    elif final_n_channels != n_channels: self.log("    Warning: Ch count mismatch. Skipping."); continue
                    sfreq = epochs.info['sfreq']
                    self.log(f"    Processing {n_epochs} epochs, {n_channels} chans, {n_times} points @ {sfreq}Hz.")

                    # --- Use Welch PSD ---
                    try: # Get high pass from entry for fmax, provide default
                         fmax_psd = float(self.high_pass_entry.get()) if self.high_pass_entry.get() else sfreq/2.0
                    except: fmax_psd = sfreq/2.0 # Default on error
                    fmax_psd = min(fmax_psd, sfreq/2.0) # Ensure <= Nyquist

                    spectrum = epochs.compute_psd(method='welch', fmin=0.5, fmax=fmax_psd,
                                                n_fft=int(sfreq * 2), n_overlap=int(sfreq * 1), # 2s window, 1s overlap
                                                window='hann', average='mean', verbose=False)
                    psd_freqs = spectrum.freqs; avg_power = spectrum.get_data(return_freqs=False)
                    file_avg_fft_amp = np.sqrt(avg_power); freqs_for_metrics = psd_freqs; n_freqs_metrics = len(freqs_for_metrics)

                    # --- Calculate Metrics ---
                    n_target = len(TARGET_FREQUENCIES)
                    file_fft_out=np.zeros((n_channels, n_target)); file_snr_out=np.zeros((n_channels, n_target))
                    file_z_out=np.zeros((n_channels, n_target)); file_bca_out=np.zeros((n_channels, n_target))
                    noise_range, noise_exclude = 12, 1 # Bins for noise estimation
                    for ch in range(n_channels):
                        for idx_freq, target_freq in enumerate(TARGET_FREQUENCIES):
                            target_bin = np.argmin(np.abs(freqs_for_metrics - target_freq))
                            lower_b = max(0, target_bin - noise_range); excl_s = max(0, target_bin - noise_exclude)
                            excl_e = min(n_freqs_metrics, target_bin + noise_exclude + 1); upper_b = min(n_freqs_metrics, target_bin + noise_range + 1)
                            neigh_idx = np.unique(np.concatenate([np.arange(lower_b, excl_s), np.arange(excl_e, upper_b)]))
                            neigh_idx = neigh_idx[neigh_idx < n_freqs_metrics]
                            if neigh_idx.size < 4: continue
                            neigh_amps = file_avg_fft_amp[ch, neigh_idx]; noise_mean = np.mean(neigh_amps); noise_std = np.std(neigh_amps)
                            fft_val = file_avg_fft_amp[ch, target_bin]; snr_val = fft_val / noise_mean if noise_mean > 1e-12 else 0
                            loc_max_r = np.arange(max(0, target_bin-1), min(n_freqs_metrics, target_bin+2));
                            if loc_max_r.size == 0: continue
                            loc_max = np.max(file_avg_fft_amp[ch, loc_max_r]); z_val = (loc_max - noise_mean) / noise_std if noise_std > 1e-12 else 0
                            bca_val = fft_val - noise_mean
                            file_fft_out[ch, idx_freq]=fft_val; file_snr_out[ch, idx_freq]=snr_val; file_z_out[ch, idx_freq]=z_val; file_bca_out[ch, idx_freq]=bca_val

                    # --- Accumulate ---
                    if accum_fft_amp is None: accum_fft_amp=file_fft_out; accum_snr=file_snr_out; accum_z_score=file_z_out; accum_bca=file_bca_out
                    else: accum_fft_amp+=file_fft_out; accum_snr+=file_snr_out; accum_z_score+=file_z_out; accum_bca+=file_bca_out
                    valid_file_count += 1; self.log(f"    Accumulated metrics from file {file_idx+1}.")
                except MemoryError: self.log(f"!!! Memory Error post-proc file {file_idx+1}. Skipping."); messagebox.showwarning("Memory Error", f"Ran out of memory post-processing {cond_name}, file {file_idx+1}.")
                except Exception as e: self.log(f"Error post-proc file {file_idx+1}: {e}\n{traceback.format_exc()}")
                finally:
                    if 'epochs' in locals() and epochs: epochs.drop_log_stats(); del epochs

            # --- Final Average & Excel ---
            if valid_file_count == 0: self.log(f"No valid data for '{cond_name}'. No Excel."); continue
            self.log(f"Averaging metrics across {valid_file_count} file(s) for '{cond_name}'.")
            avgFFT=accum_fft_amp/valid_file_count; avgSNR=accum_snr/valid_file_count; avgZ=accum_z_score/valid_file_count; avgBCA=accum_bca/valid_file_count
            subfolder_path = os.path.join(parent_folder, cond_name) # Case-sensitive folder name
            try: os.makedirs(subfolder_path, exist_ok=True)
            except OSError as e: self.log(f"Subfolder error {subfolder_path}: {e}. Saving to parent."); subfolder_path = parent_folder
            excel_filename = f"{cond_name}_Results.xlsx"; excel_path = os.path.join(subfolder_path, excel_filename)
            col_names = [f"{f:.1f}_Hz" for f in TARGET_FREQUENCIES]; n_ch_out = avgFFT.shape[0]
            electrode_col = ELECTRODE_NAMES[:n_ch_out]; # Pad if needed
            if len(electrode_col) < n_ch_out: electrode_col.extend([f"Ch{i+1}" for i in range(len(electrode_col), n_ch_out)])
            df_fft=pd.DataFrame(avgFFT, columns=col_names, index=electrode_col); df_snr=pd.DataFrame(avgSNR, columns=col_names, index=electrode_col)
            df_z=pd.DataFrame(avgZ, columns=col_names, index=electrode_col); df_bca=pd.DataFrame(avgBCA, columns=col_names, index=electrode_col)
            df_fft.insert(0, "Electrode", df_fft.index); df_snr.insert(0, "Electrode", df_snr.index); df_z.insert(0, "Electrode", df_z.index); df_bca.insert(0, "Electrode", df_bca.index)
            try: # Write Excel
                self.log(f"Writing formatted Excel: {excel_path}")
                with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                    df_fft.to_excel(writer, sheet_name="FFT_Amplitude", index=False); df_snr.to_excel(writer, sheet_name="SNR", index=False)
                    df_z.to_excel(writer, sheet_name="Z_Score", index=False); df_bca.to_excel(writer, sheet_name="BCA", index=False)
                    workbook = writer.book; center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
                    for sheet_name in writer.sheets: # Format columns
                        worksheet = writer.sheets[sheet_name]; df_map = {'FFT_Amplitude': df_fft, 'SNR': df_snr, 'Z_Score': df_z, 'BCA': df_bca}
                        current_df = df_map.get(sheet_name);
                        if current_df is None: continue
                        for col_idx, col_name in enumerate(current_df.columns):
                            header_width = len(str(col_name)); max_data_width = 0
                            try: max_data_width = current_df[col_name].astype(str).map(len).max(); max_data_width = 0 if pd.isna(max_data_width) else int(max_data_width)
                            except: pass # Ignore errors calculating max width
                            width = max(header_width, max_data_width) + 2; worksheet.set_column(col_idx, col_idx, width, center_format)
                self.log(f"Formatted Excel saved for '{cond_name}'.")
            except Exception as excel_err: self.log(f"Excel write error {excel_path}: {excel_err}\n{traceback.format_exc()}"); messagebox.showerror("Excel Error", f"Failed Excel save for '{cond_name}'.\n{excel_err}")

# --- Main execution block ---
if __name__ == "__main__":
    root = FPVSApp()
    root.mainloop()