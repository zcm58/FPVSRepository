#!/usr/bin/env python3
import os
import glob
import numpy as np
import mne
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from scipy.stats import kurtosis

#####################################################
# Fixed parameters for post-processing
#####################################################
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
    def __init__(self):
        super().__init__()
        self.title("EEG FPVS Analysis")
        self.geometry("900x700")

        # Data structures to hold preprocessed data
        self.preprocessed_data = {}    # { condition_name_lower: [list of Epochs] }
        self.condition_names_gui = []  # original condition names from the GUI

        # Will store the user-selected file paths (single mode) or list of files (batch mode)
        self.data_paths = []

        self.create_widgets()

    def create_widgets(self):
        """
        Build the GUI components:
          - Processing options (single/batch, file type)
          - Preprocessing parameters
          - Condition entry fields
          - Excel output location
          - Logging text box
          - Buttons for file/folder selection and "Start Processing"
        """
        # ========== Top Frame: Processing Options ==========
        options_frame = ttk.LabelFrame(self, text="Processing Options")
        options_frame.pack(fill="x", padx=10, pady=5)

        # File Processing Mode: Single or Batch
        self.file_mode = tk.StringVar(value="Single")
        ttk.Label(options_frame, text="File Processing Mode:").grid(
            row=0, column=0, sticky="w", padx=5, pady=2
        )
        self.radio_single = ttk.Radiobutton(
            options_frame, text="Single File",
            variable=self.file_mode, value="Single",
            command=self.update_select_button_text
        )
        self.radio_single.grid(row=0, column=1, padx=5)
        self.radio_batch = ttk.Radiobutton(
            options_frame, text="Batch Processing",
            variable=self.file_mode, value="Batch",
            command=self.update_select_button_text
        )
        self.radio_batch.grid(row=0, column=2, padx=5)

        # File Type: .BDF or .set
        self.file_type = tk.StringVar(value=".BDF")
        ttk.Label(options_frame, text="EEG File Type:").grid(
            row=1, column=0, sticky="w", padx=5, pady=2
        )
        ttk.Radiobutton(
            options_frame, text=".BDF", variable=self.file_type, value=".BDF"
        ).grid(row=1, column=1, padx=5)
        ttk.Radiobutton(
            options_frame, text=".set", variable=self.file_type, value=".set"
        ).grid(row=1, column=2, padx=5)

        # ========== Preprocessing Parameters ==========
        params_frame = ttk.LabelFrame(self, text="Preprocessing Parameters")
        params_frame.pack(fill="x", padx=10, pady=5)

        # FIR Filter range
        ttk.Label(params_frame, text="Low Pass (Hz):").grid(
            row=0, column=0, sticky="w", padx=5, pady=2
        )
        self.low_pass_entry = ttk.Entry(params_frame, width=10)
        self.low_pass_entry.insert(0, "0.1")
        self.low_pass_entry.grid(row=0, column=1, padx=5)

        ttk.Label(params_frame, text="High Pass (Hz):").grid(
            row=0, column=2, sticky="w", padx=5, pady=2
        )
        self.high_pass_entry = ttk.Entry(params_frame, width=10)
        self.high_pass_entry.insert(0, "100")
        self.high_pass_entry.grid(row=0, column=3, padx=5)

        # Downsampling Rate
        ttk.Label(params_frame, text="Downsampling Rate (Hz):").grid(
            row=1, column=0, sticky="w", padx=5, pady=2
        )
        self.downsample_entry = ttk.Entry(params_frame, width=10)
        self.downsample_entry.insert(0, "256")
        self.downsample_entry.grid(row=1, column=1, padx=5)

        # Epoch Time Window (start/end)
        ttk.Label(params_frame, text="Epoch Start (s):").grid(
            row=1, column=2, sticky="w", padx=5, pady=2
        )
        self.epoch_start_entry = ttk.Entry(params_frame, width=10)
        self.epoch_start_entry.insert(0, "-1")
        self.epoch_start_entry.grid(row=1, column=3, padx=5)

        ttk.Label(params_frame, text="Epoch End (s):").grid(
            row=1, column=4, sticky="w", padx=5, pady=2
        )
        self.epoch_end_entry = ttk.Entry(params_frame, width=10)
        self.epoch_end_entry.insert(0, "125")
        self.epoch_end_entry.grid(row=1, column=5, padx=5)

        # Rejection Threshold
        ttk.Label(params_frame, text="Rejection Threshold:").grid(
            row=2, column=0, sticky="w", padx=5, pady=2
        )
        self.reject_thresh_entry = ttk.Entry(params_frame, width=10)
        self.reject_thresh_entry.insert(0, "5")
        self.reject_thresh_entry.grid(row=2, column=1, padx=5)

        # Checkbox: Save preprocessed files
        self.save_preprocessed = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            params_frame,
            text="Save preprocessed files to disk",
            variable=self.save_preprocessed
        ).grid(row=2, column=2, columnspan=2, padx=5, pady=2)

        # ========== Conditions Frame ==========
        conditions_frame = ttk.LabelFrame(self, text="Condition Names (for epoch extraction & folder naming)")
        conditions_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.conditions_container = ttk.Frame(conditions_frame)
        self.conditions_container.pack(fill="both", expand=True)

        # We'll store references to each condition entry
        self.condition_entries = []
        self.add_condition_entry()  # add one default condition

        # Button to add another condition
        ttk.Button(
            conditions_frame, text="Add Another Condition",
            command=self.add_condition_entry
        ).pack(pady=5)

        # ========== Save Location for Excel Output ==========
        save_frame = ttk.LabelFrame(self, text="Excel Output Save Location")
        save_frame.pack(fill="x", padx=10, pady=5)

        self.save_folder_path = tk.StringVar()
        ttk.Button(
            save_frame, text="Select Parent Folder", command=self.select_save_folder
        ).pack(side="left", padx=5, pady=5)

        self.save_folder_label = ttk.Label(save_frame, textvariable=self.save_folder_path)
        self.save_folder_label.pack(side="left", padx=5)

        # ========== Logging Text Box ==========
        log_frame = ttk.LabelFrame(self, text="Log")
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.log_text = tk.Text(log_frame, height=15)
        self.log_text.pack(fill="both", expand=True)

        # ========== Dynamic File/Folder Selection Button ==========
        self.select_button_text = tk.StringVar()
        self.select_button = ttk.Button(
            self, textvariable=self.select_button_text,
            command=self.select_data_source
        )
        self.select_button.pack(pady=5)
        self.update_select_button_text()  # set initial text

        # ========== Start Processing Button ==========
        ttk.Button(self, text="Start Processing", command=self.start_processing).pack(pady=5)

    def add_condition_entry(self):
        """Dynamically add a new text entry for condition name."""
        entry = ttk.Entry(self.conditions_container, width=30)
        entry.pack(pady=2, padx=5, anchor="w")
        self.condition_entries.append(entry)

    def select_save_folder(self):
        """Prompt user to select parent folder for Excel output."""
        folder = filedialog.askdirectory(title="Select Parent Folder for Excel Output")
        if folder:
            self.save_folder_path.set(folder)
            messagebox.showinfo(
                "Save Location",
                f"Subfolders will be created under:\n{folder}"
            )

    def update_select_button_text(self):
        """Update the button text based on whether the user wants single or batch processing."""
        if self.file_mode.get() == "Single":
            self.select_button_text.set("Select file to process")
        else:
            self.select_button_text.set("Select folder containing FPVS data files")

    def select_data_source(self):
        """
        Prompt the user to either select a single file or a folder,
        depending on the file_mode. Store the results in self.data_paths.
        """
        self.data_paths = []  # clear previous selection
        file_ext = "*" + self.file_type.get().lower()

        if self.file_mode.get() == "Single":
            # Prompt user for one file
            file_path = filedialog.askopenfilename(
                title="Select EEG File",
                filetypes=[(self.file_type.get(), file_ext)]
            )
            if file_path:
                self.data_paths = [file_path]
                self.log(f"Selected single file: {file_path}")
            else:
                self.log("No file selected.")
        else:
            # Prompt user for a folder containing multiple files
            folder = filedialog.askdirectory(title="Select Folder Containing EEG Files")
            if folder:
                found_files = glob.glob(os.path.join(folder, file_ext))
                if found_files:
                    self.data_paths = found_files
                    self.log(f"Selected folder: {folder}")
                    self.log(f"Found {len(found_files)} file(s) matching '{file_ext}'.")
                else:
                    self.log(f"No files found in {folder} matching '{file_ext}'.")
            else:
                self.log("No folder selected.")

    def log(self, message):
        """Append a message to the logging text box."""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    def start_processing(self):
        """
        Triggered by the "Start Processing" button.
        Runs preprocessing and then post-processing
        on the user-selected file(s).
        """
        self.log("Starting processing...")

        # If user hasn't selected any file/folder, warn them
        if not self.data_paths:
            messagebox.showerror("Selection Error", "No file or folder selected. Please select a data source first.")
            return

        # Retrieve numeric parameters
        try:
            low_pass = float(self.low_pass_entry.get())
            high_pass = float(self.high_pass_entry.get())
            downsample_rate = float(self.downsample_entry.get())
            epoch_start = float(self.epoch_start_entry.get())
            epoch_end = float(self.epoch_end_entry.get())
            reject_thresh = float(self.reject_thresh_entry.get())
        except ValueError:
            messagebox.showerror("Parameter Error", "Please enter valid numerical values.")
            return

        # Gather condition names from GUI
        conditions_gui = [e.get().strip() for e in self.condition_entries if e.get().strip()]
        if not conditions_gui:
            messagebox.showerror("Condition Error", "Please enter at least one condition name.")
            return
        # Convert to lowercase for epoch extraction
        conditions_extract = [cond.lower() for cond in conditions_gui]
        self.condition_names_gui = conditions_gui  # store original (GUI) names

        self.log(f"Conditions for extraction: {conditions_extract}")

        # Clear any previous preprocessed data
        self.preprocessed_data = {cond: [] for cond in conditions_extract}

        # ==========================
        # Preprocessing Phase
        # ==========================
        for file_path in self.data_paths:
            self.log(f"Processing file: {os.path.basename(file_path)}")
            try:
                raw = self.load_eeg_file(file_path)
                raw = self.preprocess_raw(raw, downsample_rate, low_pass, high_pass, reject_thresh)

                # Extract epochs for each condition
                events, event_id = mne.events_from_annotations(raw)
                if not events.size:
                    self.log("No events found; skipping epoch extraction.")
                    continue

                for cond_gui, cond_extract in zip(conditions_gui, conditions_extract):
                    try:
                        # Attempt to create epochs for the given condition
                        target_id = event_id.get(cond_extract, None)
                        epochs = mne.Epochs(
                            raw, events, event_id={cond_extract: target_id},
                            tmin=epoch_start, tmax=epoch_end,
                            preload=True, verbose=False
                        )
                        if len(epochs) == 0:
                            self.log(
                                f"Warning: No epochs found for '{cond_gui}' (searched as '{cond_extract}')"
                            )
                        else:
                            self.log(f"Extracted {len(epochs)} epochs for '{cond_gui}'")
                            self.preprocessed_data[cond_extract].append(epochs)
                    except Exception as e:
                        self.log(f"Error extracting epochs for '{cond_gui}': {e}")

                # Optionally save the preprocessed file
                if self.save_preprocessed.get():
                    save_filename = os.path.splitext(os.path.basename(file_path))[0] + "_preproc.fif"
                    save_path = os.path.join(os.path.dirname(file_path), save_filename)
                    raw.save(save_path, overwrite=True)
                    self.log(f"Saved preprocessed file to: {save_path}")

            except Exception as e:
                self.log(f"Error processing file {os.path.basename(file_path)}: {e}")

        self.log("Preprocessing phase complete.")

        # ==========================
        # Post-processing Phase
        # ==========================
        self.post_process(conditions_extract)

        self.log("All processing complete.")

    def load_eeg_file(self, filepath):
        """Load an EEG file (BDF or SET) using MNE."""
        ext = os.path.splitext(filepath)[1].lower()
        if ext == ".bdf":
            self.log("Loading BDF file...")
            raw = mne.io.read_raw_bdf(filepath, preload=True, verbose=False)
        elif ext == ".set":
            self.log("Loading EEGLAB .set file...")
            raw = mne.io.read_raw_eeglab(filepath, preload=True, verbose=False)
        else:
            raise ValueError("Unsupported file format.")

        # Apply a standard 10-20 montage if available
        montage = mne.channels.make_standard_montage('standard_1020')
        raw.set_montage(montage, match_case=False)

        return raw

    def preprocess_raw(self, raw, downsample_rate, low_pass, high_pass, reject_thresh):
        """
        Replicate the preprocessing steps:
          1. Re-reference to channels 65 & 66 (if they exist)
          2. Remove channels beyond 65
          3. Downsample if above user-specified rate
          4. Filter
          5. Automatic channel rejection (kurtosis)
          6. Interpolate bad channels
          7. Re-reference to average
        """
        ch_names = raw.info['ch_names']

        # Step 1: Re-reference to channels 65 & 66 if available
        if len(ch_names) >= 66:
            ref_channels = [ch_names[64], ch_names[65]]  # zero-indexed
            self.log(f"Re-referencing to channels {ref_channels}")
            raw.set_eeg_reference(ref_channels, projection=False)
        else:
            self.log("Not enough channels for re-referencing (need >=66). Skipping custom reference.")

        # Step 2: Remove channels beyond 65
        if len(ch_names) > 65:
            channels_to_drop = ch_names[65:]
            self.log(f"Dropping channels beyond 65: {channels_to_drop}")
            raw.drop_channels(channels_to_drop)

        # Step 3: Downsample if needed
        if raw.info['sfreq'] > downsample_rate:
            self.log(f"Downsampling from {raw.info['sfreq']} Hz to {downsample_rate} Hz")
            raw.resample(downsample_rate, npad="auto")
        else:
            self.log(f"Sampling rate is {raw.info['sfreq']} Hz; no downsampling applied.")

        # Step 4: Apply FIR bandpass filter
        self.log(f"Filtering data from {low_pass} Hz to {high_pass} Hz...")
        raw.filter(l_freq=low_pass, h_freq=high_pass, method='fir', fir_design='firwin')

        # Step 5: Automatic channel rejection (kurtosis)
        data = raw.get_data()
        channel_kurt = kurtosis(data, axis=1, fisher=True)
        z_scores = (channel_kurt - np.mean(channel_kurt)) / np.std(channel_kurt)
        bad_channels = [
            raw.info['ch_names'][i] for i, z in enumerate(z_scores) if abs(z) > reject_thresh
        ]
        if bad_channels:
            self.log(f"Rejecting channels {bad_channels} based on kurtosis > {reject_thresh}")
            raw.info['bads'] = bad_channels
            raw.interpolate_bads(reset_bads=True)
        else:
            self.log("No channels rejected based on kurtosis.")

        # Step 6: Re-reference to average
        self.log("Re-referencing to average.")
        raw.set_eeg_reference(ref_channels='average', projection=False)

        return raw

    def post_process(self, conditions_extract):
        """
        For each condition, compute FFT-based metrics (FFT amplitude, SNR, Z-score, BCA)
        and write the averaged results to Excel.
        """
        if not self.save_folder_path.get():
            messagebox.showerror("Save Folder Error", "Please select a parent folder for Excel output.")
            return

        parent_folder = self.save_folder_path.get()

        for cond in conditions_extract:
            epochs_list = self.preprocessed_data.get(cond, [])
            if not epochs_list:
                self.log(f"No preprocessed data for condition '{cond}'. Skipping.")
                continue

            self.log(f"Post-processing condition '{cond}' with {len(epochs_list)} file(s)...")

            accumFFT = accumSNR = accumZ = accumBCA = None
            file_count = 0

            for epochs in epochs_list:
                try:
                    # Average across epochs (time dimension is axis=2)
                    data_avg = np.mean(epochs.get_data(), axis=2)  # shape: (n_channels, n_times)
                    n_channels, n_times = data_avg.shape

                    # Compute FFT
                    fft_vals = np.abs(np.fft.fft(data_avg, axis=1)) / n_times * 2
                    freqs = np.fft.fftfreq(n_times, d=1.0/epochs.info['sfreq'])

                    # Keep only positive frequencies
                    pos_mask = freqs >= 0
                    fft_vals = fft_vals[:, pos_mask]
                    freqs = freqs[pos_mask]

                    # Prepare outputs
                    n_target = len(TARGET_FREQUENCIES)
                    fft_out = np.zeros((n_channels, n_target))
                    snr_out = np.zeros((n_channels, n_target))
                    z_out   = np.zeros((n_channels, n_target))
                    bca_out = np.zeros((n_channels, n_target))

                    for ch in range(n_channels):
                        for idx_freq, target in enumerate(TARGET_FREQUENCIES):
                            target_bin = np.argmin(np.abs(freqs - target))
                            # ±12 bins range check
                            if target_bin - 12 < 0 or target_bin + 12 >= len(freqs):
                                continue

                            # Exclude 3 bins around the target (target_bin-1, target_bin, target_bin+1)
                            neighbor_left  = fft_vals[ch, target_bin-12 : target_bin-1]
                            neighbor_right = fft_vals[ch, target_bin+2 : target_bin+13]
                            neighbor_bins  = np.concatenate((neighbor_left, neighbor_right))

                            noise_mean = np.mean(neighbor_bins)
                            noise_std  = np.std(neighbor_bins)

                            # SNR
                            fft_val = fft_vals[ch, target_bin]
                            snr_val = fft_val / noise_mean if noise_mean != 0 else 0

                            # Z-score
                            local_max = np.max(fft_vals[ch, target_bin-1 : target_bin+2])
                            z_val = (local_max - noise_mean) / noise_std if noise_std != 0 else 0

                            # Baseline-corrected amplitude
                            bca_val = fft_val - noise_mean

                            fft_out[ch, idx_freq] = fft_val
                            snr_out[ch, idx_freq] = snr_val
                            z_out[ch, idx_freq]   = z_val
                            bca_out[ch, idx_freq] = bca_val

                    if accumFFT is None:
                        accumFFT = fft_out
                        accumSNR = snr_out
                        accumZ   = z_out
                        accumBCA = bca_out
                    else:
                        accumFFT += fft_out
                        accumSNR += snr_out
                        accumZ   += z_out
                        accumBCA += bca_out

                    file_count += 1
                except Exception as e:
                    self.log(f"Error in post-processing for condition '{cond}': {e}")

            if file_count == 0:
                self.log(f"No valid data for condition '{cond}' after processing.")
                continue

            # Average results across files
            avgFFT = accumFFT / file_count
            avgSNR = accumSNR / file_count
            avgZ   = accumZ / file_count
            avgBCA = accumBCA / file_count

            # Create subfolder named after the original GUI condition name
            idx_gui = conditions_extract.index(cond)
            cond_folder_name = self.condition_names_gui[idx_gui]
            subfolder_path = os.path.join(parent_folder, cond_folder_name)
            os.makedirs(subfolder_path, exist_ok=True)

            # Prepare dataframes
            col_names = ["Hz_" + str(f) for f in TARGET_FREQUENCIES]
            df_fft = pd.DataFrame(avgFFT, columns=col_names)
            df_snr = pd.DataFrame(avgSNR, columns=col_names)
            df_z   = pd.DataFrame(avgZ,   columns=col_names)
            df_bca = pd.DataFrame(avgBCA, columns=col_names)

            # Insert electrode labels (up to the # of channels)
            df_fft.insert(0, "Electrode", ELECTRODE_NAMES[:avgFFT.shape[0]])
            df_snr.insert(0, "Electrode", ELECTRODE_NAMES[:avgSNR.shape[0]])
            df_z.insert(0, "Electrode", ELECTRODE_NAMES[:avgZ.shape[0]])
            df_bca.insert(0, "Electrode", ELECTRODE_NAMES[:avgBCA.shape[0]])

            # Create Excel file path
            excel_filename = f"{cond_folder_name}_Results.xlsx"
            excel_path = os.path.join(subfolder_path, excel_filename)

            # Write to Excel
            with pd.ExcelWriter(excel_path) as writer:
                df_fft.to_excel(writer, sheet_name="FFT",    index=False)
                df_snr.to_excel(writer, sheet_name="SNR",    index=False)
                df_z.to_excel(writer, sheet_name="Z-score", index=False)
                df_bca.to_excel(writer, sheet_name="BCA",    index=False)

            self.log(f"Excel results saved for condition '{cond_folder_name}' at: {excel_path}")


if __name__ == "__main__":
    app = FPVSApp()
    app.mainloop()
