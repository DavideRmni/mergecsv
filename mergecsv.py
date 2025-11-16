import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import pandas as pd
import re
from pathlib import Path
import numpy as np
import threading

class CSVConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV Spectra Merge and Combine")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Variables
        self.selected_directory = tk.StringVar()
        self.csv_files = []
        self.file_vars = {}  # Dictionary to store checkbox variables
        
        # Auto-detect system locale and set default format
        self.detect_system_locale()
        
        self.setup_ui()
        
    def detect_system_locale(self):
        """Auto-detect system locale and set appropriate CSV format defaults"""
        import locale
        import platform
        
        try:
            # Get system locale
            system_locale = locale.getdefaultlocale()[0]
            system_language = locale.getdefaultlocale()[0][:2] if system_locale else 'en'
            
            # Initialize default values
            default_separator = "comma"
            default_decimal = "dot"
            detected_region = "International"
            
            # European countries typically use semicolon separator and comma decimal
            european_locales = ['de', 'fr', 'it', 'es', 'pt', 'nl', 'pl', 'cs', 'sk', 'hu', 'ro', 'bg', 'hr', 'sl', 'lt', 'lv', 'et', 'fi', 'sv', 'da', 'no']
            
            # North American countries use comma separator and dot decimal
            na_locales = ['en_US', 'en_CA', 'fr_CA']
            
            if system_locale in na_locales or (system_language == 'en' and platform.system() in ['Windows', 'Darwin']):
                # US/Canada format
                default_separator = "comma"
                default_decimal = "dot"
                detected_region = "North America"
            elif system_language in european_locales:
                # European format
                default_separator = "semicolon"
                default_decimal = "comma"
                detected_region = "Europe"
            elif system_language == 'en':
                # UK and other English-speaking countries
                default_separator = "comma"
                default_decimal = "dot"
                detected_region = "UK/International"
            else:
                # Default to European format for others
                default_separator = "semicolon"
                default_decimal = "comma"
                detected_region = "International (European format)"
            
            # Store the detected defaults
            self.default_separator = default_separator
            self.default_decimal = default_decimal
            self.detected_region = detected_region
            
        except Exception as e:
            # Fallback to European format if detection fails
            self.default_separator = "semicolon"
            self.default_decimal = "comma"
            self.detected_region = "Default (European format)"
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        main_frame.rowconfigure(5, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="CSV Spectra File Converter", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Directory selection
        dir_frame = ttk.LabelFrame(main_frame, text="1. Select Directory", padding="10")
        dir_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        dir_frame.columnconfigure(1, weight=1)
        
        ttk.Button(dir_frame, text="Browse...", 
                  command=self.select_directory).grid(row=0, column=0, padx=(0, 10))
        
        self.dir_label = ttk.Label(dir_frame, textvariable=self.selected_directory,
                                  relief="sunken", anchor="w")
        self.dir_label.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        # File selection
        files_frame = ttk.LabelFrame(main_frame, text="2. Select Files to Convert", padding="10")
        files_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        files_frame.columnconfigure(0, weight=1)
        files_frame.rowconfigure(1, weight=1)
        
        # Control buttons for file selection
        control_frame = ttk.Frame(files_frame)
        control_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(control_frame, text="Select All", 
                  command=self.select_all_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(control_frame, text="Deselect All", 
                  command=self.deselect_all_files).pack(side=tk.LEFT, padx=(0, 5))
        
        self.files_count_label = ttk.Label(control_frame, text="No files found")
        self.files_count_label.pack(side=tk.RIGHT)
        
        # Scrollable file list
        list_frame = ttk.Frame(files_frame)
        list_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # Create canvas and scrollbar for file list
        canvas = tk.Canvas(list_frame, height=200)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Bind mousewheel to canvas
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind("<MouseWheel>", _on_mousewheel)
        
        # CSV format selection
        format_frame = ttk.LabelFrame(main_frame, text="3. CSV Format Options", padding="10")
        format_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Field separator
        ttk.Label(format_frame, text="Field Separator:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.separator_var = tk.StringVar(value=self.default_separator)
        sep_frame = ttk.Frame(format_frame)
        sep_frame.grid(row=0, column=1, sticky=tk.W, padx=(0, 20))
        
        ttk.Radiobutton(sep_frame, text="Comma (,)", variable=self.separator_var, 
                       value="comma").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(sep_frame, text="Semicolon (;)", variable=self.separator_var, 
                       value="semicolon").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(sep_frame, text="Tab", variable=self.separator_var, 
                       value="tab").pack(side=tk.LEFT)
        
        # Decimal separator
        ttk.Label(format_frame, text="Decimal Separator:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        self.decimal_var = tk.StringVar(value=self.default_decimal)
        dec_frame = ttk.Frame(format_frame)
        dec_frame.grid(row=1, column=1, sticky=tk.W, padx=(0, 20), pady=(10, 0))
        
        ttk.Radiobutton(dec_frame, text="Dot (.)", variable=self.decimal_var, 
                       value="dot").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(dec_frame, text="Comma (,)", variable=self.decimal_var, 
                       value="comma").pack(side=tk.LEFT)
        
        # Format presets
        ttk.Label(format_frame, text="Quick Presets:").grid(row=0, column=2, sticky=tk.W, padx=(20, 10))
        preset_frame = ttk.Frame(format_frame)
        preset_frame.grid(row=0, column=3, sticky=tk.W, rowspan=2)
        
        ttk.Button(preset_frame, text="US/UK Format", 
                  command=lambda: self.set_format_preset("us")).pack(pady=(0, 5))
        ttk.Button(preset_frame, text="European Format", 
                  command=lambda: self.set_format_preset("eu")).pack(pady=(0, 5))
        ttk.Button(preset_frame, text="Excel Compatible", 
                  command=lambda: self.set_format_preset("excel")).pack()

        # Processing section
        process_frame = ttk.LabelFrame(main_frame, text="4. Convert Selected Files", padding="10")
        process_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        process_frame.columnconfigure(1, weight=1)
        
        self.convert_button = ttk.Button(process_frame, text="Convert to CSV", 
                                        command=self.start_conversion, state="disabled")
        self.convert_button.grid(row=0, column=0, padx=(0, 10))
        
        self.progress = ttk.Progressbar(process_frame, mode='indeterminate')
        self.progress.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        self.status_label = ttk.Label(process_frame, text="Select directory to begin")
        self.status_label.grid(row=0, column=2)
        
        # Output log
        log_frame = ttk.LabelFrame(main_frame, text="Output Log", padding="10")
        log_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, width=70)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Initial log message - Enhanced like converter2.py
        self.log("Welcome to CSV Spectra Converter!")
        self.log(f"🌍 Auto-detected region: {self.detected_region}")
        self.log(f"📄 Default CSV format: {self.default_separator} separator, {self.default_decimal} decimal")
        self.log("💡 You can change format using the presets or manual selection below")
        self.log("")
        self.log("1. Select a directory containing CSV spectral files")
        self.log("2. Choose which files to convert")
        self.log("3. Adjust CSV format options if needed")
        self.log("4. Click 'Convert to CSV' to generate unified output files")
        
    def set_format_preset(self, preset_type):
        """Set format presets for different regions/applications"""
        if preset_type == "us":
            # US/UK format: comma separator, dot decimal
            self.separator_var.set("comma")
            self.decimal_var.set("dot")
            self.log("Format set to US/UK: comma separator, dot decimal")
        elif preset_type == "eu":
            # European format: semicolon separator, comma decimal
            self.separator_var.set("semicolon")
            self.decimal_var.set("comma")
            self.log("Format set to European: semicolon separator, comma decimal")
        elif preset_type == "excel":
            # Excel compatible: tab separator, dot decimal
            self.separator_var.set("tab")
            self.decimal_var.set("dot")
            self.log("Format set to Excel compatible: tab separator, dot decimal")
        
    def log(self, message):
        """Add message to log with timestamp"""
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def select_directory(self):
        """Open directory selection dialog"""
        directory = filedialog.askdirectory(title="Select Directory with CSV spectral files")
        if directory:
            self.selected_directory.set(directory)
            self.scan_directory(directory)
            
    def scan_directory(self, directory):
        """Scan directory for CSV files"""
        self.log(f"Scanning directory: {directory}")
        
        try:
            path = Path(directory)
            self.csv_files = list(path.glob('*.csv'))
            
            # Clear previous file list
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()
            self.file_vars.clear()
            
            if self.csv_files:
                self.log(f"Found {len(self.csv_files)} CSV files:")
                
                # Create checkboxes for each file
                for i, file_path in enumerate(self.csv_files):
                    var = tk.BooleanVar(value=True)  # Default selected
                    self.file_vars[file_path] = var
                    
                    checkbox = ttk.Checkbutton(
                        self.scrollable_frame,
                        text=file_path.name,
                        variable=var,
                        command=self.update_selection_count
                    )
                    checkbox.grid(row=i, column=0, sticky=(tk.W), padx=5, pady=2)
                    
                    self.log(f"  - {file_path.name}")
                
                self.update_selection_count()
                self.convert_button.config(state="normal")
                self.status_label.config(text="Ready to convert")
                
            else:
                self.log("No CSV files found in selected directory")
                self.files_count_label.config(text="No CSV files found")
                self.convert_button.config(state="disabled")
                self.status_label.config(text="No files to convert")
                
        except Exception as e:
            self.log(f"Error scanning directory: {str(e)}")
            messagebox.showerror("Error", f"Error scanning directory:\n{str(e)}")
            
    def update_selection_count(self):
        """Update the count of selected files"""
        if self.file_vars:
            selected_count = sum(var.get() for var in self.file_vars.values())
            total_count = len(self.file_vars)
            self.files_count_label.config(text=f"{selected_count}/{total_count} files selected")
            
            # Enable/disable convert button based on selection
            if selected_count > 0:
                self.convert_button.config(state="normal")
                self.status_label.config(text="Ready to convert")
            else:
                self.convert_button.config(state="disabled")
                self.status_label.config(text="No files selected")
        
    def select_all_files(self):
        """Select all files"""
        for var in self.file_vars.values():
            var.set(True)
        self.update_selection_count()
        
    def deselect_all_files(self):
        """Deselect all files"""
        for var in self.file_vars.values():
            var.set(False)
        self.update_selection_count()
        
    def start_conversion(self):
        """Start conversion process in separate thread"""
        selected_files = [file_path for file_path, var in self.file_vars.items() if var.get()]
        
        if not selected_files:
            messagebox.showwarning("Warning", "No files selected for conversion")
            return
            
        # Disable UI during conversion
        self.convert_button.config(state="disabled")
        self.progress.start()
        self.status_label.config(text="Converting...")
        
        # Start conversion in separate thread to prevent UI freezing
        thread = threading.Thread(target=self.convert_files, args=(selected_files,))
        thread.daemon = True
        thread.start()
        
    def parse_csv_file(self, file_path):
        """Parse a CSV spectral file and extract metadata and spectral data"""
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
            content = file.read()
        
        # Dictionary to store metadata
        metadata = {}
        
        # Extract filename (without extension)
        filename = Path(file_path).stem
        metadata['Filename'] = filename
        
        # Split content into lines
        lines = content.strip().split('\n')
        
        # Find where data starts (handle both XYDATA and XYDATA; formats)
        data_start_index = -1
        for i, line in enumerate(lines):
            line_stripped = line.strip()
            if line_stripped == 'XYDATA' or line_stripped == 'XYDATA;':
                data_start_index = i + 1
                break
        
        if data_start_index == -1:
            raise ValueError("XYDATA section not found in file")
        
        # Extract metadata from lines before XYDATA
        for i in range(data_start_index - 1):
            line = lines[i].strip()
            if ';' in line and not line.startswith('XYDATA'):
                try:
                    key, value = line.split(';', 1)
                    # Clean up field names
                    clean_key = key.strip().replace(' ', '_').replace('/', '_')
                    metadata[clean_key] = value.strip()
                except:
                    pass
        
        # Extract additional metadata from end of file (if any)
        for i in range(len(lines) - 1, data_start_index, -1):
            line = lines[i].strip()
            if ';' in line and not line.replace(',', '.').replace(';', ' ').strip().replace(' ', ';').count(';') == 1:
                try:
                    key, value = line.split(';', 1)
                    clean_key = key.strip().replace(' ', '_').replace('/', '_')
                    if clean_key not in metadata:  # Don't overwrite existing metadata
                        metadata[clean_key] = value.strip()
                except:
                    pass
        
        # Extract spectral data
        x_data = []
        y_data = []
        
        for i in range(data_start_index, len(lines)):
            line = lines[i].strip()
            
            # Skip empty lines and non-data lines (metadata at end of file)
            if not line:
                continue
            if line.startswith('#') or line.startswith('[') or not ';' in line:
                continue
            
            # Split by semicolon
            parts = line.split(';')
            if len(parts) >= 2:
                try:
                    # Handle European decimal format (comma) and missing decimals
                    x_str = parts[0].strip().replace(',', '.')
                    y_str = parts[1].strip().replace(',', '.')
                    
                    # Skip if either part is empty or non-numeric
                    if not x_str or not y_str:
                        continue
                    
                    x_val = float(x_str)
                    y_val = float(y_str)
                    x_data.append(x_val)
                    y_data.append(y_val)
                except (ValueError, IndexError):
                    # Skip invalid data lines
                    continue
        
        return metadata, x_data, y_data
        
    def create_unified_x_axis(self, all_spectra_data):
        """Create a unified X axis that encompasses all spectra ranges and steps"""
        self.log("🔍 Analyzing spectral ranges and creating unified X axis...")
        
        all_x_points = set()
        file_ranges = {}
        
        # Collect all unique X points from all files
        for filename, (x_data, y_data) in all_spectra_data.items():
            if x_data:
                x_min, x_max = min(x_data), max(x_data)
                file_ranges[filename] = (x_min, x_max, len(x_data))
                
                # Add all X points to the set
                for x in x_data:
                    all_x_points.add(round(x, 4))  # Round to avoid floating point issues
        
        # Convert to sorted list
        unified_x = sorted(list(all_x_points))
        
        # Log range information
        global_min = min(unified_x)
        global_max = max(unified_x)
        self.log(f"  📊 Global range: {global_min:.1f} - {global_max:.1f} nm")
        self.log(f"  📈 Total unique X points: {len(unified_x)}")
        
        # Log individual file ranges
        for filename, (x_min, x_max, points) in file_ranges.items():
            coverage = f"{x_min:.1f}-{x_max:.1f} nm ({points} pts)"
            self.log(f"    • {filename}: {coverage}")
        
        return unified_x, file_ranges
    
    def interpolate_spectrum(self, x_original, y_original, x_target, column_name):
        """Align spectrum data onto target X axis WITHOUT interpolation - only exact matches"""
        if not x_original or not y_original:
            return [np.nan] * len(x_target)
        
        # Convert to numpy arrays for easier handling
        x_orig = np.array(x_original)
        y_orig = np.array(y_original)
        x_targ = np.array(x_target)
        
        # Initialize result with NaN
        y_aligned = np.full(len(x_target), np.nan)
        
        # Only use EXACT matches - no interpolation
        exact_matches = 0
        for i, x_val in enumerate(x_targ):
            # Find exact match in original data (with small tolerance for floating point comparison)
            matches = np.abs(x_orig - x_val) < 1e-6
            if np.any(matches):
                # Use the first exact match
                match_idx = np.where(matches)[0][0]
                y_aligned[i] = y_orig[match_idx]
                exact_matches += 1
        
        self.log(f"    ✓ {exact_matches} exact matches found (no interpolation)")
        return y_aligned.tolist()

    def convert_files(self, selected_files):
        """Convert selected files to unified CSV format with proper X-axis alignment (exact matches only)"""
        try:
            self.log(f"Starting conversion of {len(selected_files)} files...")
            
            # Dictionary to store all raw data
            all_spectra_data = {}
            all_metadata = {}
            
            # Step 1: Parse all files and collect raw data
            self.log("📁 Step 1: Parsing all files...")
            for i, file_path in enumerate(selected_files):
                self.log(f"Processing {file_path.name}...")
                
                try:
                    metadata, x_data, y_data = self.parse_csv_file(file_path)
                    
                    if x_data and y_data:
                        filename = metadata['Filename']
                        
                        # Create column name
                        title = metadata.get('TITLE', '')
                        if title:
                            column_name = f"{filename}_{title}"
                        else:
                            column_name = filename
                        
                        all_spectra_data[column_name] = (x_data, y_data)
                        all_metadata[column_name] = metadata
                        
                        x_range = f"{min(x_data):.1f}-{max(x_data):.1f}"
                        self.log(f"  ✓ {column_name}: {len(y_data)} points, range {x_range} nm")
                    else:
                        self.log(f"  ⚠ No spectral data found in {file_path.name}")
                        
                except Exception as e:
                    self.log(f"  ✗ Error processing {file_path.name}: {str(e)}")
            
            if not all_spectra_data:
                self.log("✗ No data extracted from any file!")
                messagebox.showerror("Error", "No data could be extracted from the selected files")
                return
            
            # Step 2: Create unified X axis
            self.log("⚙️ Step 2: Creating unified X axis...")
            unified_x, file_ranges = self.create_unified_x_axis(all_spectra_data)
            
            # Step 3: Align all spectra onto unified X axis (exact matches only)
            self.log("🎯 Step 3: Aligning all spectra onto unified axis (exact matches only)...")
            unified_data = {}
            
            for column_name, (x_data, y_data) in all_spectra_data.items():
                self.log(f"  Aligning {column_name} (exact matches only)...")
                aligned_y = self.interpolate_spectrum(x_data, y_data, unified_x, column_name)
                
                # Count valid (non-NaN) points
                valid_points = sum(1 for y in aligned_y if not np.isnan(y))
                total_points = len(aligned_y)
                coverage_pct = (valid_points / total_points) * 100
                
                unified_data[column_name] = aligned_y
                self.log(f"    📊 {valid_points}/{total_points} points ({coverage_pct:.1f}% coverage)")
            
            # Step 4: Create final DataFrame
            self.log(f"📋 Step 4: Creating unified CSV with {len(unified_x)} rows and {len(unified_data)} data columns...")
            
            # Start with unified X column
            df_data = {'Wavelength_nm': unified_x}
            
            # Add all interpolated Y columns
            for column_name, y_values in unified_data.items():
                df_data[column_name] = y_values
            
            # Create DataFrame
            df = pd.DataFrame(df_data)
            
            # Step 5: Apply output formatting and save
            self.log("💾 Step 5: Applying format options and saving files...")
            
            # Get selected format options
            field_sep = self.separator_var.get()
            decimal_sep = self.decimal_var.get()
            
            # Determine actual separators
            if field_sep == "comma":
                sep_char = ","
            elif field_sep == "semicolon":
                sep_char = ";"
            else:  # tab
                sep_char = "\t"
            
            self.log(f"Using field separator: {field_sep}, decimal separator: {decimal_sep}")
            
            # Save main data CSV with chosen format
            output_dir = Path(self.selected_directory.get())
            main_file = output_dir / 'unified_spectra_data.csv'
            
            if decimal_sep == "comma":
                # Replace dots with commas in numeric columns
                df_formatted = df.copy()
                for col in df_formatted.columns:
                    if df_formatted[col].dtype in ['float64', 'int64']:
                        df_formatted[col] = df_formatted[col].astype(str).str.replace('.', ',', regex=False)
                
                # Save with custom separator
                df_formatted.to_csv(main_file, index=False, sep=sep_char)
            else:
                # Use standard format
                df.to_csv(main_file, index=False, sep=sep_char)
            
            self.log(f"✓ Unified spectra data saved: {main_file}")
            self.log(f"   Format: {field_sep} field separator, {decimal_sep} decimal separator")
            
            # Create and save metadata CSV with range information
            if all_metadata:
                metadata_rows = []
                for column_name, metadata in all_metadata.items():
                    metadata_row = metadata.copy()
                    metadata_row['Column_Name'] = column_name
                    
                    # Add range information
                    if column_name in file_ranges:
                        x_min, x_max, orig_points = file_ranges[column_name]
                        metadata_row['Original_Range_Min'] = x_min
                        metadata_row['Original_Range_Max'] = x_max
                        metadata_row['Original_Points'] = orig_points
                        
                        # Calculate coverage in unified dataset
                        valid_points = sum(1 for y in unified_data[column_name] if not np.isnan(y))
                        coverage_pct = (valid_points / len(unified_x)) * 100
                        metadata_row['Unified_Valid_Points'] = valid_points
                        metadata_row['Unified_Coverage_Percent'] = round(coverage_pct, 1)
                    
                    metadata_rows.append(metadata_row)
                
                metadata_df = pd.DataFrame(metadata_rows)
                metadata_file = output_dir / 'spectra_metadata.csv'
                
                # Apply same formatting to metadata if needed
                if decimal_sep == "comma":
                    # Check for numeric columns in metadata and format them
                    for col in metadata_df.columns:
                        if metadata_df[col].dtype in ['float64', 'int64']:
                            metadata_df[col] = metadata_df[col].astype(str).str.replace('.', ',', regex=False)
                
                metadata_df.to_csv(metadata_file, index=False, sep=sep_char)
                self.log(f"✓ Enhanced metadata saved: {metadata_file}")
                self.log(f"   Includes range info and coverage statistics")
            
            self.log(f"🎉 Conversion completed successfully!")
            self.log(f"   📏 Unified dimensions: {len(df)} rows × {len(df.columns)} columns")
            self.log(f"   📊 X-axis range: {min(unified_x):.1f} - {max(unified_x):.1f} nm")
            self.log(f"   📋 Columns: Wavelength_nm + {len(df.columns)-1} aligned spectra (exact matches only)")
            
            # Show success message with detailed info
            format_info = f"{field_sep} separator, {decimal_sep} decimal"
            x_range_info = f"{min(unified_x):.1f} - {max(unified_x):.1f} nm"
            messagebox.showinfo("Success", 
                              f"Conversion completed successfully!\n\n"
                              f"Files created:\n"
                              f"• unified_spectra_data.csv\n"
                              f"  {len(df)} rows × {len(df.columns)} columns\n"
                              f"  Range: {x_range_info}\n"
                              f"  Data: Exact matches only (no interpolation)\n"
                              f"• spectra_metadata.csv\n"
                              f"  Enhanced with range & coverage info\n\n"
                              f"Format: {format_info}\n"
                              f"Location: {output_dir}")
            
        except Exception as e:
            self.log(f"✗ Conversion failed: {str(e)}")
            messagebox.showerror("Error", f"Conversion failed:\n{str(e)}")
            
        finally:
            # Re-enable UI
            self.progress.stop()
            self.convert_button.config(state="normal")
            self.status_label.config(text="Ready to convert")

def main():
    """Main function to run the application"""
    root = tk.Tk()
    app = CSVConverterGUI(root)
    
    # Center window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()