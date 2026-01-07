# CSV Spectra Merge & Combine 📊

<img width="700" height="962" alt="Screenshot 2026-01-07 171925" src="https://github.com/user-attachments/assets/399340f5-526b-4174-96a8-ae4ea3f13a1d" />

## 📝 Overview
**CSV Spectra Merge** is a GUI-based tool designed to batch process and merge multiple spectral CSV files into a single, unified dataset.

It is specifically built for laboratory spectral data containing metadata and an `XYDATA` section. The tool aligns multiple spectra onto a **single unified X-axis (Wavelength)** using **exact matches only** (no interpolation), ensuring strict data integrity.

## ✨ Key Features

* **Smart Merging:** Scans all files to create a "Master X-Axis" containing every unique wavelength point found.
* **Exact Matching:** Unlike standard interpolation tools, this software fills data only where exact wavelength matches exist. Missing points are marked as `NaN` to preserve experimental accuracy.
* **Locale Detection:** Automatically handles regional differences in CSV formats:
    * **US/UK:** Dot (`.`) for decimals, Comma (`,`) for fields.
    * **Europe:** Comma (`,`) for decimals, Semicolon (`;`) for fields.
* **Metadata Extraction:** Pulls header information (e.g., Title, Instrument Settings) regardless of whether it's at the top or bottom of the file.
* **Format Conversion:** Can output the final merged file in Excel-compatible formats or standard CSV.

## ⚠️ Input File Format
This tool is designed for **Spectral CSVs**, not generic spreadsheets. Your input files should look roughly like this:

'''text
Title;Sample A1
Date;2024-01-01
XYDATA;
350.0;0.05
350.5;0.06
351.0;0.07

* It looks for the keyword XYDATA (or XYDATA;) to start reading data.

It expects X;Y pairs (Wavelength;Intensity).

It handles both . and , as decimal separators automatically.

📋 Prerequisites
Python 3.x

pandas

numpy

🔧 Installation

Install dependencies:

pip install pandas numpy

💻 Usage

1. Run the application:

python mergecsv.py

2. Select Directory: Choose the folder containing your source .csv files.

3. Review Files: The tool lists all valid files found. You can deselect specific ones.

4. Formatting:

* Auto-detect: The tool usually guesses your region correctly.

* Manual: Use the presets (US/EU/Excel) if the output format looks wrong.

5. Convert: Click "Convert to CSV".

📂 Output Files
The tool generates two files in your source folder:

unified_spectra_data.csv:

Row 1: Headers (Filenames + Titles).

Column A: Wavelength_nm (The unified X-axis).

Columns B+: The aligned intensity data for each file.

spectra_metadata.csv:

A summary report showing the original wavelength range (Min-Max) for each file and the percentage of coverage on the new unified axis.

🤝 Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

📄 License
MIT
