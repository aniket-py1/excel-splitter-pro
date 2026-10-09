# ⚡ Excel & CSV Splitter Pro

[![Latest Release](https://img.shields.io/github/v/release/aniket-py1/excel-splitter-pro?color=10B981&label=Release&logo=github)](https://github.com/aniket-py1/excel-splitter-pro/releases/latest)
[![Direct Download](https://img.shields.io/badge/Download-ExcelSplitterPro.exe-0ea5e9?logo=windows&logoColor=white)](https://github.com/aniket-py1/excel-splitter-pro/releases/download/v2.0.0/ExcelSplitterPro.exe)
[![Python](https://img.shields.io/badge/Python-3.9+-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![GUI](https://img.shields.io/badge/GUI-PyQt5-41CD52?logo=qt&logoColor=white)](https://pypi.org/project/PyQt5/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%20%7C%2011-0078D6?logo=windows&logoColor=white)](https://www.microsoft.com/windows)
[![License: MIT](https://img.shields.io/badge/License-MIT-F59E0B.svg)](https://opensource.org/licenses/MIT)

> **A modern, ultra-fast desktop suite engineered to partition massive Excel (`.xlsx`, `.xls`) and CSV files (1M+ rows) into customizable batches with real-time speed metrics, password decryption, and zero UI freezing.**

---

### 📥 Instant Download (No Python Required)

| Platform | Download Link | Type | Size |
| :--- | :--- | :--- | :--- |
| **Windows 64-bit** | [⬇️ **Download ExcelSplitterPro.exe (v2.0.0)**](https://github.com/aniket-py1/excel-splitter-pro/releases/download/v2.0.0/ExcelSplitterPro.exe) | Portable Executable | ~81 MB |

*Simply download and run `ExcelSplitterPro.exe` directly — no Python installation or setup needed.*

---

## 🌟 Key Features

### ⚡ 1. Ultra-Fast Large Dataset Engine (1M+ Rows)
- **Vectorized Streaming**: Direct chunked disk streaming bypasses Python row-by-row iteration bottlenecks, delivering 10x–20x faster exports.
- **Asynchronous Processing**: Heavy read/write operations execute in non-blocking background `QThread` workers, keeping the GUI fluid and responsive at all times.

### 🔒 2. Encrypted Excel Workbook Support
- **In-Memory Decryption**: Seamlessly detects password-protected Excel (`.xlsx`, `.xls`) files.
- Decrypts directly into memory via `msoffcrypto` without writing insecure unencrypted temporary files to disk.

### 🎨 3. Modern Obsidian & Slate Dark UI
- **Rich Visual Aesthetics**: Curated dark theme inspired by modern developer IDEs (`#0B0F19`, `#111827`, `#1F2937`) with vibrant Emerald (`#10B981`) and Cyan accents.
- **Interactive Drag & Drop**: Drop Excel or CSV datasets directly onto the drop zone with dynamic visual feedback.
- **DPI Scaling Support**: Smooth responsive layout that adapts cleanly to 100%, 125%, and 150% Windows display scaling.

### 📊 4. Live Performance Dashboard
Monitor execution in real time with 4 live KPI cards:
- **Rows Processed**: Exact row counter dynamically updated as batches complete.
- **Peak Performance**: Live processing throughput measured in `rows/s`.
- **Time Metrics**: Elapsed counter and dynamic remaining ETA calculation.
- **Batches Generated**: Running total of produced output files.

### 🔍 5. Flexible Column Filtering & Output Choice
- **Default Output**: Pre-configured to native **Excel (`.xlsx`)** with one-click toggle to **CSV (`.csv`)**.
- **Smart Column Selection**: Customize which columns to include with a high-contrast card list or export all columns with a single click.
- **Live Search**: Instant keyword search to find specific columns among hundreds of fields.
- **Multi-Sheet Selector**: Automatically detects all worksheets in an Excel workbook and provides an instant sheet-switching dropdown.

---

## 📸 Interface Preview

```text
┌────────────────────────────────────────────────────────────────────────┐
│  ⚡ Excel & CSV Splitter Pro                                  ● READY   │
├───────────────────────────────┬────────────────────────────────────────┤
│ 1. SOURCE FILE                │  📈 ROWS      ⚡ SPEED     ⏱ TIME      │
│   [ Drag & Drop File Here ]   │  6,655        12,400 r/s  00:03 (ETA)  │
│                               ├────────────────────────────────────────┤
│ 2. BATCH & EXPORT SETTINGS    │  Execution Progress               65%  │
│   Batch Size: [ 2000 rows ]   │  [█████████████████████░░░░░░░░]       │
│   Format: (•) Excel  ( ) CSV  ├────────────────────────────────────────┤
│                               │  [ Data Preview ]  [ Terminal Log ]    │
│ 3. COLUMN SELECTION           │  ┌──┬──────────────┬────────┬────────┐ │
│   [🔍 Filter column names... ] │  │# │ LC_Email     │ Status │ Reason │ │
│   [ Select All ] [ Clear ]    │  ├──┼──────────────┼────────┼────────┤ │
│   [✔] LC_Email                │  │1 │ user@mail.com│ clean  │ High   │ │
│   [✔] LC_Status               │  └──┴──────────────┴────────┴────────┘ │
└───────────────────────────────┴────────────────────────────────────────┘
```

---

## 🚀 Running from Source

### Prerequisites
- Python 3.9, 3.10, 3.11, 3.12, 3.13, or 3.14
- pip package manager

### 1. Clone the Repository
```bash
git clone https://github.com/aniket-py1/excel-splitter-pro.git
cd excel-splitter-pro
```

### 2. Install Required Packages
```bash
pip install PyQt5 pandas openpyxl msoffcrypto-tool
```

### 3. Launch the Application
```bash
python "Excell Spliter.py"
```

---

## 🔨 Building the Standalone Executable

To compile a custom `.exe` with embedded Windows version info and multi-resolution icons:

```bash
pip install pyinstaller
pyinstaller excel_splitter.spec --clean
```

The resulting binary will be output to `dist/ExcelSplitterPro.exe`.

---

## 🛠️ Tech Stack & Architecture

| Component | Technology | Description |
| :--- | :--- | :--- |
| **GUI Framework** | PyQt5 | High-performance native Qt bindings |
| **Data Processing** | pandas & openpyxl | Vectorized chunk reading and streaming |
| **Security & Crypto** | msoffcrypto-tool | In-memory decryption of encrypted Office files |
| **Bundling** | PyInstaller 6.x | Standalone Windows 64-bit single-file compilation |
| **Icon & UI Styling** | QSS & PIL | Custom Obsidian dark mode with multi-DPI `.ico` |

---

## 🔒 Privacy & Security

- **100% Offline**: All processing occurs locally on your machine.
- **Zero Cloud Uploads**: Your sensitive spreadsheet data never leaves your computer.
- **Secure Decryption**: Encrypted files are decrypted in temporary RAM buffers and never stored unencrypted on disk.

---

## 📄 License

Distributed under the **MIT License**. See `LICENSE` for more information.

---

<p align="center">
  <b>Developed by <a href="https://github.com/aniket-py1">Aniket Mahida</a></b><br>
  Built with ❤️ for High-Performance Data Processing
</p>