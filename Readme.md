# Excel & CSV Splitter Pro ⚡

A modern, high-performance desktop utility built with Python and PyQt5 designed to split massive Excel (`.xlsx`, `.xls`) and CSV files into smaller batches (supporting 1,000,000+ rows) with real-time speed metrics and non-blocking background threading.

---

## 🌟 Key Features

* **⚡ Ultra-Fast Streaming:** Vectorized chunking reads and writes batches directly without row-by-row dictionary overhead.
* **🛡️ Zero-Freeze UI:** File preview inspection and heavy splitting execute in dedicated background `QThread` workers.
* **📂 Interactive Drag & Drop:** Drop `.csv` or `.xlsx` files directly onto the app window.
* **📊 Live KPI Dashboard:** Real-time metrics tracking:
  * Rows processed / total rows counter
  * Processing speed in `rows/sec`
  * Elapsed time & dynamic ETA calculation
  * Batches generated counter
* **📑 Multi-Sheet Excel Support:** Automatically detects all sheets in an Excel workbook with a dynamic worksheet selector.
* **🔍 Column Search & Filtering:** Filter columns by name, select specific columns, or export all columns with one click.
* **📦 Export Format Choice:** Export split files as either **CSV (`.csv`)** or native **Excel (`.xlsx`)**.
* **⏹️ Safe Cancellation:** Instantly cancel or pause operations mid-flight without losing previous files or crashing the app.
* **🎨 Modern Dark UI:** Obsidian & Slate design system with emerald/cyan accents, responsive layout that scales cleanly across 100%, 125%, and 150% Windows display scaling.

---

## 🚀 Quick Start

### 1. Install Dependencies
```bash
pip install PyQt5 pandas openpyxl
```

### 2. Run Application
```bash
python "Excell Spliter.py"
```

---

## 🔨 Building Standalone Executable (.exe)

To compile a standalone Windows executable (`ExcelSplitterPro.exe`):
```bash
pip install pyinstaller
pyinstaller excel_splitter.spec
```
The compiled executable will be generated inside the `dist/` directory.