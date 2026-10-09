import sys
import os
import time
import math
import traceback
from datetime import datetime
from pathlib import Path
import base64
import io

# Set Windows Application User Model ID before QApplication initialization
# This forces Windows Taskbar to display the custom application icon instead of python.exe
if sys.platform == 'win32':
    try:
        import ctypes
        app_id = 'cromsontechnologies.excelsplitterpro.highperformance.2'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass

# Data processing
import pandas as pd
import numpy as np

# Encryption support for password-protected Excel files
try:
    import msoffcrypto
    MSOFFCRYPTO_AVAILABLE = True
except ImportError:
    MSOFFCRYPTO_AVAILABLE = False

# PyQt5 GUI imports
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QComboBox, QSpinBox, QTableWidget, QTableWidgetItem,
    QFileDialog, QTextEdit, QProgressBar, QGroupBox, QHeaderView,
    QMessageBox, QFrame, QSizePolicy, QSpacerItem, QGridLayout,
    QListWidget, QListWidgetItem, QAbstractItemView, QCheckBox,
    QSplitter, QScrollArea, QTabWidget, QLineEdit, QRadioButton, QButtonGroup,
    QDialog
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, pyqtSlot, QMutex, QMutexLocker, QSize, QTimer
from PyQt5.QtGui import QFont, QColor, QIcon, QPixmap, QDragEnterEvent, QDropEvent

# Try loading embedded icon data
try:
    from icon_data import ICON_BASE64
except ImportError:
    ICON_BASE64 = ""

# Modern Design System Color Palette (Obsidian & Slate with Emerald & Cyan accents)
COLOR_BG_DARKEST   = "#0B0F19"
COLOR_BG_MAIN      = "#111827"
COLOR_BG_CARD      = "#1F2937"
COLOR_BG_HOVER     = "#2A3649"
COLOR_BG_INPUT     = "#162032"

COLOR_BORDER       = "#374151"
COLOR_BORDER_LIGHT = "#4B5563"
COLOR_BORDER_FOCUS = "#10B981"

COLOR_TEXT_PRIMARY   = "#F9FAFB"
COLOR_TEXT_SECONDARY = "#9CA3AF"
COLOR_TEXT_MUTED     = "#6B7280"

COLOR_EMERALD       = "#10B981"
COLOR_EMERALD_HOVER = "#059669"
COLOR_CYAN          = "#06B6D4"
COLOR_BLUE          = "#3B82F6"
COLOR_INDIGO        = "#6366F1"
COLOR_AMBER         = "#F59E0B"
COLOR_ROSE          = "#EF4444"
COLOR_ROSE_HOVER    = "#DC2626"


def get_app_icon():
    """Retrieve application QIcon from bundled MEIPASS, local file, or embedded base64."""
    candidates = [
        getattr(sys, '_MEIPASS', None),
        os.getcwd(),
        os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else None
    ]
    for base in candidates:
        if base:
            p = os.path.join(base, "app.ico")
            if os.path.exists(p):
                icon = QIcon(p)
                if not icon.isNull():
                    return icon

    try:
        if ICON_BASE64:
            icon_bytes = base64.b64decode(ICON_BASE64)
            pix = QPixmap()
            if pix.loadFromData(icon_bytes):
                return QIcon(pix)
    except Exception:
        pass

    return QIcon()


def is_file_encrypted(file_path):
    """Check if an Excel file is password protected using msoffcrypto."""
    if not MSOFFCRYPTO_AVAILABLE:
        return False
    ext = Path(file_path).suffix.lower()
    if ext not in ['.xlsx', '.xls']:
        return False
    try:
        with open(file_path, 'rb') as f:
            office_file = msoffcrypto.OfficeFile(f)
            return office_file.is_encrypted()
    except Exception:
        return False


def decrypt_file_to_buffer(file_path, password):
    """Decrypt an encrypted Excel file into an in-memory BytesIO buffer."""
    if not MSOFFCRYPTO_AVAILABLE:
        raise RuntimeError("msoffcrypto library is not available.")
    buf = io.BytesIO()
    with open(file_path, 'rb') as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=password)
        office_file.decrypt(buf)
    buf.seek(0)
    return buf


# -----------------------------------------------------------------------------
# Password Prompt Dialog for Encrypted Excel Workbooks
# -----------------------------------------------------------------------------
class PasswordDialog(QDialog):
    """Modern dark modal dialog to unlock password-protected files."""
    def __init__(self, filename, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Password Protected File")
        self.setFixedWidth(460)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.password = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)

        header_row = QHBoxLayout()
        lock_icon = QLabel("🔒")
        lock_icon.setStyleSheet("font-size: 28px; background: transparent;")
        header_row.addWidget(lock_icon)

        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        title_lbl = QLabel("Encrypted File Detected")
        title_lbl.setStyleSheet(f"font-size: 15px; font-weight: 700; color: {COLOR_TEXT_PRIMARY}; background: transparent;")
        sub_lbl = QLabel(f"'{filename}' is password protected.")
        sub_lbl.setStyleSheet(f"font-size: 12px; color: {COLOR_TEXT_SECONDARY}; background: transparent;")
        title_box.addWidget(title_lbl)
        title_box.addWidget(sub_lbl)
        header_row.addLayout(title_box)
        header_row.addStretch()
        layout.addLayout(header_row)

        prompt_lbl = QLabel("Enter the workbook password to decrypt and process:")
        prompt_lbl.setStyleSheet(f"color: {COLOR_TEXT_PRIMARY}; font-size: 12px; background: transparent;")
        layout.addWidget(prompt_lbl)

        self.pwd_input = QLineEdit()
        self.pwd_input.setEchoMode(QLineEdit.Password)
        self.pwd_input.setPlaceholderText("Enter password here...")
        self.pwd_input.returnPressed.connect(self.on_submit)
        layout.addWidget(self.pwd_input)

        self.show_pwd_check = QCheckBox("Show password characters")
        self.show_pwd_check.toggled.connect(
            lambda checked: self.pwd_input.setEchoMode(QLineEdit.Normal if checked else QLineEdit.Password)
        )
        layout.addWidget(self.show_pwd_check)

        self.err_lbl = QLabel("")
        self.err_lbl.setStyleSheet(f"color: {COLOR_ROSE}; font-size: 11px; font-weight: 600; background: transparent;")
        self.err_lbl.setVisible(False)
        layout.addWidget(self.err_lbl)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        btn_row.addStretch()

        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.setObjectName("subtleBtn")
        self.btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(self.btn_cancel)

        self.btn_unlock = QPushButton("🔓 Unlock & Continue")
        self.btn_unlock.setObjectName("openOutputBtn")
        self.btn_unlock.setCursor(Qt.PointingHandCursor)
        self.btn_unlock.clicked.connect(self.on_submit)
        btn_row.addWidget(self.btn_unlock)

        layout.addLayout(btn_row)

        self.setStyleSheet(f"""
            QDialog {{
                background-color: {COLOR_BG_MAIN};
                border: 1px solid {COLOR_BORDER};
                border-radius: 12px;
            }}
            QLineEdit {{
                background-color: {COLOR_BG_INPUT};
                border: 1px solid {COLOR_BORDER};
                border-radius: 8px;
                padding: 8px 12px;
                color: {COLOR_TEXT_PRIMARY};
                font-size: 13px;
            }}
            QLineEdit:focus {{
                border-color: {COLOR_EMERALD};
            }}
            QPushButton#openOutputBtn {{
                background-color: {COLOR_EMERALD};
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: 600;
            }}
            QPushButton#openOutputBtn:hover {{
                background-color: {COLOR_EMERALD_HOVER};
            }}
            QPushButton#subtleBtn {{
                background-color: transparent;
                color: {COLOR_TEXT_SECONDARY};
                border: 1px solid {COLOR_BORDER};
                border-radius: 6px;
                padding: 8px 14px;
            }}
            QPushButton#subtleBtn:hover {{
                background-color: {COLOR_BG_CARD};
                color: {COLOR_TEXT_PRIMARY};
            }}
            QCheckBox {{
                color: {COLOR_TEXT_SECONDARY};
                font-size: 11px;
            }}
        """)

    def on_submit(self):
        pwd = self.pwd_input.text()
        if not pwd:
            self.err_lbl.setText("Password cannot be empty.")
            self.err_lbl.setVisible(True)
            return
        self.password = pwd
        self.accept()


# -----------------------------------------------------------------------------
# Worker 1: Asynchronous Non-blocking File Preview Loader
# -----------------------------------------------------------------------------
class FilePreviewWorker(QThread):
    """
    Background worker that analyzes file structure, detects sheets,
    handles password decryption, and loads sample rows without freezing the UI.
    """
    preview_ready = pyqtSignal(object, list, bool, int)
    preview_failed = pyqtSignal(str)

    def __init__(self, file_path, sheet_name=None, password=None):
        super().__init__()
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.password = password

    def run(self):
        try:
            path = Path(self.file_path)
            ext = path.suffix.lower()
            sheets = []
            has_headers = True
            total_rows = 0

            if ext in ['.csv', '.tsv']:
                try:
                    with open(self.file_path, 'rb') as f:
                        buf_size = 1024 * 1024
                        lines = 0
                        while chunk := f.read(buf_size):
                            lines += chunk.count(b'\n')
                    total_rows = max(0, lines - 1)
                except Exception:
                    total_rows = 0

                df_preview = None
                encodings_to_try = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']
                sep = '\t' if ext == '.tsv' else ','

                for enc in encodings_to_try:
                    try:
                        df_preview = pd.read_csv(self.file_path, nrows=8, encoding=enc, sep=sep, low_memory=False)
                        break
                    except Exception:
                        continue

                if df_preview is None:
                    raise ValueError("Could not parse CSV file with supported encodings (UTF-8, Latin-1, CP1252).")

                try:
                    df_raw = pd.read_csv(self.file_path, nrows=1, header=None, sep=sep)
                    first_row_numeric = df_raw.iloc[0].apply(
                        lambda x: str(x).replace('.', '').replace('-', '').isdigit()
                    ).all()
                    has_headers = not first_row_numeric
                except Exception:
                    has_headers = True

                if not has_headers:
                    df_preview.columns = [f"Col {i+1}" for i in range(len(df_preview.columns))]

            elif ext in ['.xlsx', '.xls']:
                excel_source = self.file_path
                if self.password:
                    excel_source = decrypt_file_to_buffer(self.file_path, self.password)

                excel_file = pd.ExcelFile(excel_source)
                sheets = excel_file.sheet_names
                active_sheet = self.sheet_name if self.sheet_name in sheets else sheets[0]

                if isinstance(excel_source, io.BytesIO):
                    excel_source.seek(0)

                df_preview = pd.read_excel(excel_source, sheet_name=active_sheet, nrows=8)

                try:
                    if isinstance(excel_source, io.BytesIO):
                        excel_source.seek(0)
                    df_raw = pd.read_excel(excel_source, sheet_name=active_sheet, nrows=1, header=None)
                    first_row_numeric = df_raw.iloc[0].apply(
                        lambda x: str(x).replace('.', '').replace('-', '').isdigit()
                    ).all()
                    has_headers = not first_row_numeric
                except Exception:
                    has_headers = True

                if not has_headers:
                    df_preview.columns = [f"Col {i+1}" for i in range(len(df_preview.columns))]

                total_rows = len(df_preview)
            else:
                raise ValueError(f"Unsupported file format: {ext}")

            self.preview_ready.emit(df_preview, sheets, has_headers, total_rows)

        except Exception as e:
            self.preview_failed.emit(str(e))


# -----------------------------------------------------------------------------
# Worker 2: High-Performance Vectorized Batch Processor
# -----------------------------------------------------------------------------
class FileProcessor(QThread):
    """
    High-performance file splitter using vectorized pandas chunking,
    thread-safe cancellation, password decryption, and real-time metrics tracking.
    """
    progress_update      = pyqtSignal(int)
    metrics_update       = pyqtSignal(int, int, float, str, str, int, int)
    log_update           = pyqtSignal(str, str)
    status_update        = pyqtSignal(str)
    batch_saved          = pyqtSignal(str, int)
    processing_complete  = pyqtSignal(str, int, int)
    processing_cancelled = pyqtSignal()
    error_occurred       = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.file_path = None
        self.sheet_name = None
        self.password = None
        self.selected_columns = []
        self.batch_size = 2000
        self.output_folder = None
        self.output_format = "xlsx"  # Default Excel output
        self.has_headers = True
        self.split_all = False
        self.mutex = QMutex()
        self._is_running = True

    def setup(self, file_path, sheet_name, password, selected_columns, batch_size,
              output_folder, output_format, has_headers, split_all):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.password = password
        self.selected_columns = selected_columns
        self.batch_size = max(1, batch_size)
        self.output_folder = output_folder
        self.output_format = output_format.lower()
        self.has_headers = has_headers
        self.split_all = split_all
        with QMutexLocker(self.mutex):
            self._is_running = True

    def stop(self):
        with QMutexLocker(self.mutex):
            self._is_running = False

    def is_running(self):
        with QMutexLocker(self.mutex):
            return self._is_running

    def get_batch_label(self, size):
        if size >= 1000:
            if size % 1000 == 0:
                return f"{size // 1000}k"
            return f"{size / 1000:.1f}k"
        return str(size)

    def format_time(self, seconds):
        if seconds is None or math.isinf(seconds) or math.isnan(seconds):
            return "--:--"
        seconds = max(0, int(seconds))
        m, s = divmod(seconds, 60)
        h, m = divmod(m, 60)
        if h > 0:
            return f"{h:02d}:{m:02d}:{s:02d}"
        return f"{m:02d}:{s:02d}"

    def run(self):
        try:
            self.status_update.emit("Processing...")
            ext = Path(self.file_path).suffix.lower()

            if ext in ['.csv', '.tsv']:
                self._process_csv()
            elif ext in ['.xlsx', '.xls']:
                self._process_excel()
            else:
                raise ValueError(f"Unsupported format: {ext}")

            if not self.is_running():
                self.processing_cancelled.emit()
                self.status_update.emit("Cancelled")

        except Exception as e:
            if self.is_running():
                self.log_update.emit(f"Fatal error: {str(e)}", "error")
                self.error_occurred.emit(str(e))
                self.status_update.emit("Error")

    def _process_csv(self):
        self.log_update.emit("Scanning total rows...", "info")
        total_rows = 0
        try:
            with open(self.file_path, 'rb') as f:
                while chunk := f.read(1024 * 1024):
                    total_rows += chunk.count(b'\n')
            if self.has_headers and total_rows > 0:
                total_rows -= 1
        except Exception:
            total_rows = 0

        est_batches = max(1, math.ceil(total_rows / self.batch_size)) if total_rows > 0 else 1
        self.log_update.emit(f"Total rows detected: {total_rows:,} (Estimated {est_batches} files)", "info")

        usecols = None if self.split_all else (self.selected_columns if self.selected_columns else None)
        header_val = 0 if self.has_headers else None

        batch_number = 1
        rows_processed = 0
        start_time = time.time()
        last_metric_time = start_time

        reader = pd.read_csv(
            self.file_path,
            chunksize=self.batch_size,
            header=header_val,
            usecols=usecols,
            low_memory=False,
            encoding='utf-8-sig',
            engine='c'
        )

        for chunk_df in reader:
            if not self.is_running():
                break

            if not self.has_headers:
                chunk_df.columns = [f"Column {i+1}" for i in range(len(chunk_df.columns))]

            self._save_batch_df(chunk_df, batch_number)
            rows_processed += len(chunk_df)
            batch_number += 1

            now = time.time()
            if now - last_metric_time >= 0.15 or (total_rows and rows_processed >= total_rows):
                elapsed = max(0.001, now - start_time)
                speed = rows_processed / elapsed
                rem_rows = max(0, total_rows - rows_processed) if total_rows > 0 else 0
                eta = (rem_rows / speed) if speed > 0 and total_rows > 0 else 0
                pct = int((rows_processed / total_rows) * 100) if total_rows > 0 else 0

                self.progress_update.emit(min(100, pct))
                self.metrics_update.emit(
                    rows_processed,
                    total_rows,
                    speed,
                    self.format_time(elapsed),
                    self.format_time(eta),
                    batch_number - 1,
                    est_batches
                )
                last_metric_time = now

        if self.is_running():
            self.progress_update.emit(100)
            elapsed = time.time() - start_time
            speed = rows_processed / max(0.001, elapsed)
            self.metrics_update.emit(
                rows_processed,
                rows_processed,
                speed,
                self.format_time(elapsed),
                "00:00",
                batch_number - 1,
                batch_number - 1
            )
            self.status_update.emit("Completed")
            self.processing_complete.emit(self.output_folder, batch_number - 1, rows_processed)

    def _process_excel(self):
        self.log_update.emit("Reading Excel workbook...", "info")
        usecols = None if self.split_all else (self.selected_columns if self.selected_columns else None)
        header_val = 0 if self.has_headers else None

        excel_source = self.file_path
        if self.password:
            self.log_update.emit("Decrypting workbook with verified credentials...", "info")
            excel_source = decrypt_file_to_buffer(self.file_path, self.password)

        df = pd.read_excel(
            excel_source,
            sheet_name=self.sheet_name if self.sheet_name else 0,
            header=header_val,
            usecols=usecols
        )

        total_rows = len(df)
        if not self.has_headers:
            df.columns = [f"Column {i+1}" for i in range(len(df.columns))]

        est_batches = max(1, math.ceil(total_rows / self.batch_size))
        self.log_update.emit(f"Total rows in sheet: {total_rows:,} (Will generate {est_batches} files)", "info")

        batch_number = 1
        rows_processed = 0
        start_time = time.time()
        last_metric_time = start_time

        for start_idx in range(0, total_rows, self.batch_size):
            if not self.is_running():
                break

            end_idx = min(start_idx + self.batch_size, total_rows)
            batch_df = df.iloc[start_idx:end_idx]

            self._save_batch_df(batch_df, batch_number)
            batch_number += 1
            rows_processed = end_idx

            now = time.time()
            if now - last_metric_time >= 0.15 or rows_processed >= total_rows:
                elapsed = max(0.001, now - start_time)
                speed = rows_processed / elapsed
                rem_rows = max(0, total_rows - rows_processed)
                eta = (rem_rows / speed) if speed > 0 else 0
                pct = int((rows_processed / total_rows) * 100)

                self.progress_update.emit(pct)
                self.metrics_update.emit(
                    rows_processed,
                    total_rows,
                    speed,
                    self.format_time(elapsed),
                    self.format_time(eta),
                    batch_number - 1,
                    est_batches
                )
                last_metric_time = now

        if self.is_running():
            self.progress_update.emit(100)
            elapsed = time.time() - start_time
            speed = rows_processed / max(0.001, elapsed)
            self.metrics_update.emit(
                rows_processed,
                total_rows,
                speed,
                self.format_time(elapsed),
                "00:00",
                batch_number - 1,
                est_batches
            )
            self.status_update.emit("Completed")
            self.processing_complete.emit(self.output_folder, batch_number - 1, rows_processed)

    def _save_batch_df(self, df_chunk, batch_number):
        label = self.get_batch_label(self.batch_size)
        date_str = datetime.now().strftime("%d-%m-%Y")

        if self.output_format == "xlsx":
            filename = f"{batch_number} {label} {date_str}.xlsx"
            filepath = os.path.join(self.output_folder, filename)
            df_chunk.to_excel(filepath, index=False)
        else:
            filename = f"{batch_number} {label} {date_str}.csv"
            filepath = os.path.join(self.output_folder, filename)
            df_chunk.to_csv(filepath, index=False, encoding='utf-8-sig')

        self.log_update.emit(f"Batch {batch_number}: {filename} ({len(df_chunk):,} rows)", "success")
        self.batch_saved.emit(filename, len(df_chunk))


# -----------------------------------------------------------------------------
# Interactive Drag-and-Drop DropZone Widget
# -----------------------------------------------------------------------------
class ModernDropZone(QFrame):
    file_dropped = pyqtSignal(str)
    clicked = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setCursor(Qt.PointingHandCursor)
        self.is_drag_active = False

        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(16, 20, 16, 20)
        self.layout.setSpacing(8)
        self.layout.setAlignment(Qt.AlignCenter)

        self.icon_label = QLabel("📂")
        self.icon_label.setAlignment(Qt.AlignCenter)
        self.icon_label.setStyleSheet("font-size: 32px; background: transparent;")
        self.layout.addWidget(self.icon_label)

        self.primary_label = QLabel("Drag & Drop Excel or CSV File Here")
        self.primary_label.setAlignment(Qt.AlignCenter)
        self.primary_label.setStyleSheet(f"font-size: 14px; font-weight: 600; color: {COLOR_TEXT_PRIMARY}; background: transparent;")
        self.layout.addWidget(self.primary_label)

        self.secondary_label = QLabel("or click anywhere to browse files (.csv, .xlsx, .xls)")
        self.secondary_label.setAlignment(Qt.AlignCenter)
        self.secondary_label.setStyleSheet(f"font-size: 12px; color: {COLOR_TEXT_MUTED}; background: transparent;")
        self.layout.addWidget(self.secondary_label)

        self.file_info_frame = QFrame()
        self.file_info_frame.setStyleSheet(f"""
            QFrame {{
                background-color: {COLOR_BG_CARD};
                border: 1px solid {COLOR_EMERALD};
                border-radius: 6px;
                padding: 4px 8px;
            }}
        """)
        info_layout = QHBoxLayout(self.file_info_frame)
        info_layout.setContentsMargins(6, 4, 6, 4)
        info_layout.setSpacing(6)

        self.file_name_label = QLabel("")
        self.file_name_label.setStyleSheet(f"font-size: 12px; font-weight: 600; color: {COLOR_EMERALD}; background: transparent;")
        self.file_size_badge = QLabel("")
        self.file_size_badge.setStyleSheet(f"font-size: 11px; color: {COLOR_TEXT_SECONDARY}; background-color: {COLOR_BG_DARKEST}; border-radius: 4px; padding: 2px 6px;")

        info_layout.addWidget(self.file_name_label, 1)
        info_layout.addWidget(self.file_size_badge)

        self.file_info_frame.setVisible(False)
        self.layout.addWidget(self.file_info_frame)

        self.update_style()

    def update_style(self):
        border_color = COLOR_EMERALD if self.is_drag_active else COLOR_BORDER_LIGHT
        bg_color = COLOR_BG_HOVER if self.is_drag_active else COLOR_BG_INPUT
        self.setStyleSheet(f"""
            ModernDropZone {{
                background-color: {bg_color};
                border: 2px dashed {border_color};
                border-radius: 12px;
            }}
            ModernDropZone:hover {{
                border-color: {COLOR_EMERALD};
                background-color: {COLOR_BG_HOVER};
            }}
        """)

    def set_file_info(self, file_path, size_mb, is_locked=False):
        name = os.path.basename(file_path)
        lock_tag = "🔒 " if is_locked else "📄 "
        self.file_name_label.setText(f"{lock_tag}{name}")
        self.file_size_badge.setText(f"{size_mb:.1f} MB")
        self.primary_label.setText("Active File Loaded")
        self.secondary_label.setText("Click to choose a different file")
        self.file_info_frame.setVisible(True)

    def reset(self):
        self.file_info_frame.setVisible(False)
        self.primary_label.setText("Drag & Drop Excel or CSV File Here")
        self.secondary_label.setText("or click anywhere to browse files (.csv, .xlsx, .xls)")

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit()

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                ext = Path(url.toLocalFile()).suffix.lower()
                if ext in ['.csv', '.xlsx', '.xls', '.tsv']:
                    self.is_drag_active = True
                    self.update_style()
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dragLeaveEvent(self, event):
        self.is_drag_active = False
        self.update_style()

    def dropEvent(self, event: QDropEvent):
        self.is_drag_active = False
        self.update_style()
        urls = event.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            ext = Path(path).suffix.lower()
            if ext in ['.csv', '.xlsx', '.xls', '.tsv']:
                self.file_dropped.emit(path)


# -----------------------------------------------------------------------------
# Main Application Window
# -----------------------------------------------------------------------------
class ExcelSplitterPro(QMainWindow):
    def __init__(self):
        super().__init__()
        self.file_path = None
        self.file_password = None
        self.output_folder = None
        self.columns = []
        self.sheet_names = []
        self.has_headers = True

        self.preview_worker = None
        self.processor_thread = None

        self.init_ui()
        self.apply_theme()
        self.setWindowIcon(get_app_icon())

    def init_ui(self):
        self.setWindowTitle("Excel / CSV Splitter Pro - High Performance Suite")

        screen = QApplication.primaryScreen().geometry()
        init_width = min(1300, int(screen.width() * 0.86))
        init_height = min(840, int(screen.height() * 0.88))
        self.resize(init_width, init_height)
        self.setMinimumSize(980, 660)

        central_widget = QWidget()
        central_widget.setObjectName("centralWidget")
        self.setCentralWidget(central_widget)

        root_layout = QVBoxLayout(central_widget)
        root_layout.setContentsMargins(16, 12, 16, 16)
        root_layout.setSpacing(10)

        root_layout.addWidget(self.build_header())

        main_splitter = QSplitter(Qt.Horizontal)
        main_splitter.setChildrenCollapsible(False)

        sidebar_widget = self.build_sidebar()
        main_splitter.addWidget(sidebar_widget)

        workspace_widget = self.build_workspace()
        main_splitter.addWidget(workspace_widget)

        main_splitter.setStretchFactor(0, 0)
        main_splitter.setStretchFactor(1, 1)

        root_layout.addWidget(main_splitter, 1)

    # -------------------------------------------------------------------------
    # UI Component: Header Bar (Single '&' in label, no box artifacts)
    # -------------------------------------------------------------------------
    def build_header(self):
        header_card = QFrame()
        header_card.setObjectName("headerCard")
        header_card.setFixedHeight(64)

        header_layout = QHBoxLayout(header_card)
        header_layout.setContentsMargins(18, 8, 18, 8)

        title_box = QHBoxLayout()
        title_box.setSpacing(12)

        icon_label = QLabel()
        app_icon = get_app_icon()
        if not app_icon.isNull():
            icon_label.setPixmap(app_icon.pixmap(36, 36))
        else:
            icon_label.setText("⚡")
            icon_label.setStyleSheet("font-size: 26px; background: transparent;")
        icon_label.setStyleSheet("background: transparent;")
        title_box.addWidget(icon_label)

        name_box = QVBoxLayout()
        name_box.setSpacing(1)
        # Using single ampersand in QLabel so it displays cleanly as 'Excel & CSV Splitter Pro'
        app_title = QLabel("Excel & CSV Splitter Pro")
        app_title.setStyleSheet(f"font-size: 18px; font-weight: 700; color: {COLOR_TEXT_PRIMARY}; background: transparent;")
        app_subtitle = QLabel("Ultra-Fast Large Dataset Engine • 1M+ Row Optimization")
        app_subtitle.setStyleSheet(f"font-size: 11px; color: {COLOR_TEXT_SECONDARY}; background: transparent;")
        name_box.addWidget(app_title)
        name_box.addWidget(app_subtitle)

        title_box.addLayout(name_box)
        header_layout.addLayout(title_box)

        header_layout.addStretch()

        self.status_pill = QLabel("● READY")
        self.status_pill.setStyleSheet(f"""
            QLabel {{
                background-color: rgba(16, 185, 129, 0.15);
                color: {COLOR_EMERALD};
                border: 1px solid rgba(16, 185, 129, 0.4);
                border-radius: 12px;
                padding: 4px 14px;
                font-size: 11px;
                font-weight: 700;
            }}
        """)
        header_layout.addWidget(self.status_pill)

        return header_card

    # -------------------------------------------------------------------------
    # UI Component: Left Sidebar (Scrollable & Responsive)
    # -------------------------------------------------------------------------
    def build_sidebar(self):
        scroll_area = QScrollArea()
        scroll_area.setObjectName("sidebarScrollArea")
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.setMinimumWidth(440)

        container = QWidget()
        container.setObjectName("sidebarContainer")
        sidebar_layout = QVBoxLayout(container)
        sidebar_layout.setContentsMargins(6, 4, 18, 6)
        sidebar_layout.setSpacing(12)

        # 1. Source File Group
        file_box = QGroupBox("1. SOURCE FILE")
        file_layout = QVBoxLayout(file_box)
        file_layout.setContentsMargins(12, 14, 12, 12)
        file_layout.setSpacing(10)

        self.drop_zone = ModernDropZone()
        self.drop_zone.file_dropped.connect(self.load_selected_file)
        self.drop_zone.clicked.connect(self.browse_file)
        file_layout.addWidget(self.drop_zone)

        self.sheet_container = QWidget()
        sheet_layout = QHBoxLayout(self.sheet_container)
        sheet_layout.setContentsMargins(0, 0, 0, 0)
        sheet_label = QLabel("Worksheet:")
        sheet_label.setStyleSheet(f"color: {COLOR_TEXT_SECONDARY}; font-weight: 600; background: transparent;")
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentIndexChanged.connect(self.on_sheet_changed)
        sheet_layout.addWidget(sheet_label)
        sheet_layout.addWidget(self.sheet_combo, 1)
        self.sheet_container.setVisible(False)
        file_layout.addWidget(self.sheet_container)

        sidebar_layout.addWidget(file_box)

        # 2. Batch & Export Settings Group (Default to Excel XLSX output)
        config_box = QGroupBox("2. BATCH && EXPORT SETTINGS")
        config_layout = QVBoxLayout(config_box)
        config_layout.setContentsMargins(12, 14, 12, 12)
        config_layout.setSpacing(10)

        batch_row = QHBoxLayout()
        batch_label = QLabel("Batch Size:")
        batch_label.setStyleSheet(f"color: {COLOR_TEXT_PRIMARY}; font-weight: 600; background: transparent;")
        self.batch_spin = QSpinBox()
        self.batch_spin.setRange(1, 5_000_000)
        self.batch_spin.setValue(2000)
        self.batch_spin.setSingleStep(500)
        self.batch_spin.setSuffix(" rows")
        self.batch_spin.valueChanged.connect(self.update_estimated_files)
        batch_row.addWidget(batch_label)
        batch_row.addWidget(self.batch_spin, 1)
        config_layout.addLayout(batch_row)

        preset_row = QHBoxLayout()
        preset_row.setSpacing(6)
        presets = [1000, 2000, 5000, 10000, 25000, 50000]
        for p in presets:
            btn = QPushButton(f"{p//1000}k" if p >= 1000 else str(p))
            btn.setObjectName("presetBtn")
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda _, val=p: self.batch_spin.setValue(val))
            preset_row.addWidget(btn)
        config_layout.addLayout(preset_row)

        self.est_files_label = QLabel("⚡ Estimated: -- files")
        self.est_files_label.setStyleSheet(f"color: {COLOR_CYAN}; font-size: 11px; font-weight: 600; background: transparent;")
        config_layout.addWidget(self.est_files_label)

        # Output Format Selection (EXCEL IS DEFAULT)
        format_row = QHBoxLayout()
        format_label = QLabel("Output Format:")
        format_label.setStyleSheet(f"color: {COLOR_TEXT_SECONDARY}; font-weight: 600; background: transparent;")
        self.format_csv_radio = QRadioButton("CSV (.csv)")
        self.format_xlsx_radio = QRadioButton("Excel (.xlsx)")

        # DEFAULT TO EXCEL OUTPUT as requested
        self.format_xlsx_radio.setChecked(True)
        self.format_csv_radio.setChecked(False)

        format_group = QButtonGroup(self)
        format_group.addButton(self.format_xlsx_radio)
        format_group.addButton(self.format_csv_radio)

        format_row.addWidget(format_label)
        format_row.addWidget(self.format_xlsx_radio)
        format_row.addWidget(self.format_csv_radio)
        config_layout.addLayout(format_row)

        sidebar_layout.addWidget(config_box)

        # 3. Column Selection Group (BY DEFAULT DO NOT SELECT ALL COLUMNS)
        col_box = QGroupBox("3. COLUMN SELECTION")
        col_layout = QVBoxLayout(col_box)
        col_layout.setContentsMargins(12, 14, 12, 12)
        col_layout.setSpacing(8)

        # Row 1: Checkbox for Export All Columns (UNCHECKED BY DEFAULT)
        self.all_cols_checkbox = QCheckBox("Export All Columns (Bypass column filter)")
        self.all_cols_checkbox.setChecked(False)  # UNCHECKED BY DEFAULT
        self.all_cols_checkbox.toggled.connect(self.on_all_columns_toggled)
        col_layout.addWidget(self.all_cols_checkbox)

        # Row 2: Search input for filtering columns
        self.col_search_input = QLineEdit()
        self.col_search_input.setPlaceholderText("🔍 Filter column names...")
        self.col_search_input.textChanged.connect(self.filter_columns)
        self.col_search_input.setEnabled(True)
        col_layout.addWidget(self.col_search_input)

        # Row 3: Quick action buttons + High-contrast Count badge
        quick_btn_row = QHBoxLayout()
        quick_btn_row.setSpacing(6)
        self.btn_select_all = QPushButton("Select All")
        self.btn_select_all.setObjectName("subtleBtn")
        self.btn_select_all.clicked.connect(self.select_all_columns)

        self.btn_clear_cols = QPushButton("Clear")
        self.btn_clear_cols.setObjectName("subtleBtn")
        self.btn_clear_cols.clicked.connect(self.clear_column_selection)

        self.col_count_badge = QLabel("0 columns selected")
        self.col_count_badge.setStyleSheet(f"""
            QLabel {{
                color: {COLOR_EMERALD};
                font-size: 11px;
                font-weight: 700;
                background-color: {COLOR_BG_CARD};
                border: 1px solid {COLOR_BORDER};
                border-radius: 4px;
                padding: 2px 8px;
            }}
        """)

        quick_btn_row.addWidget(self.btn_select_all)
        quick_btn_row.addWidget(self.btn_clear_cols)
        quick_btn_row.addStretch()
        quick_btn_row.addWidget(self.col_count_badge)
        col_layout.addLayout(quick_btn_row)

        # Row 4: High-Contrast Column list widget
        self.column_list = QListWidget()
        self.column_list.setSelectionMode(QAbstractItemView.MultiSelection)
        self.column_list.setFixedHeight(180)
        self.column_list.setEnabled(True)
        self.column_list.itemSelectionChanged.connect(self.update_col_badge)
        col_layout.addWidget(self.column_list)

        sidebar_layout.addWidget(col_box)

        # 4. Output Destination Group
        out_box = QGroupBox("4. OUTPUT DESTINATION")
        out_layout = QVBoxLayout(out_box)
        out_layout.setContentsMargins(12, 14, 12, 12)
        out_layout.setSpacing(6)

        out_btn_row = QHBoxLayout()
        self.output_btn = QPushButton("📁 Browse Destination...")
        self.output_btn.clicked.connect(self.browse_output_folder)
        out_btn_row.addWidget(self.output_btn)
        out_layout.addLayout(out_btn_row)

        self.output_label = QLabel("Same directory as input file")
        self.output_label.setWordWrap(True)
        self.output_label.setStyleSheet(f"color: {COLOR_TEXT_MUTED}; font-size: 11px; background: transparent;")
        out_layout.addWidget(self.output_label)

        sidebar_layout.addWidget(out_box)

        # Action Buttons
        action_layout = QVBoxLayout()
        action_layout.setSpacing(8)

        self.split_btn = QPushButton("⚡ START SPLITTING")
        self.split_btn.setObjectName("splitBtn")
        self.split_btn.setCursor(Qt.PointingHandCursor)
        self.split_btn.setEnabled(False)
        self.split_btn.clicked.connect(self.start_split)
        action_layout.addWidget(self.split_btn)

        self.cancel_btn = QPushButton("⏹ CANCEL OPERATION")
        self.cancel_btn.setObjectName("cancelBtn")
        self.cancel_btn.setCursor(Qt.PointingHandCursor)
        self.cancel_btn.setVisible(False)
        self.cancel_btn.clicked.connect(self.cancel_split)
        action_layout.addWidget(self.cancel_btn)

        self.reset_btn = QPushButton("🔄 Reset All")
        self.reset_btn.setObjectName("resetBtn")
        self.reset_btn.setCursor(Qt.PointingHandCursor)
        self.reset_btn.clicked.connect(self.reset_ui)
        action_layout.addWidget(self.reset_btn)

        sidebar_layout.addLayout(action_layout)
        sidebar_layout.addStretch()

        scroll_area.setWidget(container)
        return scroll_area

    # -------------------------------------------------------------------------
    # UI Component: Right Workspace (Live KPI Dashboard, Table & Log)
    # -------------------------------------------------------------------------
    def build_workspace(self):
        workspace = QWidget()
        workspace.setObjectName("workspaceWidget")
        layout = QVBoxLayout(workspace)
        layout.setContentsMargins(8, 4, 4, 4)
        layout.setSpacing(12)

        # 1. Four Live Metric KPI Cards
        kpi_grid = QGridLayout()
        kpi_grid.setSpacing(10)

        self.kpi_rows_card = self.create_kpi_card("📊 ROWS PROCESSED", "0", "Ready to start")
        self.kpi_speed_card = self.create_kpi_card("⚡ PROCESSING SPEED", "0 rows/s", "Peak performance")
        self.kpi_time_card = self.create_kpi_card("⏱ TIME METRICS", "00:00", "ETA: --:--")
        self.kpi_batch_card = self.create_kpi_card("📦 BATCHES GENERATED", "0", "Target size: 2,000")

        kpi_grid.addWidget(self.kpi_rows_card['frame'], 0, 0)
        kpi_grid.addWidget(self.kpi_speed_card['frame'], 0, 1)
        kpi_grid.addWidget(self.kpi_time_card['frame'], 0, 2)
        kpi_grid.addWidget(self.kpi_batch_card['frame'], 0, 3)

        layout.addLayout(kpi_grid)

        # 2. Main Progress Bar Card
        progress_card = QFrame()
        progress_card.setObjectName("progressCard")
        prog_layout = QVBoxLayout(progress_card)
        prog_layout.setContentsMargins(14, 10, 14, 10)
        prog_layout.setSpacing(6)

        prog_title_row = QHBoxLayout()
        self.progress_title = QLabel("Execution Progress")
        self.progress_title.setStyleSheet(f"font-size: 12px; font-weight: 600; color: {COLOR_TEXT_PRIMARY}; background: transparent;")
        self.progress_percent = QLabel("0%")
        self.progress_percent.setStyleSheet(f"font-size: 13px; font-weight: 700; color: {COLOR_EMERALD}; background: transparent;")
        prog_title_row.addWidget(self.progress_title)
        prog_title_row.addStretch()
        prog_title_row.addWidget(self.progress_percent)
        prog_layout.addLayout(prog_title_row)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(8)
        prog_layout.addWidget(self.progress_bar)

        layout.addWidget(progress_card)

        # 3. Tabbed View: Clean Non-Clipping Tabs
        self.tab_widget = QTabWidget()
        self.tab_widget.setObjectName("workspaceTabs")

        # Tab A: Data Preview
        preview_tab = QWidget()
        preview_layout = QVBoxLayout(preview_tab)
        preview_layout.setContentsMargins(0, 8, 0, 0)

        self.preview_table = QTableWidget()
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.preview_table.horizontalHeader().setStretchLastSection(True)
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        preview_layout.addWidget(self.preview_table)

        self.tab_widget.addTab(preview_tab, "Data Preview (First 8 Rows)")

        # Tab B: Execution Terminal Console
        log_tab = QWidget()
        log_layout = QVBoxLayout(log_tab)
        log_layout.setContentsMargins(0, 8, 0, 0)
        log_layout.setSpacing(6)

        log_controls = QHBoxLayout()
        self.auto_scroll_check = QCheckBox("Auto-scroll to bottom")
        self.auto_scroll_check.setChecked(True)
        btn_clear_log = QPushButton("Clear Console")
        btn_clear_log.setObjectName("subtleBtn")
        btn_clear_log.clicked.connect(lambda: self.log_text.clear())

        log_controls.addWidget(self.auto_scroll_check)
        log_controls.addStretch()
        log_controls.addWidget(btn_clear_log)
        log_layout.addLayout(log_controls)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setObjectName("terminalLog")
        log_layout.addWidget(self.log_text)

        self.tab_widget.addTab(log_tab, "Execution Terminal & Batch Log")

        layout.addWidget(self.tab_widget, 1)

        # 4. Post-processing Completion Banner
        self.completion_banner = QFrame()
        self.completion_banner.setObjectName("completionBanner")
        self.completion_banner.setFixedHeight(48)
        banner_layout = QHBoxLayout(self.completion_banner)
        banner_layout.setContentsMargins(16, 6, 16, 6)

        self.completion_label = QLabel("✅ All batches successfully generated!")
        self.completion_label.setStyleSheet(f"color: {COLOR_EMERALD}; font-weight: 600; font-size: 13px; background: transparent;")
        banner_layout.addWidget(self.completion_label)

        banner_layout.addStretch()

        self.open_output_btn = QPushButton("📂 Open Output Folder")
        self.open_output_btn.setObjectName("openOutputBtn")
        self.open_output_btn.setCursor(Qt.PointingHandCursor)
        self.open_output_btn.clicked.connect(self.open_output_folder)
        banner_layout.addWidget(self.open_output_btn)

        self.completion_banner.setVisible(False)
        layout.addWidget(self.completion_banner)

        return workspace

    def create_kpi_card(self, title, initial_val, initial_sub):
        frame = QFrame()
        frame.setObjectName("kpiCard")
        card_layout = QVBoxLayout(frame)
        card_layout.setContentsMargins(14, 12, 14, 12)
        card_layout.setSpacing(4)

        title_lbl = QLabel(title)
        title_lbl.setStyleSheet(f"font-size: 10px; font-weight: 700; color: {COLOR_TEXT_MUTED}; letter-spacing: 0.5px; background: transparent;")
        card_layout.addWidget(title_lbl)

        val_lbl = QLabel(initial_val)
        val_lbl.setStyleSheet(f"font-size: 18px; font-weight: 800; color: {COLOR_TEXT_PRIMARY}; background: transparent;")
        card_layout.addWidget(val_lbl)

        sub_lbl = QLabel(initial_sub)
        sub_lbl.setStyleSheet(f"font-size: 11px; color: {COLOR_TEXT_SECONDARY}; background: transparent;")
        card_layout.addWidget(sub_lbl)

        return {'frame': frame, 'val': val_lbl, 'sub': sub_lbl}

    # -------------------------------------------------------------------------
    # File Selection, Password Prompt & Async Inspection
    # -------------------------------------------------------------------------
    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Data File",
            "",
            "Excel & CSV Files (*.csv *.xlsx *.xls *.tsv);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;All Files (*.*)"
        )
        if file_path:
            self.load_selected_file(file_path)

    def load_selected_file(self, file_path):
        self.file_path = file_path
        self.file_password = None
        size_mb = os.path.getsize(file_path) / (1024 * 1024)

        if is_file_encrypted(file_path):
            pwd_dialog = PasswordDialog(os.path.basename(file_path), self)
            if pwd_dialog.exec_() == QDialog.Accepted and pwd_dialog.password:
                self.file_password = pwd_dialog.password
                self.append_log(f"Password credentials provided for '{os.path.basename(file_path)}'", "info")
            else:
                self.append_log("File loading cancelled: Password required for encrypted file.", "warning")
                self.set_status("LOCKED", COLOR_AMBER)
                return

        self.drop_zone.set_file_info(file_path, size_mb, is_locked=bool(self.file_password))

        if not self.output_folder:
            self.output_folder = os.path.dirname(file_path)
            self.output_label.setText(self.output_folder)

        self.set_status("ANALYZING", COLOR_CYAN)
        self.append_log(f"Inspecting '{os.path.basename(file_path)}' ({size_mb:.1f} MB)...", "info")

        self.preview_table.clear()
        self.preview_table.setRowCount(0)
        self.preview_table.setColumnCount(0)

        self.preview_worker = FilePreviewWorker(self.file_path, password=self.file_password)
        self.preview_worker.preview_ready.connect(self.on_preview_ready)
        self.preview_worker.preview_failed.connect(self.on_preview_failed)
        self.preview_worker.start()

    @pyqtSlot(object, list, bool, int)
    def on_preview_ready(self, df_preview, sheets, has_headers, estimated_rows):
        self.has_headers = has_headers
        self.sheet_names = sheets

        if len(sheets) > 1:
            self.sheet_combo.blockSignals(True)
            self.sheet_combo.clear()
            self.sheet_combo.addItems(sheets)
            self.sheet_combo.blockSignals(False)
            self.sheet_container.setVisible(True)
        else:
            self.sheet_container.setVisible(False)

        self.populate_preview_table(df_preview)

        # Populate Column List with bright high-visibility items
        self.columns = list(df_preview.columns)
        self.column_list.clear()
        for col in self.columns:
            item = QListWidgetItem(str(col))
            # By default do NOT select columns as requested
            item.setSelected(False)
            item.setForeground(QColor("#FFFFFF"))
            item.setBackground(QColor("#1E293B"))
            font = item.font()
            font.setPointSize(10)
            font.setBold(True)
            item.setFont(font)
            self.column_list.addItem(item)

        # DEFAULT: "Export All Columns" is UNCHECKED
        self.all_cols_checkbox.setChecked(False)
        self.on_all_columns_toggled(False)
        self.update_col_badge()

        if estimated_rows > 0:
            self.kpi_rows_card['val'].setText(f"{estimated_rows:,}")
            self.kpi_rows_card['sub'].setText("Source dataset rows")
        else:
            self.kpi_rows_card['val'].setText(f"{len(self.columns)} Cols")
            self.kpi_rows_card['sub'].setText("Headers detected" if has_headers else "Generated headers")

        self.update_estimated_files()
        self.split_btn.setEnabled(True)
        self.set_status("READY", COLOR_EMERALD)
        self.append_log(f"Ready: Loaded {len(self.columns)} columns in '{os.path.basename(self.file_path)}'", "success")

    @pyqtSlot(str)
    def on_preview_failed(self, err_msg):
        self.set_status("ERROR", COLOR_ROSE)
        if "password" in err_msg.lower() or "decryption" in err_msg.lower() or "encrypted" in err_msg.lower():
            self.append_log("Decryption failed: Incorrect password provided for file.", "error")
            QMessageBox.critical(self, "Decryption Error", "Incorrect password or decryption failure.\nPlease verify the password and try again.")
        else:
            self.append_log(f"Error loading preview: {err_msg}", "error")
            QMessageBox.warning(self, "Preview Error", f"Unable to read file preview:\n{err_msg}")

    def on_sheet_changed(self):
        selected_sheet = self.sheet_combo.currentText()
        if self.file_path and selected_sheet:
            self.append_log(f"Switching preview to worksheet: '{selected_sheet}'", "info")
            self.preview_worker = FilePreviewWorker(self.file_path, sheet_name=selected_sheet, password=self.file_password)
            self.preview_worker.preview_ready.connect(self.on_preview_ready)
            self.preview_worker.preview_failed.connect(self.on_preview_failed)
            self.preview_worker.start()

    def populate_preview_table(self, df):
        if df is None or df.empty:
            return
        self.preview_table.setRowCount(len(df))
        self.preview_table.setColumnCount(len(df.columns))
        self.preview_table.setHorizontalHeaderLabels([str(c) for c in df.columns])

        for r_idx in range(len(df)):
            for c_idx in range(len(df.columns)):
                val = str(df.iloc[r_idx, c_idx])
                if len(val) > 60:
                    val = val[:57] + "..."
                item = QTableWidgetItem(val)
                self.preview_table.setItem(r_idx, c_idx, item)

        self.preview_table.resizeColumnsToContents()

    # -------------------------------------------------------------------------
    # Column Filtering & Search
    # -------------------------------------------------------------------------
    def on_all_columns_toggled(self, checked):
        self.column_list.setEnabled(not checked)
        self.col_search_input.setEnabled(not checked)
        self.btn_select_all.setEnabled(not checked)
        self.btn_clear_cols.setEnabled(not checked)

        if checked:
            self.column_list.clearSelection()
            self.col_count_badge.setText(f"All {len(self.columns)} cols selected")
        else:
            self.update_col_badge()

    def filter_columns(self, query):
        query = query.strip().lower()
        for i in range(self.column_list.count()):
            item = self.column_list.item(i)
            item.setHidden(query not in item.text().lower())

    def select_all_columns(self):
        for i in range(self.column_list.count()):
            item = self.column_list.item(i)
            if not item.isHidden():
                item.setSelected(True)
        self.update_col_badge()

    def clear_column_selection(self):
        self.column_list.clearSelection()
        self.update_col_badge()

    def update_col_badge(self):
        if self.all_cols_checkbox.isChecked():
            self.col_count_badge.setText(f"All {len(self.columns)} cols selected")
        else:
            selected = len(self.column_list.selectedItems())
            self.col_count_badge.setText(f"{selected} of {len(self.columns)} selected")

        # Programmatically guarantee high contrast on every list item
        for i in range(self.column_list.count()):
            it = self.column_list.item(i)
            if it.isSelected():
                it.setBackground(QColor("#10B981"))
                it.setForeground(QColor("#FFFFFF"))
            else:
                it.setBackground(QColor("#1E293B"))
                it.setForeground(QColor("#FFFFFF"))

    def update_estimated_files(self):
        batch_size = self.batch_spin.value()
        self.kpi_batch_card['sub'].setText(f"Target size: {batch_size:,}")
        try:
            val_text = self.kpi_rows_card['val'].text().replace(',', '')
            if val_text.isdigit():
                rows = int(val_text)
                if rows > 0:
                    est = math.ceil(rows / batch_size)
                    self.est_files_label.setText(f"⚡ Estimated: ~{est:,} output files")
                    return
        except Exception:
            pass
        self.est_files_label.setText("⚡ Estimated: depends on total rows")

    def browse_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Destination")
        if folder:
            self.output_folder = folder
            self.output_label.setText(folder)

    # -------------------------------------------------------------------------
    # Execution: Start, Cancel, Metrics & Completion
    # -------------------------------------------------------------------------
    def start_split(self):
        if not self.file_path:
            return

        if not self.output_folder:
            self.output_folder = os.path.dirname(self.file_path)

        selected_cols = []
        if not self.all_cols_checkbox.isChecked():
            selected_items = self.column_list.selectedItems()
            if not selected_items:
                QMessageBox.warning(self, "Selection Required", "Please select at least one column from the list or check 'Export All Columns'.")
                return
            selected_cols = [item.text() for item in selected_items]

        self.set_controls_enabled(False)
        self.cancel_btn.setVisible(True)
        self.cancel_btn.setEnabled(True)
        self.completion_banner.setVisible(False)
        self.tab_widget.setCurrentIndex(1)

        self.progress_bar.setValue(0)
        self.progress_percent.setText("0%")
        self.set_status("PROCESSING", COLOR_AMBER)

        out_fmt = "xlsx" if self.format_xlsx_radio.isChecked() else "csv"
        active_sheet = self.sheet_combo.currentText() if self.sheet_container.isVisible() else None

        self.append_log("=" * 60, "info")
        self.append_log("🚀 INITIATING HIGH-SPEED SPLIT", "info")
        self.append_log(f"Source File  : {self.file_path}", "info")
        if active_sheet:
            self.append_log(f"Worksheet    : {active_sheet}", "info")
        if self.file_password:
            self.append_log("Encryption   : Decrypted via verified password", "info")
        self.append_log(f"Batch Size   : {self.batch_spin.value():,} rows", "info")
        self.append_log(f"Output Format: {out_fmt.upper()}", "info")
        self.append_log(f"Destination  : {self.output_folder}", "info")
        self.append_log("=" * 60, "info")

        self.processor_thread = FileProcessor()
        self.processor_thread.setup(
            file_path=self.file_path,
            sheet_name=active_sheet,
            password=self.file_password,
            selected_columns=selected_cols,
            batch_size=self.batch_spin.value(),
            output_folder=self.output_folder,
            output_format=out_fmt,
            has_headers=self.has_headers,
            split_all=self.all_cols_checkbox.isChecked()
        )

        self.processor_thread.progress_update.connect(self.on_progress)
        self.processor_thread.metrics_update.connect(self.on_metrics)
        self.processor_thread.log_update.connect(self.append_log)
        self.processor_thread.status_update.connect(lambda s: self.set_status(s.upper(), COLOR_AMBER if s == "Processing..." else COLOR_EMERALD))
        self.processor_thread.processing_complete.connect(self.on_complete)
        self.processor_thread.processing_cancelled.connect(self.on_cancelled)
        self.processor_thread.error_occurred.connect(self.on_error)

        self.processor_thread.start()

    def cancel_split(self):
        if self.processor_thread and self.processor_thread.isRunning():
            self.cancel_btn.setEnabled(False)
            self.cancel_btn.setText("Stopping...")
            self.set_status("CANCELLING", COLOR_ROSE)
            self.append_log("Cancellation requested by user. Halting worker gracefully...", "warning")
            self.processor_thread.stop()

    @pyqtSlot(int)
    def on_progress(self, val):
        self.progress_bar.setValue(val)
        self.progress_percent.setText(f"{val}%")

    @pyqtSlot(int, int, float, str, str, int, int)
    def on_metrics(self, proc_rows, total_rows, speed, elapsed, eta, curr_batch, est_batches):
        self.kpi_rows_card['val'].setText(f"{proc_rows:,}")
        if total_rows > 0:
            pct = int((proc_rows / total_rows) * 100)
            self.kpi_rows_card['sub'].setText(f"{pct}% of {total_rows:,} rows")

        self.kpi_speed_card['val'].setText(f"{int(speed):,} rows/s")
        self.kpi_time_card['val'].setText(elapsed)
        self.kpi_time_card['sub'].setText(f"ETA: {eta}")

        self.kpi_batch_card['val'].setText(f"{curr_batch}")
        self.kpi_batch_card['sub'].setText(f"of ~{est_batches} files")

    @pyqtSlot(str, int, int)
    def on_complete(self, out_folder, total_batches, total_rows):
        self.set_status("COMPLETED", COLOR_EMERALD)
        self.set_controls_enabled(True)
        self.cancel_btn.setVisible(False)
        self.cancel_btn.setText("⏹ CANCEL OPERATION")

        self.completion_label.setText(f"✅ Splitting Complete: {total_batches:,} files ({total_rows:,} total rows) created!")
        self.completion_banner.setVisible(True)

        self.append_log("=" * 60, "success")
        self.append_log("🎉 Process finished successfully!", "success")
        self.append_log(f"Generated {total_batches:,} batch files in '{out_folder}'", "success")
        self.append_log("=" * 60, "success")

    def open_output_folder(self):
        if self.output_folder and os.path.exists(self.output_folder):
            if sys.platform == 'win32':
                os.startfile(self.output_folder)
            elif sys.platform == 'darwin':
                os.system(f'open "{self.output_folder}"')
            else:
                os.system(f'xdg-open "{self.output_folder}"')

    @pyqtSlot()
    def on_cancelled(self):
        self.set_status("CANCELLED", COLOR_ROSE)
        self.set_controls_enabled(True)
        self.cancel_btn.setVisible(False)
        self.cancel_btn.setText("⏹ CANCEL OPERATION")
        self.append_log("Operation cancelled by user.", "warning")

    @pyqtSlot(str)
    def on_error(self, err_text):
        self.set_status("ERROR", COLOR_ROSE)
        self.set_controls_enabled(True)
        self.cancel_btn.setVisible(False)
        self.cancel_btn.setText("⏹ CANCEL OPERATION")
        QMessageBox.critical(self, "Processing Error", f"An error occurred while splitting:\n{err_text}")

    def set_controls_enabled(self, enabled):
        self.split_btn.setEnabled(enabled)
        self.split_btn.setVisible(enabled)
        self.drop_zone.setEnabled(enabled)
        self.output_btn.setEnabled(enabled)
        self.batch_spin.setEnabled(enabled)
        self.all_cols_checkbox.setEnabled(enabled)
        self.format_csv_radio.setEnabled(enabled)
        self.format_xlsx_radio.setEnabled(enabled)
        self.reset_btn.setEnabled(enabled)
        if not self.all_cols_checkbox.isChecked():
            self.column_list.setEnabled(enabled)
            self.col_search_input.setEnabled(enabled)

    def set_status(self, text, color_hex):
        self.status_pill.setText(f"● {text}")
        self.status_pill.setStyleSheet(f"""
            QLabel {{
                background-color: {color_hex}22;
                color: {color_hex};
                border: 1px solid {color_hex}66;
                border-radius: 12px;
                padding: 4px 14px;
                font-size: 11px;
                font-weight: 700;
            }}
        """)

    def append_log(self, text, level="info"):
        color_map = {
            "info": COLOR_CYAN,
            "success": COLOR_EMERALD,
            "warning": COLOR_AMBER,
            "error": COLOR_ROSE
        }
        color = color_map.get(level, COLOR_TEXT_PRIMARY)
        timestamp = datetime.now().strftime("%H:%M:%S")
        html_msg = f"<span style='color: {COLOR_TEXT_MUTED};'>[{timestamp}]</span> <span style='color: {color};'>{text}</span>"
        self.log_text.append(html_msg)

        if self.auto_scroll_check.isChecked():
            sb = self.log_text.verticalScrollBar()
            sb.setValue(sb.maximum())

    def reset_ui(self):
        reply = QMessageBox.question(
            self,
            "Confirm Reset",
            "Are you sure you want to reset all configurations?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.file_path = None
            self.file_password = None
            self.output_folder = None
            self.columns = []
            self.sheet_names = []
            self.drop_zone.reset()
            self.output_label.setText("Same directory as input file")
            self.sheet_container.setVisible(False)
            self.sheet_combo.clear()
            self.batch_spin.setValue(2000)
            self.format_xlsx_radio.setChecked(True)
            self.all_cols_checkbox.setChecked(False)
            self.column_list.clear()
            self.preview_table.clear()
            self.preview_table.setRowCount(0)
            self.preview_table.setColumnCount(0)
            self.progress_bar.setValue(0)
            self.progress_percent.setText("0%")
            self.completion_banner.setVisible(False)

            self.kpi_rows_card['val'].setText("0")
            self.kpi_rows_card['sub'].setText("Ready to start")
            self.kpi_speed_card['val'].setText("0 rows/s")
            self.kpi_time_card['val'].setText("00:00")
            self.kpi_time_card['sub'].setText("ETA: --:--")
            self.kpi_batch_card['val'].setText("0")
            self.kpi_batch_card['sub'].setText("Target size: 2,000")

            self.split_btn.setEnabled(False)
            self.set_status("READY", COLOR_EMERALD)
            self.append_log("Interface reset to default state.", "info")

    def closeEvent(self, event):
        if self.processor_thread and self.processor_thread.isRunning():
            reply = QMessageBox.question(
                self,
                "Confirm Exit",
                "File splitting is actively in progress. Are you sure you want to cancel and exit?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.processor_thread.stop()
                self.processor_thread.wait(1200)
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

    # -------------------------------------------------------------------------
    # Theme: Scoped Dark Styling (No dark rectangle artifacts on labels)
    # -------------------------------------------------------------------------
    def apply_theme(self):
        self.setStyleSheet(f"""
            /* Scoped Window and Background Container */
            QMainWindow, QWidget#centralWidget, QWidget#sidebarContainer, QWidget#workspaceWidget {{
                background-color: {COLOR_BG_DARKEST};
                color: {COLOR_TEXT_PRIMARY};
                font-family: 'Segoe UI', 'SF Pro Display', Inter, -apple-system, sans-serif;
                font-size: 13px;
            }}

            /* Labels have transparent backgrounds to prevent ugly box artifacts */
            QLabel {{
                background-color: transparent;
                background: transparent;
            }}

            /* Header Card */
            QFrame#headerCard {{
                background-color: {COLOR_BG_MAIN};
                border: 1px solid {COLOR_BORDER};
                border-radius: 12px;
            }}

            /* KPI Cards */
            QFrame#kpiCard {{
                background-color: {COLOR_BG_MAIN};
                border: 1px solid {COLOR_BORDER};
                border-radius: 10px;
            }}
            QFrame#kpiCard:hover {{
                border-color: {COLOR_BORDER_LIGHT};
                background-color: {COLOR_BG_CARD};
            }}

            /* Progress Card */
            QFrame#progressCard {{
                background-color: {COLOR_BG_MAIN};
                border: 1px solid {COLOR_BORDER};
                border-radius: 10px;
            }}
            QProgressBar {{
                background-color: {COLOR_BG_DARKEST};
                border: none;
                border-radius: 4px;
            }}
            QProgressBar::chunk {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {COLOR_EMERALD}, stop:1 {COLOR_CYAN});
                border-radius: 4px;
            }}

            /* Group Boxes */
            QGroupBox {{
                background-color: {COLOR_BG_MAIN};
                border: 1px solid {COLOR_BORDER};
                border-radius: 10px;
                margin-top: 10px;
                padding-top: 16px;
                font-weight: 700;
                font-size: 11px;
                letter-spacing: 0.5px;
                color: {COLOR_TEXT_SECONDARY};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 14px;
                padding: 0 6px;
                background-color: {COLOR_BG_MAIN};
                border-radius: 4px;
            }}

            /* Inputs, Combos, SpinBoxes */
            QLineEdit, QComboBox, QSpinBox {{
                background-color: {COLOR_BG_INPUT};
                border: 1px solid {COLOR_BORDER};
                border-radius: 8px;
                padding: 7px 10px;
                color: {COLOR_TEXT_PRIMARY};
                font-size: 13px;
            }}
            QLineEdit:focus, QComboBox:focus, QSpinBox:focus {{
                border: 1px solid {COLOR_EMERALD};
            }}
            QComboBox::drop-down {{
                border: none;
                width: 24px;
            }}
            QComboBox::down-arrow {{
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 5px solid {COLOR_TEXT_SECONDARY};
                margin-right: 8px;
            }}
            QComboBox QAbstractItemView {{
                background-color: {COLOR_BG_CARD};
                border: 1px solid {COLOR_BORDER};
                selection-background-color: {COLOR_EMERALD};
                selection-color: white;
                padding: 4px;
            }}

            /* PushButtons */
            QPushButton {{
                background-color: {COLOR_BG_CARD};
                color: {COLOR_TEXT_PRIMARY};
                border: 1px solid {COLOR_BORDER};
                border-radius: 8px;
                padding: 8px 14px;
                font-weight: 600;
                font-size: 13px;
            }}
            QPushButton:hover {{
                background-color: {COLOR_BG_HOVER};
                border-color: {COLOR_BORDER_LIGHT};
            }}
            QPushButton:pressed {{
                background-color: {COLOR_BG_DARKEST};
            }}

            /* Action Buttons */
            QPushButton#splitBtn {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {COLOR_EMERALD}, stop:1 #059669);
                color: white;
                border: none;
                border-radius: 8px;
                padding: 12px 18px;
                font-size: 15px;
                font-weight: 700;
                min-height: 24px;
            }}
            QPushButton#splitBtn:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #34D399, stop:1 {COLOR_EMERALD});
            }}
            QPushButton#splitBtn:disabled {{
                background-color: {COLOR_BG_CARD};
                color: {COLOR_TEXT_MUTED};
                border: 1px solid {COLOR_BORDER};
            }}

            QPushButton#cancelBtn {{
                background-color: {COLOR_ROSE};
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px 16px;
                font-weight: 700;
            }}
            QPushButton#cancelBtn:hover {{
                background-color: {COLOR_ROSE_HOVER};
            }}

            QPushButton#resetBtn {{
                background-color: transparent;
                color: {COLOR_TEXT_SECONDARY};
                border: 1px solid {COLOR_BORDER};
            }}
            QPushButton#resetBtn:hover {{
                background-color: {COLOR_BG_CARD};
                color: {COLOR_TEXT_PRIMARY};
            }}

            QPushButton#presetBtn {{
                background-color: {COLOR_BG_CARD};
                color: {COLOR_TEXT_SECONDARY};
                border: 1px solid {COLOR_BORDER};
                border-radius: 6px;
                padding: 4px 8px;
                font-size: 11px;
                font-weight: 600;
            }}
            QPushButton#presetBtn:hover {{
                border-color: {COLOR_CYAN};
                color: {COLOR_CYAN};
            }}

            QPushButton#subtleBtn {{
                background-color: transparent;
                color: {COLOR_TEXT_SECONDARY};
                border: 1px solid {COLOR_BORDER};
                border-radius: 6px;
                padding: 4px 10px;
                font-size: 11px;
            }}
            QPushButton#subtleBtn:hover {{
                background-color: {COLOR_BG_CARD};
                color: {COLOR_TEXT_PRIMARY};
            }}

            QPushButton#openOutputBtn {{
                background-color: {COLOR_EMERALD};
                color: white;
                border: none;
                border-radius: 6px;
                padding: 6px 14px;
                font-weight: 600;
                font-size: 12px;
            }}
            QPushButton#openOutputBtn:hover {{
                background-color: {COLOR_EMERALD_HOVER};
            }}

            /* Completion Banner */
            QFrame#completionBanner {{
                background-color: rgba(16, 185, 129, 0.12);
                border: 1px solid rgba(16, 185, 129, 0.4);
                border-radius: 10px;
            }}

            /* Table Widget */
            QTableWidget {{
                background-color: {COLOR_BG_MAIN};
                alternate-background-color: {COLOR_BG_CARD};
                gridline-color: {COLOR_BORDER};
                border: 1px solid {COLOR_BORDER};
                border-radius: 8px;
                color: {COLOR_TEXT_PRIMARY};
            }}
            QTableWidget::item {{
                padding: 6px 10px;
            }}
            QTableWidget::item:selected {{
                background-color: #065F46;
                color: #FFFFFF;
            }}
            QHeaderView::section {{
                background-color: {COLOR_BG_CARD};
                color: {COLOR_TEXT_SECONDARY};
                font-weight: 700;
                font-size: 12px;
                padding: 8px 10px;
                border: none;
                border-bottom: 2px solid {COLOR_BORDER};
            }}

            /* Tabs - Clean Non-Clipping Tabs */
            QTabWidget::pane {{
                border: 1px solid {COLOR_BORDER};
                border-radius: 8px;
                background-color: {COLOR_BG_MAIN};
                top: -1px;
            }}
            QTabBar::tab {{
                background-color: {COLOR_BG_DARKEST};
                color: {COLOR_TEXT_SECONDARY};
                border: 1px solid {COLOR_BORDER};
                border-bottom: none;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                padding: 8px 20px;
                min-width: 170px;
                margin-right: 4px;
                font-weight: 600;
                font-size: 12px;
            }}
            QTabBar::tab:selected {{
                background-color: {COLOR_BG_MAIN};
                color: {COLOR_EMERALD};
                border-top: 2px solid {COLOR_EMERALD};
            }}
            QTabBar::tab:hover:!selected {{
                color: {COLOR_TEXT_PRIMARY};
                background-color: {COLOR_BG_CARD};
            }}

            /* Terminal Log */
            QTextEdit#terminalLog {{
                background-color: {COLOR_BG_INPUT};
                border: none;
                border-radius: 8px;
                font-family: 'Consolas', 'Cascadia Code', monospace;
                font-size: 12px;
                padding: 10px;
                line-height: 1.4;
            }}

            /* CheckBoxes & RadioButtons */
            QCheckBox, QRadioButton {{
                spacing: 8px;
                color: {COLOR_TEXT_PRIMARY};
                font-size: 12px;
                background: transparent;
            }}
            QCheckBox::indicator, QRadioButton::indicator {{
                width: 16px;
                height: 16px;
                border: 1px solid {COLOR_BORDER_LIGHT};
                border-radius: 4px;
                background-color: {COLOR_BG_INPUT};
            }}
            QRadioButton::indicator {{
                border-radius: 8px;
            }}
            QCheckBox::indicator:checked, QRadioButton::indicator:checked {{
                background-color: {COLOR_EMERALD};
                border-color: {COLOR_EMERALD};
            }}

            /* Column ListWidget - High Contrast, Bright White Text & Clear Emerald Selection */
            QListWidget {{
                background-color: #0F172A;
                border: 1px solid #334155;
                border-radius: 8px;
                padding: 6px;
                color: #FFFFFF;
                font-size: 13px;
                outline: none;
            }}
            QListWidget::item {{
                color: #F8FAFC;
                background-color: #1E293B;
                border: 1px solid #334155;
                border-radius: 6px;
                padding: 7px 12px;
                margin-bottom: 4px;
                font-weight: 500;
            }}
            QListWidget::item:hover {{
                background-color: #334155;
                color: #38BDF8;
                border-color: #38BDF8;
            }}
            QListWidget::item:selected {{
                background-color: #10B981;
                color: #FFFFFF;
                font-weight: 700;
                border: 1px solid #34D399;
            }}
            QListWidget::item:selected:hover {{
                background-color: #059669;
                color: #FFFFFF;
            }}
            QListWidget:disabled {{
                background-color: #0F172A;
                color: #64748B;
            }}
            QListWidget::item:disabled {{
                color: #64748B;
                background-color: #111827;
                border-color: #1E293B;
            }}

            /* Custom Minimalist Scrollbars */
            QScrollBar:vertical {{
                background-color: {COLOR_BG_DARKEST};
                width: 8px;
                margin: 0px;
            }}
            QScrollBar::handle:vertical {{
                background-color: {COLOR_BORDER};
                border-radius: 4px;
                min-height: 24px;
            }}
            QScrollBar::handle:vertical:hover {{
                background-color: {COLOR_BORDER_LIGHT};
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0px;
            }}

            /* Splitter */
            QSplitter::handle {{
                background-color: {COLOR_BORDER};
                width: 2px;
            }}
            QSplitter::handle:hover {{
                background-color: {COLOR_EMERALD};
            }}
        """)


# -----------------------------------------------------------------------------
# Main Entry Point
# -----------------------------------------------------------------------------
def main():
    if hasattr(Qt, 'AA_EnableHighDpiScaling'):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setApplicationName("Excel/CSV Splitter Pro")
    app.setOrganizationName("Cromson Technologies LLP")
    app.setStyle('Fusion')

    # Explicitly set application icon on QApplication so Windows taskbar uses it
    icon = get_app_icon()
    if not icon.isNull():
        app.setWindowIcon(icon)

    window = ExcelSplitterPro()
    window.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()