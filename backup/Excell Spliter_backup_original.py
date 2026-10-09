import sys
import os
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path
import threading
from queue import Queue
import traceback

# GUI imports
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QLabel, QComboBox, 
                            QSpinBox, QTableWidget, QTableWidgetItem, 
                            QFileDialog, QTextEdit, QProgressBar, QGroupBox,
                            QHeaderView, QMessageBox, QFrame, QSizePolicy,
                            QSpacerItem, QGridLayout, QListWidget, QListWidgetItem,
                            QAbstractItemView, QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QPalette, QColor, QIcon, QPixmap

# Modern color scheme
BG_PRIMARY = "#0f0f0f"
BG_SECONDARY = "#1a1a1a"
BG_CARD = "#242424"
ACCENT_GREEN = "#10b981"
ACCENT_HOVER = "#059669"
ACCENT_BLUE = "#3b82f6"
TEXT_PRIMARY = "#f3f4f6"
TEXT_SECONDARY = "#9ca3af"
BORDER = "#374151"
ERROR = "#ef4444"
WARNING = "#f59e0b"
RESET_COLOR = "#6366f1"

class FileProcessor(QThread):
    """Worker thread for processing files"""
    progress_update = pyqtSignal(int)
    log_update = pyqtSignal(str)
    status_update = pyqtSignal(str)
    error_occurred = pyqtSignal(str)
    processing_complete = pyqtSignal()
    rows_update = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        self.file_path = None
        self.selected_columns = []
        self.batch_size = 2000
        self.output_folder = None
        self.has_headers = True
        self.split_all = True
        self._is_running = True
        
    def setup(self, file_path, selected_columns, batch_size, output_folder, has_headers, split_all):
        self.file_path = file_path
        self.selected_columns = selected_columns
        self.batch_size = batch_size
        self.output_folder = output_folder
        self.has_headers = has_headers
        self.split_all = split_all
        self._is_running = True
        
    def stop(self):
        self._is_running = False
        
    def get_batch_size_label(self, size):
        if size >= 1000:
            if size % 1000 == 0:
                return f"{size // 1000}k"
            else:
                return f"{size / 1000:.1f}k"
        else:
            return str(size)
        
    def run(self):
        try:
            self.status_update.emit("Processing...")
            file_extension = Path(self.file_path).suffix.lower()
            
            if file_extension == '.csv':
                self._process_csv()
            elif file_extension in ['.xlsx', '.xls']:
                self._process_excel()
            else:
                raise ValueError(f"Unsupported file format: {file_extension}")
                
            if self._is_running:
                self.status_update.emit("Completed")
                self.processing_complete.emit()
            
        except Exception as e:
            if self._is_running:
                self.error_occurred.emit(str(e))
                self.status_update.emit("Error occurred")
            
    def _process_csv(self):
        """Process CSV files efficiently"""
        # Count total rows
        total_rows = 0
        with open(self.file_path, 'r', encoding='utf-8', errors='ignore') as f:
            total_rows = sum(1 for _ in f) - (1 if self.has_headers else 0)
        
        self.rows_update.emit(f"{total_rows:,}")
        
        # Determine columns to use
        if self.split_all:
            usecols = None
        else:
            usecols = self.selected_columns if self.selected_columns else None
        
        chunk_size = min(50000, max(10000, self.batch_size * 5))
        batch_number = 1
        current_batch_data = []
        rows_processed = 0
        
        for chunk in pd.read_csv(self.file_path, 
                                chunksize=chunk_size,
                                header=0 if self.has_headers else None,
                                usecols=usecols,
                                low_memory=False,
                                engine='c'):
            
            if not self._is_running:
                break
                
            if not self.has_headers:
                chunk.columns = [f"Column {i+1}" for i in range(len(chunk.columns))]
                
            chunk_dict = chunk.to_dict('records')
            
            for row in chunk_dict:
                if not self._is_running:
                    break
                    
                current_batch_data.append(row)
                rows_processed += 1
                
                if len(current_batch_data) >= self.batch_size:
                    self._save_batch(pd.DataFrame(current_batch_data), batch_number)
                    batch_number += 1
                    current_batch_data = []
                
                if rows_processed % 1000 == 0:
                    progress = int((rows_processed / total_rows) * 100)
                    self.progress_update.emit(progress)
        
        if current_batch_data and self._is_running:
            self._save_batch(pd.DataFrame(current_batch_data), batch_number)
            
        if self._is_running:
            self.progress_update.emit(100)
        
    def _process_excel(self):
        """Process Excel files efficiently"""
        # Determine columns to use
        if self.split_all:
            usecols = None
        else:
            usecols = self.selected_columns if self.selected_columns else None
            
        df = pd.read_excel(self.file_path, 
                          header=0 if self.has_headers else None,
                          usecols=usecols)
        total_rows = len(df)
        self.rows_update.emit(f"{total_rows:,}")
        
        if not self.has_headers:
            df.columns = [f"Column {i+1}" for i in range(len(df.columns))]
        
        batch_number = 1
        rows_processed = 0
        
        for start_idx in range(0, len(df), self.batch_size):
            if not self._is_running:
                break
                
            end_idx = min(start_idx + self.batch_size, len(df))
            batch_df = df.iloc[start_idx:end_idx]
            
            self._save_batch(batch_df, batch_number)
            batch_number += 1
            
            rows_processed = end_idx
            progress = int((rows_processed / total_rows) * 100)
            self.progress_update.emit(progress)
            
        if self._is_running:
            self.progress_update.emit(100)
        
    def _save_batch(self, df, batch_number):
        """Save batch to CSV"""
        batch_size_label = self.get_batch_size_label(self.batch_size)
        date_str = datetime.now().strftime("%d-%m-%Y")
        filename = f"{batch_number} {batch_size_label} {date_str}.csv"
        filepath = os.path.join(self.output_folder, filename)
        
        df.to_csv(filepath, index=False, encoding='utf-8-sig')
        
        log_msg = f"Batch {batch_number}: {filename} ({len(df)} rows)"
        self.log_update.emit(log_msg)

class ExcelSplitterGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.file_path = None
        self.output_folder = None
        self.columns = []
        self.has_headers = True
        self.processor_thread = None
        
        self.init_ui()
        self.apply_theme()
        self.set_app_icon()
        
    def set_app_icon(self):
        """Set application icon"""
        icon_path = "app.ico"
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
    
    def init_ui(self):
        """Initialize UI with working components"""
        self.setWindowTitle("Excel/CSV Splitter Pro")
        
        # Window sizing
        screen = QApplication.primaryScreen().geometry()
        width = min(1100, int(screen.width() * 0.8))
        height = min(750, int(screen.height() * 0.85))
        
        x = (screen.width() - width) // 2
        y = (screen.height() - height) // 2
        self.setGeometry(x, y, width, height)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Title section with logo
        title_section = QVBoxLayout()
        title_section.setSpacing(0)

        # Create a grid layout for better control
        title_grid = QGridLayout()
        title_grid.setSpacing(0)
        title_grid.setContentsMargins(0, 0, 0, 0)

        # Add logo/icon
        logo_label = QLabel()
        icon_path = "app.ico"
        if os.path.exists(icon_path):
            pixmap = QPixmap(icon_path)
            scaled_pixmap = pixmap.scaled(48, 48, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_label.setPixmap(scaled_pixmap)
        else:
            # If no icon file, use emoji as fallback
            logo_label.setText("📊")
            logo_label.setStyleSheet("font-size: 36px;")

        # Position logo in grid
        title_grid.addWidget(logo_label, 0, 0, 2, 1, Qt.AlignRight | Qt.AlignVCenter)

        # Add spacing
        spacing_label = QLabel()
        spacing_label.setFixedWidth(15)
        title_grid.addWidget(spacing_label, 0, 1)

        # Main title
        title = QLabel("Excel/CSV File Splitter Pro")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("""
            font-size: 26px;
            font-weight: bold;
            color: #10b981;
            padding: 5px;
        """)
        title_grid.addWidget(title, 0, 2, Qt.AlignLeft)

        # Subtitle directly under the title
        subtitle = QLabel("Split large files efficiently - Handles 1M+ rows")
        subtitle.setAlignment(Qt.AlignLeft)
        subtitle.setStyleSheet("""
            font-size: 14px;
            color: #9ca3af;
            padding-bottom: 10px;
            padding-top: 0px;
        """)
        title_grid.addWidget(subtitle, 1, 2, Qt.AlignLeft)

        # Center the entire grid
        grid_container = QHBoxLayout()
        grid_container.addStretch()
        grid_container.addLayout(title_grid)
        grid_container.addStretch()

        title_section.addLayout(grid_container)
        main_layout.addLayout(title_section)
        
        # Separator line
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet(f"background-color: {BORDER}; height: 2px;")
        main_layout.addWidget(line)
        
        # Content area with two columns
        content_layout = QHBoxLayout()
        content_layout.setSpacing(20)
        
        # Left column
        left_column = QVBoxLayout()
        left_column.setSpacing(15)
        
        # File selection group
        file_group = QGroupBox("File Selection")
        file_layout = QVBoxLayout()
        
        self.file_btn = QPushButton("Select File")
        self.file_btn.setCursor(Qt.PointingHandCursor)
        self.file_btn.clicked.connect(self.select_file)
        file_layout.addWidget(self.file_btn)
        
        self.file_label = QLabel("No file selected")
        self.file_label.setWordWrap(True)
        self.file_label.setStyleSheet("color: #9ca3af; padding: 5px;")
        file_layout.addWidget(self.file_label)
        
        file_group.setLayout(file_layout)
        left_column.addWidget(file_group)
        
        # Configuration group
        config_group = QGroupBox("Configuration")
        config_layout = QGridLayout()
        config_layout.setSpacing(10)
        
        # Column selection mode - Changed default to False
        config_layout.addWidget(QLabel("Column Selection:"), 0, 0)
        
        self.all_columns_radio = QCheckBox("All Columns")
        self.all_columns_radio.setChecked(False)  # Changed from True to False
        self.all_columns_radio.toggled.connect(self.toggle_column_selection)
        config_layout.addWidget(self.all_columns_radio, 0, 1)
        
        # Column list for multiple selection - Now enabled by default
        config_layout.addWidget(QLabel("Select Columns:"), 1, 0)
        self.column_list = QListWidget()
        self.column_list.setSelectionMode(QAbstractItemView.MultiSelection)
        self.column_list.setMaximumHeight(100)
        self.column_list.setEnabled(True)  # Changed from False to True
        config_layout.addWidget(self.column_list, 1, 1)
        
        # Batch size
        config_layout.addWidget(QLabel("Batch Size:"), 2, 0)
        self.batch_size_spin = QSpinBox()
        self.batch_size_spin.setMinimum(1)
        self.batch_size_spin.setMaximum(1000000)
        self.batch_size_spin.setValue(2000)
        self.batch_size_spin.setSingleStep(100)
        self.batch_size_spin.setSuffix(" rows")
        config_layout.addWidget(self.batch_size_spin, 2, 1)
        
        # Output folder
        config_layout.addWidget(QLabel("Output Folder:"), 3, 0)
        self.output_btn = QPushButton("Select Folder")
        self.output_btn.setCursor(Qt.PointingHandCursor)
        self.output_btn.clicked.connect(self.select_output_folder)
        config_layout.addWidget(self.output_btn, 3, 1)
        
        self.output_label = QLabel("Same as input file")
        self.output_label.setWordWrap(True)
        self.output_label.setStyleSheet("color: #9ca3af; padding: 5px;")
        config_layout.addWidget(self.output_label, 4, 0, 1, 2)
        
        config_group.setLayout(config_layout)
        left_column.addWidget(config_group)
        
        # Action buttons
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        # Reset button
        self.reset_btn = QPushButton("🔄 Reset")
        self.reset_btn.setCursor(Qt.PointingHandCursor)
        self.reset_btn.clicked.connect(self.reset_all)
        self.reset_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {RESET_COLOR};
                color: white;
                font-size: 14px;
                font-weight: bold;
                padding: 10px;
                border-radius: 8px;
                min-height: 35px;
            }}
            QPushButton:hover {{
                background-color: #4f46e5;
            }}
        """)
        button_layout.addWidget(self.reset_btn)
        
        # Split button
        self.split_btn = QPushButton("⚡ SPLIT FILE")
        self.split_btn.setCursor(Qt.PointingHandCursor)
        self.split_btn.setEnabled(False)
        self.split_btn.clicked.connect(self.split_file)
        self.split_btn.setStyleSheet("""
            QPushButton {
                background-color: #10b981;
                color: white;
                font-size: 16px;
                font-weight: bold;
                padding: 12px;
                border-radius: 8px;
                min-height: 40px;
            }
            QPushButton:hover {
                background-color: #059669;
            }
            QPushButton:disabled {
                background-color: #374151;
                color: #9ca3af;
            }
        """)
        button_layout.addWidget(self.split_btn, 2)
        
        left_column.addLayout(button_layout)
        left_column.addStretch()
        
        # Right column
        right_column = QVBoxLayout()
        right_column.setSpacing(15)
        
        # Preview group
        preview_group = QGroupBox("File Preview")
        preview_layout = QVBoxLayout()
        
        self.preview_table = QTableWidget()
        self.preview_table.setMaximumHeight(150)
        self.preview_table.setAlternatingRowColors(True)
        preview_layout.addWidget(self.preview_table)
        
        preview_group.setLayout(preview_layout)
        right_column.addWidget(preview_group)
        
        # Progress group
        progress_group = QGroupBox("Progress")
        progress_layout = QVBoxLayout()
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        progress_layout.addWidget(self.progress_bar)
        
        status_layout = QHBoxLayout()
        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("color: #10b981; font-weight: bold;")
        status_layout.addWidget(self.status_label)
        
        status_layout.addStretch()
        
        self.rows_label = QLabel("")
        self.rows_label.setStyleSheet("color: #3b82f6;")
        status_layout.addWidget(self.rows_label)
        
        progress_layout.addLayout(status_layout)
        progress_group.setLayout(progress_layout)
        right_column.addWidget(progress_group)
        
        # Log group
        log_group = QGroupBox("Processing Log")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(200)
        log_layout.addWidget(self.log_text)
        
        log_group.setLayout(log_layout)
        right_column.addWidget(log_group)
        
        # Add columns to content layout
        content_layout.addLayout(left_column, 1)
        content_layout.addLayout(right_column, 2)
        
        main_layout.addLayout(content_layout)
        
    def apply_theme(self):
        """Apply dark theme to the entire application"""
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: {BG_PRIMARY};
            }}
            
            QWidget {{
                background-color: {BG_PRIMARY};
                color: {TEXT_PRIMARY};
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 13px;
            }}
            
            QGroupBox {{
                background-color: {BG_CARD};
                border: 2px solid {BORDER};
                border-radius: 10px;
                margin-top: 10px;
                padding-top: 15px;
                font-weight: bold;
                font-size: 14px;
            }}
            
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 10px 0 10px;
                background-color: {BG_CARD};
                border-radius: 5px;
            }}
            
            QPushButton {{
                background-color: {ACCENT_BLUE};
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: 500;
                min-height: 30px;
            }}
            
            QPushButton:hover {{
                background-color: #2563eb;
            }}
            
            QPushButton:pressed {{
                background-color: #1d4ed8;
            }}
            
            QCheckBox {{
                spacing: 5px;
            }}
            
            QCheckBox::indicator {{
                width: 18px;
                height: 18px;
                border-radius: 3px;
                border: 2px solid {BORDER};
                background-color: {BG_CARD};
            }}
            
            QCheckBox::indicator:checked {{
                background-color: {ACCENT_GREEN};
                border-color: {ACCENT_GREEN};
            }}
            
            QListWidget {{
                background-color: {BG_CARD};
                border: 1px solid {BORDER};
                border-radius: 6px;
                padding: 5px;
            }}
            
            QListWidget::item {{
                padding: 3px;
                border-radius: 3px;
            }}
            
            QListWidget::item:selected {{
                background-color: {ACCENT_BLUE};
                color: white;
            }}
            
            QListWidget::item:hover {{
                background-color: {BORDER};
            }}
            
            QComboBox {{
                background-color: {BG_CARD};
                border: 1px solid {BORDER};
                border-radius: 6px;
                padding: 6px;
                min-height: 25px;
            }}
            
            QComboBox:hover {{
                border-color: {ACCENT_BLUE};
            }}
            
            QComboBox::drop-down {{
                border: none;
                width: 25px;
            }}
            
            QComboBox::down-arrow {{
                image: none;
                border-style: solid;
                border-width: 5px;
                border-color: {TEXT_SECONDARY} transparent transparent transparent;
            }}
            
            QComboBox QAbstractItemView {{
                background-color: {BG_CARD};
                border: 1px solid {BORDER};
                selection-background-color: {ACCENT_BLUE};
            }}
            
            QSpinBox {{
                background-color: {BG_CARD};
                border: 1px solid {BORDER};
                border-radius: 6px;
                padding: 6px;
                min-height: 25px;
            }}
            
            QSpinBox:hover {{
                border-color: {ACCENT_BLUE};
            }}
            
            QSpinBox::up-button, QSpinBox::down-button {{
                background-color: transparent;
                border: none;
                width: 20px;
            }}
            
            QSpinBox::up-arrow {{
                image: none;
                border-style: solid;
                border-width: 0 4px 5px 4px;
                border-color: transparent transparent {TEXT_SECONDARY} transparent;
            }}
            
            QSpinBox::down-arrow {{
                image: none;
                border-style: solid;
                border-width: 5px 4px 0 4px;
                border-color: {TEXT_SECONDARY} transparent transparent transparent;
            }}
            
            QTableWidget {{
                background-color: {BG_CARD};
                alternate-background-color: {BG_SECONDARY};
                gridline-color: {BORDER};
                border: 1px solid {BORDER};
                border-radius: 6px;
            }}
            
            QTableWidget::item {{
                padding: 5px;
            }}
            
            QHeaderView::section {{
                background-color: {BG_SECONDARY};
                padding: 8px;
                border: none;
                border-bottom: 2px solid {ACCENT_GREEN};
                font-weight: bold;
            }}
            
            QTextEdit {{
                background-color: {BG_CARD};
                border: 1px solid {BORDER};
                border-radius: 6px;
                padding: 8px;
                font-family: 'Consolas', monospace;
                font-size: 12px;
            }}
            
            QProgressBar {{
                background-color: {BG_CARD};
                border: 1px solid {BORDER};
                border-radius: 6px;
                text-align: center;
                font-weight: bold;
                min-height: 22px;
            }}
            
            QProgressBar::chunk {{
                background-color: {ACCENT_GREEN};
                border-radius: 5px;
            }}
            
            QLabel {{
                background-color: transparent;
            }}
            
            QScrollBar:vertical {{
                background-color: {BG_SECONDARY};
                width: 12px;
                border: none;
            }}
            
            QScrollBar::handle:vertical {{
                background-color: {BORDER};
                border-radius: 6px;
                min-height: 30px;
            }}
            
            QScrollBar::handle:vertical:hover {{
                background-color: {ACCENT_BLUE};
            }}
        """)
        
    def toggle_column_selection(self, checked):
        """Toggle between all columns and specific column selection"""
        self.column_list.setEnabled(not checked)
        if checked:
            self.column_list.clearSelection()
            
    def select_file(self):
        """Handle file selection"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel or CSV File",
            "",
            "Data Files (*.csv *.xlsx *.xls);;All Files (*.*)"
        )
        
        if file_path:
            self.file_path = file_path
            filename = os.path.basename(file_path)
            size_mb = os.path.getsize(file_path) / (1024 * 1024)
            
            self.file_label.setText(f"{filename}\nSize: {size_mb:.1f} MB")
            self.file_label.setStyleSheet("color: #f3f4f6; padding: 5px;")
            
            # Set default output folder
            if not self.output_folder:
                self.output_folder = os.path.dirname(file_path)
                self.output_label.setText("Same as input file")
            
            # Load preview
            self.load_file_preview()
            self.split_btn.setEnabled(True)
            
    def select_output_folder(self):
        """Handle output folder selection"""
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        
        if folder:
            self.output_folder = folder
            self.output_label.setText(folder)
            self.output_label.setStyleSheet("color: #f3f4f6; padding: 5px;")
            
    def load_file_preview(self):
        """Load and display file preview"""
        try:
            file_ext = Path(self.file_path).suffix.lower()
            
            # Read first 5 rows
            if file_ext == '.csv':
                df = pd.read_csv(self.file_path, nrows=5)
                df_no_header = pd.read_csv(self.file_path, nrows=5, header=None)
            else:
                df = pd.read_excel(self.file_path, nrows=5)
                df_no_header = pd.read_excel(self.file_path, nrows=5, header=None)
            
            # Check for headers
            try:
                first_row_numeric = df_no_header.iloc[0].apply(
                    lambda x: str(x).replace('.', '').replace('-', '').isdigit()
                ).all()
                self.has_headers = not first_row_numeric
            except:
                self.has_headers = True
            
            if not self.has_headers:
                df = df_no_header
                df.columns = [f"Column {i+1}" for i in range(len(df.columns))]
            
            # Update preview table
            self.preview_table.setRowCount(len(df))
            self.preview_table.setColumnCount(len(df.columns))
            self.preview_table.setHorizontalHeaderLabels([str(col) for col in df.columns])
            
            for row in range(len(df)):
                for col in range(len(df.columns)):
                    value = str(df.iloc[row, col])
                    if len(value) > 50:
                        value = value[:47] + "..."
                    item = QTableWidgetItem(value)
                    self.preview_table.setItem(row, col, item)
            
            self.preview_table.resizeColumnsToContents()
            
            # Update column list
            self.columns = list(df.columns)
            self.column_list.clear()
            for col in self.columns:
                self.column_list.addItem(str(col))
            
            # Log
            self.log_text.append(f"✅ Loaded: {os.path.basename(self.file_path)}")
            self.log_text.append(f"📊 Found {len(df.columns)} columns")
            self.log_text.append(f"📄 Headers detected: {'Yes' if self.has_headers else 'No'}")
            
        except Exception as e:
            self.log_text.append(f"❌ Error loading file: {str(e)}")
            QMessageBox.warning(self, "Error", f"Failed to load file preview:\n{str(e)}")
            
    def reset_all(self):
        """Reset all fields and clear the interface"""
        # Confirm reset
        reply = QMessageBox.question(
            self,
            "Confirm Reset",
            "Are you sure you want to reset all fields?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            # Reset file selection
            self.file_path = None
            self.file_label.setText("No file selected")
            self.file_label.setStyleSheet("color: #9ca3af; padding: 5px;")
            
            # Reset output folder
            self.output_folder = None
            self.output_label.setText("Same as input file")
            self.output_label.setStyleSheet("color: #9ca3af; padding: 5px;")
            
            # Reset configuration - Changed to keep default as False
            self.column_list.clear()
            self.all_columns_radio.setChecked(False)  # Changed from True to False
            self.column_list.setEnabled(True)  # Keep enabled
            self.batch_size_spin.setValue(2000)
            
            # Clear preview table
            self.preview_table.clear()
            self.preview_table.setRowCount(0)
            self.preview_table.setColumnCount(0)
            
            # Reset progress
            self.progress_bar.setValue(0)
            self.status_label.setText("Ready")
            self.status_label.setStyleSheet("color: #10b981; font-weight: bold;")
            self.rows_label.setText("")


                        # Clear log
            self.log_text.clear()
            self.log_text.append("🔄 Interface reset successfully")
            
            # Disable split button
            self.split_btn.setEnabled(False)
            
    def split_file(self):
        """Start the file splitting process"""
        if not self.file_path:
            return
            
        # Validate output folder
        if not self.output_folder:
            self.output_folder = os.path.dirname(self.file_path)
            
        # Get selected columns
        selected_columns = []
        if not self.all_columns_radio.isChecked():
            selected_items = self.column_list.selectedItems()
            if not selected_items:
                QMessageBox.warning(self, "Warning", "Please select at least one column or choose 'All Columns'")
                return
            selected_columns = [item.text() for item in selected_items]
            
        # Disable controls during processing
        self.split_btn.setEnabled(False)
        self.file_btn.setEnabled(False)
        self.output_btn.setEnabled(False)
        self.column_list.setEnabled(False)
        self.all_columns_radio.setEnabled(False)
        self.batch_size_spin.setEnabled(False)
        self.reset_btn.setEnabled(False)
        
        # Reset progress
        self.progress_bar.setValue(0)
        self.status_label.setText("Processing...")
        self.status_label.setStyleSheet("color: #f59e0b; font-weight: bold;")
        
        # Clear and update log
        self.log_text.append("\n" + "="*50)
        self.log_text.append(f"🚀 Starting split process")
        self.log_text.append(f"📁 Output folder: {self.output_folder}")
        self.log_text.append(f"📏 Batch size: {self.batch_size_spin.value():,} rows")
        
        if self.all_columns_radio.isChecked():
            self.log_text.append(f"🔍 Splitting: All columns")
        else:
            self.log_text.append(f"🔍 Splitting: {len(selected_columns)} selected columns")
            for col in selected_columns:
                self.log_text.append(f"   • {col}")
                
        self.log_text.append("="*50 + "\n")
        
        # Create and start processor thread
        self.processor_thread = FileProcessor()
        self.processor_thread.setup(
            self.file_path,
            selected_columns,
            self.batch_size_spin.value(),
            self.output_folder,
            self.has_headers,
            self.all_columns_radio.isChecked()
        )
        
        # Connect signals
        self.processor_thread.progress_update.connect(self.update_progress)
        self.processor_thread.log_update.connect(self.update_log)
        self.processor_thread.status_update.connect(self.update_status)
        self.processor_thread.error_occurred.connect(self.handle_error)
        self.processor_thread.processing_complete.connect(self.processing_complete)
        self.processor_thread.rows_update.connect(self.update_rows)
        
        # Start processing
        self.processor_thread.start()
        
    def update_progress(self, value):
        """Update progress bar"""
        self.progress_bar.setValue(value)
        
    def update_log(self, message):
        """Add message to log"""
        self.log_text.append(f"  ✓ {message}")
        # Auto-scroll to bottom
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
        
    def update_status(self, status):
        """Update status label"""
        self.status_label.setText(status)
        if status == "Completed":
            self.status_label.setStyleSheet("color: #10b981; font-weight: bold;")
        elif "Error" in status:
            self.status_label.setStyleSheet("color: #ef4444; font-weight: bold;")
        
    def update_rows(self, rows):
        """Update row count"""
        self.rows_label.setText(f"Total rows: {rows}")
        
    def handle_error(self, error):
        """Handle processing errors"""
        self.log_text.append(f"\n❌ ERROR: {error}")
        QMessageBox.critical(self, "Processing Error", f"An error occurred:\n{error}")
        self.enable_controls()
        
    def processing_complete(self):
        """Handle successful completion"""
        self.log_text.append("\n" + "="*50)
        self.log_text.append("✅ File splitting completed successfully!")
        self.log_text.append("="*50)
        
        # Show success dialog
        msg = QMessageBox(self)
        msg.setWindowTitle("Success")
        msg.setText("File splitting completed successfully!")
        msg.setInformativeText(f"Output files saved to:\n{self.output_folder}")
        msg.setIcon(QMessageBox.Information)
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Open)
        msg.setDefaultButton(QMessageBox.Ok)
        
        if msg.exec_() == QMessageBox.Open:
            # Open output folder
            if sys.platform == 'win32':
                os.startfile(self.output_folder)
            elif sys.platform == 'darwin':
                os.system(f'open "{self.output_folder}"')
            else:
                os.system(f'xdg-open "{self.output_folder}"')
        
        self.enable_controls()
        
    def enable_controls(self):
        """Re-enable all controls"""
        self.split_btn.setEnabled(True)
        self.file_btn.setEnabled(True)
        self.output_btn.setEnabled(True)
        self.column_list.setEnabled(not self.all_columns_radio.isChecked())
        self.all_columns_radio.setEnabled(True)
        self.batch_size_spin.setEnabled(True)
        self.reset_btn.setEnabled(True)
        self.status_label.setText("Ready")
        self.status_label.setStyleSheet("color: #10b981; font-weight: bold;")
        
    def closeEvent(self, event):
        """Handle window close event"""
        if self.processor_thread and self.processor_thread.isRunning():
            reply = QMessageBox.question(
                self,
                "Confirm Exit",
                "File processing is in progress. Are you sure you want to exit?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.processor_thread.stop()
                self.processor_thread.wait(1000)
                if self.processor_thread.isRunning():
                    self.processor_thread.terminate()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

def main():
    """Main entry point"""
    # Enable High DPI support BEFORE creating QApplication
    if hasattr(Qt, 'AA_EnableHighDpiScaling'):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    app = QApplication(sys.argv)
    
    # App settings
    app.setApplicationName("Excel/CSV Splitter Pro")
    app.setOrganizationName("Cromson Technologies LLP")
    app.setStyle('Fusion')
    
    # Create and show window
    window = ExcelSplitterGUI()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()