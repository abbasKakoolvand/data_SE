import os
import time
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QLineEdit, QPushButton, QProgressBar,
                             QLabel, QFileDialog, QMessageBox, QGroupBox, QGridLayout, QDialog)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, pyqtSlot, QObject, QSize
from PyQt6.QtGui import QIcon, QFont, QPixmap
import pandas as pd
from openpyxl import Workbook

STYLE_SHEET = """
QMainWindow {
    background-color: #2D2D2D;
    color: #FFFFFF;  /* Light font color */
}

QGroupBox {
    border: 1px solid #3D3D3D;
    border-radius: 5px;
    margin-top: 10px;
    padding-top: 15px;
    color: #FFFFFF;  /* Light font color */
}

QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 3px;
    color: #88B0FF;  /* Light accent color */
}

QLineEdit {
    background-color: #404040;
    border: 1px solid #3D3D3D;
    border-radius: 3px;
    padding: 5px;
    font-size: 12px;
    color: #FFFFFF;  /* Light font color */
}

QPushButton {
    background-color: #404040;
    border: 1px solid #3D3D3D;
    border-radius: 3px;
    padding: 5px 10px;
    min-width: 80px;
    color: #FFFFFF;  /* Light font color */
}

QPushButton:hover {
    background-color: #505050;
}

QPushButton:pressed {
    background-color: #303030;
}

QProgressBar {
    border: 1px solid #3D3D3D;
    border-radius: 3px;
    text-align: center;
    background-color: #404040;
    color: white;  /* Light font color */
}

QProgressBar::chunk {
    background-color: #4CAF50;
    border-radius: 2px;
}

QLabel {
    font-size: 12px;
    color: #FFFFFF;  /* Light font color */
}

QDialog {
    background-color: #2D2D2D;
    color: #FFFFFF;  /* Light font color */
}

QMessageBox {
    background-color: #2D2D2D;
    color: #FFFFFF;  /* Light font color */
}

QMessageBox QLabel {
    color: #FFFFFF;  /* Light font color */
}

QMessageBox QPushButton {
    background-color: #404040;
    color: #FFFFFF;  /* Light font color */
}
"""


class Worker(QObject):
    progress = pyqtSignal(int, int, float, float, float, float, int)  # Added records count
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, folder_path, output_path):
        super().__init__()
        self.folder_path = folder_path
        self.output_path = output_path
        self._is_running = True
        self.total_records = 0  # Track total records

    def stop(self):
        self._is_running = False

    def run(self):
        try:
            # Phase 1: File enumeration
            file_list = []
            total_size = 0
            total_files = 0

            for root, _, files in os.walk(self.folder_path):
                for file in files:
                    if not self._is_running:
                        return
                    if file.lower().endswith(('.xlsx', '.xls', '.csv')):
                        file_path = os.path.join(root, file)
                        file_size = os.path.getsize(file_path)
                        file_list.append((file_path, file_size))
                        total_size += file_size
                        total_files += 1

            # Phase 2: File processing
            processed_files = 0
            processed_size = 0
            start_time = time.time()

            wb = Workbook(write_only=True)
            ws = wb.create_sheet()
            ws.append(['File Name', 'Path', 'Sheet Name', 'Column Name'])

            for file_path, file_size in file_list:
                if not self._is_running:
                    break

                processed_size += file_size
                processed_files += 1
                elapsed = time.time() - start_time
                avg_time_per_file = elapsed / processed_files if processed_files else 0
                remaining_time = avg_time_per_file * (total_files - processed_files)
                records_added = 0  # Records added for current file

                try:
                    if file_path.lower().endswith(('.xlsx', '.xls')):
                        with pd.ExcelFile(file_path) as excel:
                            for sheet_name in excel.sheet_names:
                                # Read the Excel file without specifying the header
                                df = pd.read_excel(excel, sheet_name=sheet_name, header=None, nrows=10)
                                print([col for col in df.columns.tolist()])
                                # Find the first non-null row
                                first_non_null_index = df.first_valid_index()

                                if first_non_null_index is not None:
                                    # Set the first non-null row as the header
                                    new_header = df.iloc[first_non_null_index]  # Get the first non-null row
                                    df = df[1:]  # Take the data less the header row
                                    df.columns = new_header  # Set the new header

                                # Now df has the first non-null row as the columns
                                columns = df.columns.tolist()
                                records_added += len(columns)
                                for col in columns:
                                    ws.append([
                                        os.path.basename(file_path),
                                        file_path,
                                        sheet_name,
                                        col
                                    ])
                    elif file_path.lower().endswith('.csv'):
                        df = pd.read_csv(file_path, nrows=0)
                        columns = df.columns.tolist()
                        records_added += len(columns)
                        for col in columns:
                            ws.append([
                                os.path.basename(file_path),
                                file_path,
                                'CSV',
                                col
                            ])

                    self.total_records += records_added

                except Exception as e:
                    self.error.emit(f"Error processing {file_path}: {str(e)}")

                # Emit progress with records count
                self.progress.emit(
                    processed_files,
                    total_files,
                    processed_size,
                    total_size - processed_size,
                    elapsed,
                    remaining_time,
                    self.total_records
                )

            # Create directory if needed and save
            os.makedirs(os.path.dirname(self.output_path), exist_ok=True)
            wb.save(self.output_path)
            self.finished.emit()

        except Exception as e:
            self.error.emit(str(e))


TEAM_LOGO_PATH = "team_logo.png"  # Path to your team logo (PNG format recommended)
APP_LOGO_PATH = "icon.png"  # Path to your team logo (PNG format recommended)
APP_NAME = "Excel/CSV Metadata Scanner"  # Path to your team logo (PNG format recommended)
TEAM_NAME = "SPM BI"
TEAM_DESCRIPTION = """
CSO
MCI

"""


class AboutTeamDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About")
        self.setWindowIcon(QIcon(TEAM_LOGO_PATH))
        self.setModal(True)
        self.setMaximumHeight(400)
        self.setMaximumHeight(400)
        self.setMaximumWidth(300)
        self.setMinimumWidth(300)

        layout = QVBoxLayout()

        # Team Logo
        logo_label_app = QLabel(self)
        logo_pixmap_app = QPixmap(APP_LOGO_PATH)
        logo_label_app.setPixmap(logo_pixmap_app.scaled(200, 200, Qt.AspectRatioMode.KeepAspectRatio))
        logo_label_app.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(logo_label_app)

        # Team Name
        app_name_label = QLabel(f"{APP_NAME}\n")
        app_name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        app_name_label.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(app_name_label)

        # Team Name
        team_name_label = QLabel(f"A product of {TEAM_NAME} Team")
        team_name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        team_name_label.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(team_name_label)

        # Team Logo
        logo_label = QLabel(self)
        logo_pixmap = QPixmap(TEAM_LOGO_PATH)
        logo_label.setPixmap(logo_pixmap.scaled(200, 200, Qt.AspectRatioMode.KeepAspectRatio))
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(logo_label)

        # Team Description
        team_description_label = QLabel(TEAM_DESCRIPTION)
        team_description_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        team_description_label.setWordWrap(True)
        layout.addWidget(team_description_label)

        self.setLayout(layout)

    def mouseMoveEvent(self, event):
        # Override to do nothing on mouse move
        pass


class ModernMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} (SPM BI team - CSO - MCI)")
        self.setMinimumSize(800, 500)
        self.setWindowIcon(QIcon("icon.ico"))

        # Central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)

        # Folder selection section
        self.folder_input = QLineEdit()
        self.folder_input.setPlaceholderText("Select a folder to scan...")
        self.folder_input.setReadOnly(True)
        self.browse_button = QPushButton("Browse Folder")
        self.browse_button.setIcon(QIcon.fromTheme("folder-open"))

        folder_layout = QHBoxLayout()
        folder_layout.addWidget(self.folder_input)
        folder_layout.addWidget(self.browse_button)

        # Progress section
        progress_group = QGroupBox("Scan Progress")
        progress_layout = QVBoxLayout(progress_group)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFixedHeight(25)

        # Stats grid with added records count
        stats_grid = QGridLayout()
        self.processed_files_label = self.create_stat_label()
        self.remaining_files_label = self.create_stat_label()
        self.processed_volume_label = self.create_stat_label()
        self.remaining_volume_label = self.create_stat_label()
        self.elapsed_time_label = self.create_stat_label()
        self.remaining_time_label = self.create_stat_label()
        self.total_records_label = self.create_stat_label()  # New records counter

        stats_grid.addWidget(QLabel("Files Processed:"), 0, 0)
        stats_grid.addWidget(self.processed_files_label, 0, 1)
        stats_grid.addWidget(QLabel("Files Remaining:"), 0, 2)
        stats_grid.addWidget(self.remaining_files_label, 0, 3)

        stats_grid.addWidget(QLabel("Data Processed:"), 1, 0)
        stats_grid.addWidget(self.processed_volume_label, 1, 1)
        stats_grid.addWidget(QLabel("Data Remaining:"), 1, 2)
        stats_grid.addWidget(self.remaining_volume_label, 1, 3)

        stats_grid.addWidget(QLabel("Elapsed Time:"), 2, 0)
        stats_grid.addWidget(self.elapsed_time_label, 2, 1)
        stats_grid.addWidget(QLabel("Estimated Time Left:"), 2, 2)
        stats_grid.addWidget(self.remaining_time_label, 2, 3)

        stats_grid.addWidget(QLabel("Total Records Added:"), 3, 0)
        stats_grid.addWidget(self.total_records_label, 3, 1)

        progress_layout.addWidget(self.progress_bar)
        progress_layout.addLayout(stats_grid)

        # Control buttons
        self.scan_button = QPushButton("Start Scanning")
        self.scan_button.setIcon(QIcon.fromTheme("document-save"))
        self.scan_button.setFixedHeight(35)
        font = self.scan_button.font()
        font.setBold(True)
        self.scan_button.setFont(font)

        # Team logo button in bottom-left corner
        self.team_logo_button = QPushButton(self)
        self.team_logo_button.setIcon(QIcon(TEAM_LOGO_PATH))
        self.team_logo_button.setIconSize(QSize(100, 100))  # Set logo size
        self.team_logo_button.setStyleSheet("border: none;")  # Remove button border
        self.team_logo_button.setFixedSize(10, 55)  # Set button size

        self.team_logo_button.clicked.connect(self.show_about_team)

        # Assemble main layout
        main_layout.addLayout(folder_layout)
        main_layout.addWidget(progress_group)
        main_layout.addWidget(self.scan_button)
        main_layout.addWidget(self.team_logo_button, alignment=Qt.AlignmentFlag.AlignRight)

        # Connections
        self.browse_button.clicked.connect(self.browse_folder)
        self.scan_button.clicked.connect(self.start_scanning)

        # Thread
        self.thread = None
        self.worker = None

    def create_stat_label(self):
        label = QLabel("0")
        label.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignBottom)
        label.setStyleSheet("font-weight: bold; color: #88B0FF;")
        return label

    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            self.folder_input.setText(folder)

    def show_about_team(self):
        """Show the About Team dialog."""
        dialog = AboutTeamDialog(self)
        dialog.exec()

    def start_scanning(self):
        if not self.folder_input.text():
            QMessageBox.warning(self, "Warning", "Please select a folder first!")
            return

        # Get output file path
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Report File",
            os.path.join(self.folder_input.text(), "file_columns_report.xlsx"),
            "Excel Files (*.xlsx)"
        )

        if not output_path:
            return  # User canceled

        self.scan_button.setEnabled(False)
        self.browse_button.setEnabled(False)
        self.thread = QThread()
        self.worker = Worker(self.folder_input.text(), output_path)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.finished.connect(self.state_change)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.progress.connect(self.update_progress)
        self.worker.error.connect(self.show_error)

        self.thread.start()

    def state_change(self):
        self.progress_bar.setFormat(f"100% - Process finished.")
        self.browse_button.setEnabled(True)
        self.scan_button.setEnabled(True)

    @pyqtSlot(int, int, float, float, float, float, int)
    def update_progress(self, current, total, processed_size, remaining_size,
                        elapsed, remaining_time, total_records):
        progress = int((current / total) * 100) if total else 0
        self.progress_bar.setValue(progress)
        self.progress_bar.setFormat(f"{progress}% - Processing...")

        self.processed_files_label.setText(f"{current}/{total}")
        self.remaining_files_label.setText(f"{total - current}")

        self.processed_volume_label.setText(f"{processed_size / (1024 ** 3):.2f} GB")
        self.remaining_volume_label.setText(f"{remaining_size / (1024 ** 3):.2f} GB")

        self.elapsed_time_label.setText(time.strftime('%H:%M:%S', time.gmtime(elapsed)))
        self.remaining_time_label.setText(time.strftime('%H:%M:%S', time.gmtime(remaining_time)))
        self.total_records_label.setText(f"{total_records}")  # Update records count

    @pyqtSlot(str)
    def show_error(self, message):
        QMessageBox.critical(self, "Error", message)
        self.scan_button.setEnabled(True)

    def closeEvent(self, event):
        if self.worker:
            self.worker.stop()
        event.accept()


if __name__ == "__main__":
    app = QApplication([])
    app.setStyle("Fusion")
    app.setStyleSheet(STYLE_SHEET)

    # Set default font
    font = QFont()
    font.setFamily("Segoe UI")
    font.setPointSize(10)
    app.setFont(font)

    window = ModernMainWindow()
    window.show()
    app.exec()
