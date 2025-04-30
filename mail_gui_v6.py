import os
import sys
import json
import time
import random
from datetime import datetime, timedelta
import requests
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QTextEdit, QListWidget, QMessageBox, QFileDialog,
    QStackedWidget, QSystemTrayIcon, QMenu, QDialog
)
from PyQt6.QtCore import pyqtSignal, QTimer, Qt
from PyQt6.QtGui import QIcon, QPixmap
import pythoncom
import win32com.client

# Constants
APP_NAME = "Outlook Attachment Processor"
APP_LOGO_PATH = "mci_mail.png"
TEAM_LOGO_PATH = "team_logo.png"
CONFIG_FILE = "config.json"
with open(CONFIG_FILE, "r") as f:
    config = json.load(f)
destination_folder = config.get("destination_folder")
SAVE_PATH = os.path.join(os.getcwd(), destination_folder)


# Global COM state
COM_INITIALIZED = False


def ensure_com_initialized():
    global COM_INITIALIZED
    if not COM_INITIALIZED:
        pythoncom.CoInitialize()
        COM_INITIALIZED = True


class OutlookClientHandler:
    def __init__(self, tray_icon):
        self.tray_icon = tray_icon
        self.progress_dialog = None
        self.load_config()

        # Initialize Outlook
        self.outlook = None
        self.inbox = None
        self.connect_to_outlook()

        # Schedule daily task
        self.daily_timer = QTimer()
        self.daily_timer.timeout.connect(self.check_daily_task)
        self.daily_timer.start(random.randint(50, 55) * 60 * 1000)

    def connect_to_outlook(self):
        try:
            ensure_com_initialized()
            self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            self.inbox = self.outlook.GetDefaultFolder(6)  # Inbox
        except Exception as e:
            self.show_error("Outlook Connection Failed", str(e))

    def load_config(self):
        try:
            with open(CONFIG_FILE, "r") as f:
                config = json.load(f)
                self.schedule_hour = config.get("schedule_hour", 8)
                self.start_date = config.get("start_date", "2025-01-01")
        except Exception:
            self.schedule_hour = 8
            self.start_date = "2025-01-01"

    def show_error(self, title, message):
        self.tray_icon.showMessage(title, message, QSystemTrayIcon.MessageIcon.Critical)
        print(f"{title}: {message}")

    def check_daily_task(self):
        now = datetime.now()
        if now.hour == self.schedule_hour and now.minute < 5:
            self.process_attachments()
            time.sleep(60)
            self.upload_attachments()

    def process_attachments(self):
        try:
            os.makedirs(SAVE_PATH, exist_ok=True)

            # Format date for DASL query
            date_str = self.start_date
            query = f"@SQL=(\"urn:schemas:httpmail:datereceived\" > '{date_str}')"
            messages = self.inbox.Items.Restrict(query)

            if messages.Count == 0:
                self.show_info("No Emails", "No emails found after the specified date.")
                return

            # Show progress dialog
            self.progress_dialog = ProgressDialog()
            self.progress_dialog.set_phase("Processing Emails")
            self.progress_dialog.set_total_emails(messages.Count)
            self.progress_dialog.show()

            # Process each email
            processed = 0
            total_attachments = 0
            saved_attachments = 0

            for i, msg in enumerate(messages, 1):
                try:
                    print(f"[{i}] Processing email: '{msg.Subject}' from {msg.SenderName}")
                    attachments = msg.Attachments

                    for j, attachment in enumerate(attachments, 1):
                        filename = attachment.FileName
                        if filename.lower().endswith(('.xlsx', '.xls', '.csv')):
                            file_path = os.path.join(SAVE_PATH, filename)
                            attachment.SaveAsFile(file_path)
                            saved_attachments += 1
                            self.progress_dialog.update_attachment_counts(total_attachments, saved_attachments)
                            print(f"  [{j}] Saved attachment: {filename}")

                    processed += 1
                    self.progress_dialog.update_email_progress(processed)
                except Exception as e:
                    print(f"Error processing email: {e}")

            self.progress_dialog.close()
            self.show_info("Processing Complete", f"Saved {saved_attachments} attachments.")

            # Update config date
            self.update_start_date()

        except Exception as e:
            self.show_error("Processing Failed", str(e))
            self.progress_dialog.close()

    def update_start_date(self):
        try:
            yesterday = datetime.now() - timedelta(days=1)
            new_date = yesterday.strftime("%Y-%m-%d")

            with open(CONFIG_FILE, "r") as f:
                config = json.load(f)
                config["start_date"] = new_date

            with open(CONFIG_FILE, "w") as f:
                json.dump(config, f, indent=2)

        except Exception as e:
            print(f"Failed to update start date: {e}")

    def upload_attachments(self):

        try:
            files = os.listdir(SAVE_PATH)
            if not files:
                self.show_info("No Files", "No files to upload.")
                return

            self.progress_dialog = ProgressDialog()
            self.progress_dialog.set_phase("Uploading Files")
            self.progress_dialog.set_upload_total(len(files))
            self.progress_dialog.show()

            uploaded = 0
            upload_done = True

            for file in files:
                full_path = os.path.join(SAVE_PATH, file)
                try:
                    with open(full_path, 'rb') as f:
                        response = requests.post(
                            "https://bi.mci.ir/myflask/upload",
                            files={'file': f},
                            verify=False
                        )
                        print(f"Uploaded: {file}, Status: {response.status_code}")
                        if response.status_code != 200:
                            upload_done = False

                    os.remove(full_path)
                    uploaded += 1
                    self.progress_dialog.update_upload_progress(uploaded)
                    time.sleep(1)

                except Exception as e:
                    print(f"Upload error for {file}: {e}")
                    upload_done = False

            self.progress_dialog.close()
            if upload_done:
                self.show_info("Upload Success", "All files uploaded successfully.")
            else:
                self.show_error("Upload Failed", "Some files failed to upload.")

        except Exception as e:
            self.show_error("Upload Error", str(e))
            self.progress_dialog.close()

    def show_info(self, title, message):
        self.tray_icon.showMessage(title, message, QSystemTrayIcon.MessageIcon.Information)
        print(f"{title}: {message}")


class ProgressDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Processing Progress")
        self.setModal(True)
        self.layout = QVBoxLayout()

        self.phase_label = QLabel("Initializing...", self)
        self.layout.addWidget(self.phase_label)

        self.progress_bar = QProgressBar(self)
        self.layout.addWidget(self.progress_bar)

        self.email_label = QLabel("Emails processed: 0/0", self)
        self.attachment_found_label = QLabel("Attachments found: 0", self)
        self.attachment_saved_label = QLabel("Attachments saved: 0", self)
        self.uploaded_label = QLabel("Files uploaded: 0/0", self)

        self.layout.addWidget(self.email_label)
        self.layout.addWidget(self.attachment_found_label)
        self.layout.addWidget(self.attachment_saved_label)
        self.layout.addWidget(self.uploaded_label)

        self.setLayout(self.layout)

    def set_phase(self, phase_text):
        self.phase_label.setText(phase_text)

    def set_total_emails(self, total):
        self.progress_bar.setMaximum(total)
        self.email_label.setText(f"Emails processed: 0/{total}")

    def update_email_progress(self, processed):
        self.progress_bar.setValue(processed)
        self.email_label.setText(f"Emails processed: {processed}/{self.progress_bar.maximum()}")

    def update_attachment_counts(self, found, saved):
        self.attachment_found_label.setText(f"Attachments found: {found}")
        self.attachment_saved_label.setText(f"Attachments saved: {saved}")

    def set_upload_total(self, total):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(0)
        self.uploaded_label.setText(f"Files uploaded: 0/{total}")

    def update_upload_progress(self, uploaded):
        self.progress_bar.setValue(uploaded)
        self.uploaded_label.setText(f"Files uploaded: {uploaded}/{self.progress_bar.maximum()}")


class TrayIcon(QSystemTrayIcon):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setIcon(QIcon(APP_LOGO_PATH))
        self.setVisible(True)
        self.email_client = OutlookClientHandler(self)
        self.create_menu()

    def create_menu(self):
        menu = QMenu()
        schedule_action = QAction("Set Schedule Time", self)
        schedule_action.triggered.connect(self.set_schedule_time)
        menu.addAction(schedule_action)

        process_action = QAction("Process Mail", self)
        process_action.triggered.connect(self.process_mail)
        menu.addAction(process_action)

        upload_action = QAction("Upload Attachments", self)
        upload_action.triggered.connect(self.upload_attachments)
        menu.addAction(upload_action)

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(sys.exit)
        menu.addAction(exit_action)

        self.setContextMenu(menu)

    def set_schedule_time(self):
        hour, ok = QInputDialog.getInt(
            None, "Set Schedule Hour", "Enter hour (0-23):",
            self.email_client.schedule_hour, 0, 23
        )
        if ok:
            self.email_client.schedule_hour = hour
            with open(CONFIG_FILE, "r+") as f:
                config = json.load(f)
                config["schedule_hour"] = hour
                f.seek(0)
                json.dump(config, f, indent=2)
                f.truncate()

    def process_mail(self):
        self.email_client.process_attachments()

    def upload_attachments(self):
        self.email_client.upload_attachments()


class AboutTeamDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"{APP_NAME} - About")
        self.setModal(True)
        layout = QVBoxLayout()

        logo_label = QLabel(self)
        logo_pixmap = QPixmap(TEAM_LOGO_PATH)
        logo_label.setPixmap(logo_pixmap.scaled(200, 200, Qt.AspectRatioMode.KeepAspectRatio))
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(logo_label)

        app_name = QLabel(f"{APP_NAME}\n\nA product of SPM BI Team")
        app_name.setAlignment(Qt.AlignmentFlag.AlignCenter)
        app_name.setStyleSheet("font-size: 16px;")
        layout.addWidget(app_name)

        self.setLayout(layout)


if __name__ == "__main__":
    pythoncom.CoInitialize()

    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(APP_LOGO_PATH))

    if not QSystemTrayIcon.isSystemTrayAvailable():
        QMessageBox.critical(None, "Error", "System tray not supported.")
        sys.exit(1)

    tray_icon = TrayIcon()
    tray_icon.show()

    try:
        sys.exit(app.exec())
    finally:
        pythoncom.CoUninitialize()
