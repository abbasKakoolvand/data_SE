import random
import sys
import os
import json
import time
from datetime import datetime, timedelta
from cryptography.fernet import Fernet
import requests
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QListWidget,
    QMessageBox, QFileDialog, QStackedWidget, QSystemTrayIcon, QMenu, QInputDialog, QDialog
)
from PyQt6.QtCore import pyqtSignal, QTimer, QDate, Qt
from PyQt6.QtGui import QAction, QIcon, QPixmap
from exchangelib import Credentials, Account, DELEGATE, Configuration, Message, Mailbox, FileAttachment
from exchangelib.errors import UnauthorizedError
import hashlib
import base64

# Global variables
account = None
logged_in = False
LOGIN_CHECK_INTERVAL = 30000

# Encryption setup
email = "Abbas.Kakoolvand@gmail.com"
hashed_email = hashlib.sha256(email.encode()).digest()
cipher_key = base64.urlsafe_b64encode(hashed_email)
cipher = Fernet(cipher_key)
from exchangelib import Configuration, NTLM  # Ensure NTLM is imported

class TrayIcon(QSystemTrayIcon):
    global account, logged_in, about_page_is_open
    print(f"account:{account}, logged_in:{logged_in}")

    def __init__(self, parent=None):
        super().__init__(parent)
        self.dialog = None
        self.setIcon(QIcon("mci_mail.png"))
        self.setVisible(True)
        self.email_client = None  # Don't initialize here
        # Create menu
        self.menu = QMenu()
        self.create_actions()
        self.setContextMenu(self.menu)
        # Load config settings
        self.load_config_settings()

    def create_actions(self):
        # Schedule Time
        schedule_action = QAction("Set Schedule Time", self)
        schedule_action.triggered.connect(self.set_schedule_time)
        self.menu.addAction(schedule_action)

        # Process Mail
        process_action = QAction("Process Mail", self)
        process_action.triggered.connect(self.process_mail)
        self.menu.addAction(process_action)

        # Set Credentials
        creds_action = QAction("Set New Credentials", self)
        creds_action.triggered.connect(self.show_login_window)
        self.menu.addAction(creds_action)

        # Set Retry Time
        retry_action = QAction("Set Retry Time", self)
        retry_action.triggered.connect(self.set_retry_time)
        self.menu.addAction(retry_action)

        # Compose Email
        compose_action = QAction("Write Email", self)
        compose_action.triggered.connect(self.show_composer)
        self.menu.addAction(compose_action)

        # View Inbox
        inbox_action = QAction("View Inbox", self)
        inbox_action.triggered.connect(self.show_inbox)
        self.menu.addAction(inbox_action)

        # About
        exit_action = QAction("About", self)
        exit_action.triggered.connect(self.show_about_dialog)
        self.menu.addAction(exit_action)

        # Exit
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.exit_app)
        self.menu.addAction(exit_action)

    def exit_app(self):
        try:
            sys.exit(0)
        except Exception as e:
            print(e)
        try:
            os._exit(0)
        except Exception as e:
            print(e)

    def show_about_dialog(self):
        global about_page_is_open
        if not about_page_is_open:
            self.dialog = AboutTeamDialog()  # Keep reference
            self.dialog.show()
        else:
            # Bring existing dialog to front
            self.dialog.activateWindow()
            self.dialog.raise_()
            self.dialog.setFocus()

    def load_config_settings(self):
        try:
            with open("config.json", "r") as f:
                config = json.load(f)
                self.schedule_hour = config.get("schedule_hour", 8)
                self.retry_interval = config.get("retry_interval", 8 * 3600 * 1000)
        except Exception as e:
            self.schedule_hour = 8
            self.retry_interval = 8 * 3600 * 1000

    def set_schedule_time(self):
        if hasattr(self, "dialog") and (not self.dialog == None):
            if self.dialog.isVisible():
                self.dialog.close()
        try:
            print("Opening dialog for schedule hour...")
            hour, ok = QInputDialog.getInt(None, "Set Schedule Hour",
                                           "Enter hour (0-23):",
                                           self.schedule_hour, 0, 23)
            if ok:
                print(f"User entered hour: {hour}")
                self.schedule_hour = hour
                self.update_config("schedule_hour", hour)
                print("Config updated successfully.")
                self.showMessage("Schedule Updated",
                                 f"Daily tasks scheduled for {hour}:00",
                                 QSystemTrayIcon.MessageIcon.Information)
                print("Message shown successfully.")
            else:
                print("User canceled the dialog.")
        except Exception as e:
            print(f"Error in set_schedule_time: {e}")

    def set_retry_time(self):
        if hasattr(self, "dialog") and (not self.dialog == None):
            if self.dialog.isVisible():
                self.dialog.close()
        hours, ok = QInputDialog.getInt(None, "Set Retry Time",
                                        "Enter retry interval (hours):",
                                        self.retry_interval // 3600000, 1, 24)
        if ok:
            self.retry_interval = hours * 3600 * 1000
            self.update_config("retry_interval", self.retry_interval)
            self.showMessage("Retry Time Updated",
                             f"Retry interval set to {hours} hours",
                             QSystemTrayIcon.MessageIcon.Information)

    def update_config(self, key, value):
        try:
            with open("config.json", "r") as f:
                config = json.load(f)
            config[key] = value
            with open("config.json", "w") as f:
                json.dump(config, f, indent=4)
        except Exception as e:
            print(f"Error updating config: {e}")

    def process_mail(self):
        if hasattr(self, "dialog") and (not self.dialog == None):
            if self.dialog.isVisible():
                self.dialog.close()
        global account, logged_in
        print(f"account:{account}, logged_in:{logged_in}")
        if not account or not logged_in:
            self.showMessage("Error", "Not logged in!", QSystemTrayIcon.MessageIcon.Critical)
            return
        if not self.email_client:
            self.email_client = EmailClientHandler(self)
        try:
            self.email_client.process_attachments()
            self.email_client.upload_attachments()
            self.showMessage("Processing Complete",
                             "Mail processing finished",
                             QSystemTrayIcon.MessageIcon.Information)
        except Exception as e:
            self.showMessage("Error", str(e),
                             QSystemTrayIcon.MessageIcon.Critical)
            print(f"Error processing mail: {e}")

    def show_login_window(self):
        if hasattr(self, "dialog") and (not self.dialog == None):
            if self.dialog.isVisible():
                self.dialog.close()
        self.login_window = LoginWindow(self)
        self.login_window.login_success.connect(self.on_login_success)
        self.login_window.show()

    def show_composer(self):
        if hasattr(self, "dialog") and (not self.dialog == None):
            if self.dialog.isVisible():
                self.dialog.close()
        global account, logged_in
        print(f"account:{account}, logged_in:{logged_in}")
        if not account or not logged_in:
            self.showMessage("Error", "Not logged in!", QSystemTrayIcon.MessageIcon.Critical)
            return
        self.composer = EmailComposer(self)
        self.composer.show()

    def show_inbox(self):
        if hasattr(self, "dialog") and (not self.dialog == None):
            if self.dialog.isVisible():
                self.dialog.close()
        global account, logged_in
        print(f"account:{account}, logged_in:{logged_in}")
        if not account or not logged_in:
            self.showMessage("Error", "Not logged in!", QSystemTrayIcon.MessageIcon.Critical)
            return
        self.inbox = InboxViewer(self)
        self.inbox.show()

    def on_login_success(self):
        global account, logged_in

        logged_in = True
        print(f"account:{account}, logged_in:{logged_in}")
        self.showMessage("Login Success",
                         "Credentials updated successfully",
                         QSystemTrayIcon.MessageIcon.Information)


class LoginWindow(QMainWindow):
    global account, logged_in
    login_success = pyqtSignal()

    def __init__(self, tray_icon):
        super().__init__()
        self.tray_icon = tray_icon
        self.setWindowTitle(f"{APP_NAME} - Login")
        self.setWindowIcon(QIcon(APP_LOGO_PATH))
        self.setGeometry(100, 100, 400, 250)
        self.otp_required = False
        self.retry_timer = QTimer()
        self.init_ui()

    def init_ui(self):
        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)

        # Basic login page
        self.basic_login_widget = QWidget()
        basic_layout = QVBoxLayout()
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("Email")
        basic_layout.addWidget(self.email_input)
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Password")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        basic_layout.addWidget(self.password_input)
        self.login_btn = QPushButton("Login")
        self.login_btn.clicked.connect(self.attempt_basic_login)
        basic_layout.addWidget(self.login_btn)
        self.basic_login_widget.setLayout(basic_layout)
        self.stacked_widget.addWidget(self.basic_login_widget)

        # OTP login page
        self.otp_login_widget = QWidget()
        otp_layout = QVBoxLayout()
        self.otp_label = QLabel("Enter OTP:")
        self.otp_input = QLineEdit()
        otp_layout.addWidget(self.otp_label)
        otp_layout.addWidget(self.otp_input)
        self.otp_login_btn = QPushButton("Complete Login")
        self.otp_login_btn.clicked.connect(self.attempt_otp_login)
        otp_layout.addWidget(self.otp_login_btn)
        self.otp_login_widget.setLayout(otp_layout)
        self.stacked_widget.addWidget(self.otp_login_widget)

    def load_encrypted_credentials(self):
        try:
            with open("config.json", "r") as f:
                config = json.load(f)
                if 'encrypted_user' in config and 'encrypted_pass' in config:
                    email_user = cipher.decrypt(config['encrypted_user'].encode()).decode()
                    password = cipher.decrypt(config['encrypted_pass'].encode()).decode()
                    self.email_input.setText(email_user)
                    self.password_input.setText(password)
                    return True
        except Exception as e:
            print(f"Error loading credentials: {e}")
            self.tray_icon.showMessage("Credentials Error",
                                       f"Failed to load credentials: {e}",
                                       QSystemTrayIcon.MessageIcon.Warning)
        return False

    def attempt_basic_login(self):
        global account, logged_in
        email = self.email_input.text()
        password = self.password_input.text()

        try:
            # Split email to get domain and username
            domain_part = email.split('@')[1].split('.')[0].upper()  # Extracts "MCI" from "a.kakoolvand@mci.ir"
            username = f"{domain_part}\\{email.split('@')[0]}"  # Format as "MCI\a.kakoolvand"

            # Use email as primary_smtp_address, username for authentication
            credentials = Credentials(username=username, password=password)

            config = Configuration(
                credentials=credentials,
                service_endpoint="https://mail.mci.ir/ews/exchange.asmx",
                auth_type=NTLM  # Force NTLM authentication
            )

            account = Account(
                primary_smtp_address=email,  # Keep the email address here
                config=config,
                autodiscover=False,
                access_type=DELEGATE
            )

            # Disable SSL verification
            session = requests.Session()
            session.verify = False
            account.protocol.session = session  # Apply session before any requests

            account.root.refresh()  # Test connection

            # Success handling
            self.login_success.emit()
            logged_in = True
            self.save_encrypted_credentials(email, password)
            self.close()
            self.tray_icon.showMessage("Login Success", "Connected to email server",
                                       QSystemTrayIcon.MessageIcon.Information)

        except UnauthorizedError as e:
            error_msg = f"Authorization Error: {str(e)}"
            print(error_msg)
            self.tray_icon.showMessage("Login Failed", error_msg, QSystemTrayIcon.MessageIcon.Critical)
            if "multi-factor" in str(e).lower():
                self.stacked_widget.setCurrentWidget(self.otp_login_widget)
        except Exception as e:
            error_msg = f"Connection Error: {str(e)}"
            print(error_msg)
            self.tray_icon.showMessage("Connection Error", error_msg, QSystemTrayIcon.MessageIcon.Critical)
            self.handle_connection_error(self.attempt_basic_login)
    def attempt_otp_login(self):
        global account, logged_in
        email = self.email_input.text()
        password = self.password_input.text()
        otp = self.otp_input.text()
        try:
            credentials = Credentials(username=email, password=f"{password}{otp}")
            config = Configuration(
                credentials=credentials,
                service_endpoint="https://mail.mci.ir/ews/exchange.asmx"
            )
            account = Account(
                primary_smtp_address=email,
                config=config,
                autodiscover=False,
                access_type=DELEGATE
            )
            account.root.refresh()
            self.save_encrypted_credentials(email, f"{password}{otp}")
            self.login_success.emit()
            self.close()
            self.tray_icon.showMessage("OTP Login Success",
                                       "Multi-factor authentication completed",
                                       QSystemTrayIcon.MessageIcon.Information)
        except Exception as e:
            error_msg = f"OTP Error: {str(e)}"
            print(error_msg)
            self.tray_icon.showMessage("OTP Failed", error_msg,
                                       QSystemTrayIcon.MessageIcon.Critical)

    def save_encrypted_credentials(self, email, password):
        """New: Save encrypted credentials"""
        try:
            encrypted_user = cipher.encrypt(email.encode()).decode()
            encrypted_pass = cipher.encrypt(password.encode()).decode()
            with open("config.json", "r") as f:
                config = json.load(f)
            config['encrypted_user'] = encrypted_user
            config['encrypted_pass'] = encrypted_pass
            with open("config.json", 'w') as f:
                json.dump(config, f, indent=4)
        except Exception as e:
            print(f"Error saving credentials: {e}")

    def handle_connection_error(self, retry_action):
        self.tray_icon.showMessage("Connection Error",
                                   "Will retry in configured interval",
                                   QSystemTrayIcon.MessageIcon.Warning)
        self.retry_timer.timeout.connect(retry_action)
        self.retry_timer.start(self.tray_icon.retry_interval)


TEAM_LOGO_PATH = "team_logo.png"  # Path to your team logo (PNG format recommended)
APP_LOGO_PATH = "mci_mail.png"  # Path to your team logo (PNG format recommended)
APP_NAME = "MCI Mail Attachment Aggregator"
TEAM_NAME = "SPM BI"
TEAM_DESCRIPTION = """
CSO
MCI

"""
about_page_is_open = False


class AboutTeamDialog(QDialog):
    global about_page_is_open

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"{APP_NAME} - About")
        self.setWindowIcon(QIcon(APP_LOGO_PATH))  # Use app logo
        self.setModal(True)
        self.setMaximumHeight(400)
        self.setMaximumHeight(400)
        self.setMaximumWidth(350)
        self.setMinimumWidth(350)

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

    def showEvent(self, event):
        global about_page_is_open
        """Set flag to True when the dialog is shown."""
        about_page_is_open = True
        super().showEvent(event)

    def closeEvent(self, event):
        global about_page_is_open
        """Set flag to False when the dialog is closed."""
        about_page_is_open = False
        super().closeEvent(event)

    def mouseMoveEvent(self, event):
        # Override to do nothing on mouse move
        pass


class EmailClientHandler:
    global account, logged_in
    def __init__(self, tray_icon):
        self.tray_icon = tray_icon
        self.progress_dialog = None  # Progress dialog instance
        self.total_emails = 0
        self.processed_emails = 0
        self.total_attachments_found = 0
        self.total_attachments_saved = 0
        self.total_files_uploaded = 0
        self.last_daily_run = None
        self.daily_timer = QTimer()
        self.daily_timer.timeout.connect(self.check_daily_task)
        self.daily_timer.start(random.randint(50, 55) * 60 * 1000)
        self.start_time = None  # Track when processing starts
        self.total_emails = 0  # Track total emails in the current run
        self.session_check_timer = QTimer()
        self.session_check_timer.timeout.connect(self.is_account_valid)
        self.session_check_timer.start(3600000)  # Check every hour

    def process_initial_run(self):
        global account
        if account:  # Only run if account is initialized
            self.process_attachments()
            time.sleep(60)
            self.upload_attachments()

    def is_account_valid(self):
        global account, logged_in
        try:
            if not account or not logged_in:
                return False
            inbox = list(account.inbox.all().order_by('-datetime_received')[:10])  # Validate connection
            return True
        except Exception:
            return False

    def handle_reauthentication(self):
        try:
            # Try re-login with stored credentials

            self.login_window = LoginWindow(self.tray_icon)
            self.login_window.load_encrypted_credentials()
            self.login_window.attempt_basic_login()
        except Exception as e:
            self.tray_icon.showMessage(
                "Re-authentication Failed",
                f"Stored credentials invalid. Please check error.{e}",
                QSystemTrayIcon.MessageIcon.Critical
            )

    def check_daily_task(self):
        global account, logged_in
        now = datetime.now()
        current_date = QDate.currentDate()
        self.is_account_valid()
        if now.hour == self.tray_icon.schedule_hour and 0 <= now.minute < 60:
            if not self.is_account_valid():
                self.handle_reauthentication()
                return
            if self.last_daily_run == current_date:
                return
            self.last_daily_run = current_date
            self.process_attachments()
            time.sleep(60)
            self.upload_attachments()

    def process_attachments(self):
        global account, logged_in
        try:
            if not account or not logged_in:
                raise ValueError("Not authenticated with email server")
            # Rest of the original process_attachments code
            with open("config.json", "r") as f:
                config = json.load(f)
            start_date_str = config.get("start_date")
            destination_folder = config.get("destination_folder")
            destination_folder = os.path.join(os.getcwd(), destination_folder)
            if not start_date_str or not destination_folder:
                raise ValueError("Missing config parameters")
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            if not os.path.exists(destination_folder):
                os.makedirs(destination_folder)
            all_emails = list(account.inbox.all().order_by('datetime_received'))
            self.emails_to_process = [
                email for email in all_emails
                if email.datetime_received.replace(tzinfo=None) > start_date
            ]
            self.emails_to_process.sort(key=lambda email: email.datetime_received)
            self.destination_folder = destination_folder
            # Initialize logging variables
            self.start_time = datetime.now()
            self.total_emails = len(self.emails_to_process)
            self.log_run(type_process="Save Attachment")
            self.progress_dialog = ProgressDialog()
            self.progress_dialog.set_phase("Processing Emails")
            self.progress_dialog.set_total_emails(len(self.emails_to_process))
            self.progress_dialog.show()

            # Start processing
            self.process_next_email()
        except Exception as e:
            self.progress_dialog.close()
            error_msg = f"Processing Error: {str(e)}"
            print(error_msg)
            self.tray_icon.showMessage("Processing Failed", error_msg,
                                       QSystemTrayIcon.MessageIcon.Critical)
            # Log the error
            self.log_run(type_process="Save Attachment", error=error_msg)

    def process_next_email(self):
        if not self.emails_to_process:
            self.tray_icon.showMessage("Processing Complete",
                                       "All emails attachments were saved.",
                                       QSystemTrayIcon.MessageIcon.Information)
            return
        email = self.emails_to_process.pop(0)
        self.processed_emails += 1
        try:
            # Process attachments
            for attachment in email.attachments:
                self.total_attachments_found += 1
                if isinstance(attachment, FileAttachment) and attachment.name.lower().endswith(
                        ('.xlsx', '.xls', '.csv')):
                    file_path = os.path.join(self.destination_folder, attachment.name)
                    with open(file_path, 'wb') as f:
                        f.write(attachment.content)
                    self.total_attachments_saved += 1

            # Update progress dialog
            self.progress_dialog.update_email_progress(self.processed_emails)
            self.progress_dialog.update_attachment_counts(
                self.total_attachments_found,
                self.total_attachments_saved
            )
            time.sleep(5)
            self.process_next_email()
        except Exception as e:
            self.progress_dialog.set_phase(f"Error: {str(e)}")
            error_msg = f"Attachment Error: {str(e)}"
            print(error_msg)
            self.tray_icon.showMessage("Attachment Failed", error_msg,
                                       QSystemTrayIcon.MessageIcon.Warning)

    def upload_attachments(self):
        upload_Done = True
        try:
            with open("config.json", "r") as f:
                config = json.load(f)

            destination_folder = config.get("destination_folder")
            destination_folder = os.path.join(os.getcwd(), destination_folder)
            files = os.listdir(destination_folder)
            total_files = len(files)
            self.progress_dialog.set_phase("Uploading Files")
            self.progress_dialog.set_upload_total(total_files)
            for i, file in enumerate(files, 1):
                full_path = os.path.join(destination_folder, file)
                if True:
                    url = "https://bi.mci.ir/myflask/upload"

                    # Use the full path to the file
                    with open(full_path, 'rb') as f:
                        response = requests.post(url, files={'file': f}, verify=False)

                    print(response.status_code)
                    print(response.json())
                    if response.status_code != 200:
                        upload_Done = False
                        error_msg_upload = response.text
                        print(error_msg_upload)

                    time.sleep(15)
                    os.remove(full_path)
                    self.total_files_uploaded += 1
                    self.progress_dialog.update_upload_progress(i)
            if upload_Done:
                update_time_config()

                self.tray_icon.showMessage("Upload Complete",
                                           "Files uploaded successfully",
                                           QSystemTrayIcon.MessageIcon.Information)
                self.log_run(type_process="Upload")
            else:

                self.tray_icon.showMessage("Upload Failed", error_msg_upload,
                                           QSystemTrayIcon.MessageIcon.Critical)
                self.log_run(type_process="Upload", error=error_msg_upload)
            clear_directory_contents(destination_folder)
            self.progress_dialog.close()

        except Exception as e:
            error_msg = f"Upload Error: {str(e)}"
            print(error_msg)
            self.tray_icon.showMessage("Upload Failed", error_msg,
                                       QSystemTrayIcon.MessageIcon.Critical)
            self.log_run(type_process="Upload", error=error_msg)
            try:
                self.progress_dialog.close()
            except:
                pass

    def log_run(self, type_process=None, error=None):
        """Logs the runtime and number of emails processed"""
        if type_process == "Save Attachment":
            count = self.total_emails
        else:
            count = self.total_attachments_saved
        log_entry = (
            f"{type_process} - {self.start_time.strftime('%Y-%m-%d %H:%M:%S')} - "
            f"Processed {count} (emails/files)"
        )
        if error:
            log_entry += f"{type_process} - Error: {error}"
        log_entry += "\n"

        try:
            with open("process_log.txt", "a") as log_file:
                log_file.write(log_entry)
        except Exception as e:
            print(f"Failed to write log: {e}")


class EmailComposer(QMainWindow):
    global account, logged_in

    def __init__(self, tray_icon):
        super().__init__()
        self.tray_icon = tray_icon
        global account, logged_in
        if not account or not logged_in:
            self.tray_icon.showMessage("Error", "Not logged in!", QSystemTrayIcon.MessageIcon.Critical)
            return
        self.setWindowTitle("Compose Email")
        self.setGeometry(200, 200, 600, 400)
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        layout = QVBoxLayout()
        self.to_input = QLineEdit()
        self.to_input.setPlaceholderText("To")
        layout.addWidget(self.to_input)
        self.subject_input = QLineEdit()
        self.subject_input.setPlaceholderText("Subject")
        layout.addWidget(self.subject_input)
        self.body_input = QTextEdit()
        self.body_input.setPlaceholderText("Message body")
        layout.addWidget(self.body_input)
        self.attachment_btn = QPushButton("Add Attachment")
        self.attachment_btn.clicked.connect(self.add_attachment)
        layout.addWidget(self.attachment_btn)
        self.send_btn = QPushButton("Send")
        self.send_btn.clicked.connect(self.send_email)
        layout.addWidget(self.send_btn)
        self.attachments = []
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def add_attachment(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Attachment")
        if file_path:
            self.attachments.append(file_path)
            print(f"Attachment added: {os.path.basename(file_path)}")
            self.tray_icon.showMessage("Attachment Added",
                                       f"Added attachment: {os.path.basename(file_path)}",
                                       QSystemTrayIcon.MessageIcon.Information)

    def send_email(self):
        global account
        try:
            msg = Message(
                account=account,
                subject=self.subject_input.text(),
                body=self.body_input.toPlainText(),
                to_recipients=[Mailbox(email_address=self.to_input.text())]
            )
            for file_path in self.attachments:
                with open(file_path, 'rb') as f:
                    attachment = FileAttachment(
                        name=os.path.basename(file_path),
                        content=f.read()
                    )
                msg.attach(attachment)
                print(f"Attached file: {os.path.basename(file_path)}")
            msg.send()
            print("Email sent successfully!")
            self.tray_icon.showMessage("Email Sent",
                                       "Message delivered successfully",
                                       QSystemTrayIcon.MessageIcon.Information)
        except Exception as e:
            error_msg = f"Sending Error: {str(e)}"
            print(error_msg)
            self.tray_icon.showMessage("Send Failed", error_msg,
                                       QSystemTrayIcon.MessageIcon.Critical)


from PyQt6.QtWidgets import QDialog, QVBoxLayout, QLabel, QProgressBar


class ProgressDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Processing Progress")
        self.setModal(True)  # Make it modal to block other interactions
        self.layout = QVBoxLayout()

        # Phase label (e.g., "Processing emails", "Uploading files")
        self.phase_label = QLabel("Initializing...", self)
        self.layout.addWidget(self.phase_label)

        # Progress bar
        self.progress_bar = QProgressBar(self)
        self.layout.addWidget(self.progress_bar)

        # Status labels
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


class InboxViewer(QMainWindow):
    global account, logged_in

    def __init__(self, tray_icon):
        super().__init__()
        self.tray_icon = tray_icon
        global account, logged_in
        if not account or not logged_in:
            self.tray_icon.showMessage("Error", "Not logged in!", QSystemTrayIcon.MessageIcon.Critical)
            return
        self.setWindowTitle("Inbox")
        self.setGeometry(200, 200, 800, 600)
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        layout = QHBoxLayout()
        # Email list
        self.email_list = QListWidget()
        self.email_list.itemClicked.connect(self.show_email)
        layout.addWidget(self.email_list, 1)
        # Email viewer
        self.email_viewer = QTextEdit()
        self.email_viewer.setReadOnly(True)
        layout.addWidget(self.email_viewer, 2)
        self.load_emails()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def load_emails(self):
        global account, logged_in
        print("Loading emails for Inbox Viewer...")
        try:
            inbox = list(account.inbox.all().order_by('-datetime_received')[:10])
            for email in inbox:
                self.email_list.addItem(f"{email.subject} - {email.sender.email_address}")
            self.emails = inbox
            print(f"Loaded {len(inbox)} email(s) into the inbox list.")
        except Exception as e:
            print(f"Error loading emails: {str(e)}")
            self.tray_icon.showMessage("Error", f"Error loading emails: {str(e)}", QSystemTrayIcon.MessageIcon.Critical)

    def show_email(self, item):
        global account, logged_in
        try:
            index = self.email_list.row(item)
            email = self.emails[index]
            print(f"Displaying email: {email.subject}")
            content = f"From: {email.sender.email_address}\n"
            content += f"Date: {email.datetime_received}\n"
            content += email.text_body
            self.email_viewer.setText(content)
        except Exception as e:
            print(f"Error displaying email: {str(e)}")
            self.tray_icon.showMessage("Error", f"Failed to display email: {str(e)}",
                                       QSystemTrayIcon.MessageIcon.Critical)


def update_time_config():
    # Define the path to your JSON file
    with open("config.json", "r") as f:
        data = json.load(f)
    # Get today's date and subtract one day
    yesterday = datetime.now() - timedelta(days=1)
    new_start_date = yesterday.strftime('%Y-%m-%d')
    # Update the start_date in the JSON data
    data['start_date'] = new_start_date
    # Save the updated data back to the JSON file
    with open("config.json", 'w') as file:
        json.dump(data, file, indent=4)
    print(f"Updated start_date to {new_start_date} in config.json.")


import os


def clear_directory_contents(target_dir):
    """Remove all files and subdirectories within a directory, but keep the directory itself."""
    for root, dirs, files in os.walk(target_dir, topdown=False):
        # Remove all files
        for file in files:
            file_path = os.path.join(root, file)
            os.remove(file_path)
            print(f"Deleted file: {file_path}")

        # Remove subdirectories (except the target directory itself)
        if root != target_dir:
            os.rmdir(root)
            print(f"Deleted directory: {root}")


if __name__ == "__main__":
    print(f"account:{account}, logged_in:{logged_in}")

    # Initialize the application
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(APP_LOGO_PATH))

    # Check if system tray is available
    if not QSystemTrayIcon.isSystemTrayAvailable():
        print("System tray not available!")
        sys.exit(1)  # Exit only if the system tray is unavailable

    # Create tray icon and login window
    tray = TrayIcon()
    login_window = LoginWindow(tray)

    # Connect login success signal to initialize email client
    login_window.login_success.connect(lambda: setattr(tray, 'email_client', EmailClientHandler(tray)))

    # Attempt to load encrypted credentials and perform auto-login
    if login_window.load_encrypted_credentials():
        try:
            login_window.attempt_basic_login()
        except Exception as e:
            print(f"Auto-login failed: {e}")
            # Do not exit the application on login failure
            tray.showMessage("Login Failed", f"Auto-login failed: {e}", QSystemTrayIcon.MessageIcon.Critical)

    # Start the main event loop
    while True:
        try:
            app.exec()  # Keep the application running
        except Exception as e:
            print(f"Critical error: {e}")
            tray.showMessage("Critical Error", str(e), QSystemTrayIcon.MessageIcon.Critical)
# pyinstaller --windowed --icon=mci_mail.ico --name "MCI Mail Attachment Aggregator" --add-data "mci_mail.png;." --add-data "team_logo.png;." mail_gui_v5.py --noconfirm
