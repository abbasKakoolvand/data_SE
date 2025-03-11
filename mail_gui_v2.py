import sys
import os
import json
import requests
from datetime import datetime, timedelta
from tempfile import gettempdir
import shutil
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QListWidget,
    QMessageBox, QFileDialog, QStackedWidget, QSystemTrayIcon, QMenu, QCheckBox
)
from PyQt6.QtCore import Qt, pyqtSignal, QTimer, QDate
from PyQt6.QtGui import QIcon, QAction
from exchangelib import Credentials, Account, DELEGATE, Configuration, Message, Mailbox, FileAttachment
from exchangelib.errors import UnauthorizedError
from cryptography.fernet import Fernet

# Configuration
CONFIG_FILE = "email_client_config.json"
TEMP_FOLDER = os.path.join(gettempdir(), "email_attachments")
ALLOWED_EXTENSIONS = {'.xlsx', '.xls', '.csv', '.xlsm', '.xlsb'}
UPLOAD_URL = "https://bi.mci.ir/myflask/upload"

# Add near the top of your code
KEY_FILE = "encryption.key"


# Generate or load encryption key
def get_or_create_key():
    if not os.path.exists(KEY_FILE):
        key = Fernet.generate_key()
        with open(KEY_FILE, 'wb') as key_file:
            key_file.write(key)
    with open(KEY_FILE, 'rb') as key_file:
        return key_file.read()


KEY = get_or_create_key()
cipher_suite = Fernet(KEY)


class SecureCredentials:
    @staticmethod
    def save_credentials(email, password):
        try:
            encrypted_email = cipher_suite.encrypt(email.encode())
            encrypted_password = cipher_suite.encrypt(password.encode())
            with open(CONFIG_FILE, 'w') as f:
                json.dump({
                    'email': encrypted_email.decode(),
                    'password': encrypted_password.decode()
                }, f)
        except Exception as e:
            QMessageBox.critical(None, "Error", f"Failed to save credentials: {str(e)}")
            SecureCredentials.clear_credentials()

    @staticmethod
    def load_credentials():
        try:
            if not os.path.exists(CONFIG_FILE):
                return None, None

            with open(CONFIG_FILE, 'r') as f:
                data = json.load(f)
                return (
                    cipher_suite.decrypt(data['email'].encode()).decode(),
                    cipher_suite.decrypt(data['password'].encode()).decode()
                )
        except Exception as e:
            QMessageBox.critical(None, "Error",
                                 "Failed to load credentials. Please log in again.")
            SecureCredentials.clear_credentials()
            return None, None


class EmailClient(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Email Client")
        self.setGeometry(100, 100, 800, 600)
        self.tray_icon = None
        self.last_run_date = None
        self.init_ui()  # Now properly defined
        self.init_tray()
        self.load_last_run_date()
        self.setup_daily_task()

    def init_ui(self):
        central_widget = QWidget()
        layout = QVBoxLayout()

        # Status Label
        self.status_label = QLabel("Service is running")
        layout.addWidget(self.status_label)

        # Buttons
        btn_layout = QHBoxLayout()
        self.send_btn = QPushButton("Send Email")
        self.send_btn.clicked.connect(self.open_composer)
        btn_layout.addWidget(self.send_btn)

        self.inbox_btn = QPushButton("Check Inbox")
        self.inbox_btn.clicked.connect(self.open_inbox)
        btn_layout.addWidget(self.inbox_btn)

        layout.addLayout(btn_layout)
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def open_composer(self):
        self.composer = EmailComposer()
        self.composer.show()

    def open_inbox(self):
        self.inbox = InboxViewer()
        self.inbox.show()

    def init_tray(self):
        self.tray_icon = QSystemTrayIcon(QIcon("icon.ico"), self)
        menu = QMenu()

        self.show_action = QAction("Show", self)
        self.show_action.triggered.connect(self.change_current_situation)
        menu.addAction(self.show_action)

        self.exit_action = QAction("Exit", self)
        self.exit_action.triggered.connect(self.close_completely)
        menu.addAction(self.exit_action)

        self.tray_icon.setContextMenu(menu)
        self.tray_icon.show()

    def close_completely(self):
        self.close()
        try:
            sys.exit(0)  # Graceful exit
        except Exception as e:
            os._exit(0)
        finally:
            os._exit(0)

    def change_current_situation(self):
        if self.show_action.text() == "Show":
            self.show()
            self.show_action.setText("Hide")
        else:
            self.hide()
            self.show_action.setText("Show")

    def setup_daily_task(self):
        self.timer = QTimer()
        self.timer.timeout.connect(self.daily_task)
        self.timer.start(24 * 60 * 60 * 1000)  # 24 hours
        QTimer.singleShot(0, self.daily_task)  # Immediate first run

    def load_last_run_date(self):
        try:
            with open("last_run.txt", "r") as f:
                self.last_run_date = datetime.fromisoformat(f.read().strip())
        except FileNotFoundError:
            self.last_run_date = datetime(2000, 1, 1)

    def save_last_run_date(self):
        with open("last_run.txt", "w") as f:
            f.write(datetime.now().isoformat())

    def daily_task(self):
        try:
            if self.process_attachments():
                self.clean_temp_folder()
                self.save_last_run_date()
        except Exception as e:
            self.show_error(f"Daily task failed: {str(e)}")

    def process_attachments(self):
        try:
            # Clear temp folder before processing
            if os.path.exists(TEMP_FOLDER):
                shutil.rmtree(TEMP_FOLDER)
            os.makedirs(TEMP_FOLDER, exist_ok=True)

            # Get emails since last run
            inbox = account.inbox.filter(
                datetime_received__gt=self.last_run_date
            ).order_by('-datetime_received')

            # Process attachments
            for email in inbox:
                for attachment in email.attachments:
                    if isinstance(attachment, FileAttachment):
                        ext = os.path.splitext(attachment.name)[1].lower()
                        if ext in ALLOWED_EXTENSIONS:
                            file_path = os.path.join(TEMP_FOLDER, attachment.name)
                            with open(file_path, 'wb') as f:
                                f.write(attachment.content)
                            if not self.upload_file(file_path):
                                return False
            return True
        except Exception as e:
            self.show_error(f"Attachment processing failed: {str(e)}")
            return False

    def upload_file(self, file_path):
        try:
            with open(file_path, 'rb') as f:
                response = requests.post(
                    UPLOAD_URL,
                    files={'file': f},
                    timeout=30
                )
            if response.status_code != 200:
                self.show_error(f"Upload failed for {file_path}: {response.text}")
                return False
            return True
        except Exception as e:
            self.show_error(f"Upload error: {str(e)}")
            return False

    def clean_temp_folder(self):
        if os.path.exists(TEMP_FOLDER):
            shutil.rmtree(TEMP_FOLDER)

    def show_error(self, message):
        self.tray_icon.showMessage("Error", message, QSystemTrayIcon.MessageIcon.Critical)

    def closeEvent(self, event):
        self.tray_icon.hide()
        event.accept()


class InboxViewer(QMainWindow):
    def __init__(self):
        super().__init__()
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
        try:
            # Convert the inbox to a list to support indexing
            inbox = list(account.inbox.all().order_by('-datetime_received')[:10])
            for email in inbox:
                self.email_list.addItem(f"{email.subject} - {email.sender.email_address}")
            self.emails = inbox  # Store the list of emails
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load emails: {str(e)}")

    def show_email(self, item):
        try:
            index = self.email_list.row(item)
            email = self.emails[index]  # Access the email from the list
            content = f"From: {email.sender.email_address}\n"
            content += f"Date: {email.datetime_received}\n\n"
            content += email.text_body
            self.email_viewer.setText(content)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to display email: {str(e)}")


class EmailComposer(QMainWindow):
    def __init__(self):
        super().__init__()
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
            QMessageBox.information(self, "Attachment Added",
                                    f"Added attachment: {os.path.basename(file_path)}")

    def send_email(self):
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

            msg.send()
            QMessageBox.information(self, "Success", "Email sent successfully!")
            self.close()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to send email: {str(e)}")


class LoginWindow(QMainWindow):
    login_success = pyqtSignal()
    login_failed = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.account = None  # Store account here
        self.setWindowTitle("Login")
        self.setGeometry(100, 100, 400, 250)
        self.otp_required = False
        self.init_ui()
        # Delay the credential check to ensure cipher suite is ready
        QTimer.singleShot(100, self.check_saved_credentials)

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

        self.remember_check = QCheckBox("Remember credentials")
        basic_layout.addWidget(self.remember_check)

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

    def check_saved_credentials(self):
        email, password = SecureCredentials.load_credentials()
        if email and password:
            self.email_input.setText(email)
            self.password_input.setText(password)
            self.attempt_basic_login()

    def attempt_basic_login(self):
        try:
            email = self.email_input.text()
            password = self.password_input.text()
            remember = self.remember_check.isChecked()

            try:
                global account
                credentials = Credentials(username=email, password=password)
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

                if remember:
                    SecureCredentials.save_credentials(email, password)
                else:
                    SecureCredentials.clear_credentials()

                self.login_success.emit()
                self.close()
            except UnauthorizedError as e:
                if "multi-factor authentication" in str(e).lower():
                    self.stacked_widget.setCurrentWidget(self.otp_login_widget)
                    QMessageBox.warning(self, "OTP Required",
                                        "Multi-factor authentication required. Please enter your OTP.")
                else:
                    self.handle_login_error(e)
            except Exception as e:
                self.handle_login_error(e)
        except Exception as e:
            SecureCredentials.clear_credentials()
            QMessageBox.critical(self, "Error",
                                 "Session expired. Please log in again.")
            self.show()

    def attempt_otp_login(self):
        email = self.email_input.text()
        password = self.password_input.text()
        otp = self.otp_input.text()

        try:
            global account
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

            if self.remember_check.isChecked():
                SecureCredentials.save_credentials(email, password)
            else:
                SecureCredentials.clear_credentials()

            self.login_success.emit()
            self.close()
        except Exception as e:
            self.handle_login_error(e, "OTP login failed: ")

    def handle_login_error(self, error, prefix=""):
        SecureCredentials.clear_credentials()
        QMessageBox.critical(self, "Login Failed", f"{prefix}{str(error)}")
        self.login_failed.emit()
        self.show()
        self.email_input.clear()
        self.password_input.clear()
        self.otp_input.clear()
        self.stacked_widget.setCurrentIndex(0)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    login_window = LoginWindow()
    client_window = EmailClient()

    login_window.login_success.connect(client_window.show)
    login_window.login_failed.connect(login_window.show)

    login_window.show()
    sys.exit(app.exec())
