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
from PyQt6.QtCore import Qt, pyqtSignal, QTimer
from PyQt6.QtGui import QIcon, QAction
from exchangelib import Credentials, Account, DELEGATE, Configuration, Message, Mailbox, FileAttachment
from exchangelib.errors import UnauthorizedError
from cryptography.fernet import Fernet

# Configuration
CONFIG_FILE = "email_client_config.json"
TEMP_FOLDER = os.path.join(gettempdir(), "email_attachments")
ALLOWED_EXTENSIONS = {'.xlsx', '.xls', '.csv', '.xlsm', '.xlsb'}
UPLOAD_URL = "https://bi.mci.ir/myflask/upload"
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
        except Exception:
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
        except Exception:
            SecureCredentials.clear_credentials()
            return None, None


class LoginWindow(QMainWindow):
    login_success = pyqtSignal(object)  # Signal to pass the account object

    def __init__(self):
        super().__init__()
        self.account = None
        self.setWindowTitle("Login")
        self.setGeometry(100, 100, 400, 250)
        self.init_ui()
        QTimer.singleShot(100, self.check_saved_credentials)

    def init_ui(self):
        layout = QVBoxLayout()
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("Email")
        layout.addWidget(self.email_input)

        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Password")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.password_input)

        self.remember_check = QCheckBox("Remember credentials")
        layout.addWidget(self.remember_check)

        self.login_btn = QPushButton("Login")
        self.login_btn.clicked.connect(self.attempt_basic_login)
        layout.addWidget(self.login_btn)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def check_saved_credentials(self):
        email, password = SecureCredentials.load_credentials()
        if email and password:
            self.attempt_basic_login(email, password, auto=True)

    def attempt_basic_login(self, email=None, password=None, auto=False):
        email = email or self.email_input.text()
        password = password or self.password_input.text()
        remember = self.remember_check.isChecked()
        try:
            credentials = Credentials(username=email, password=password)
            config = Configuration(
                credentials=credentials,
                service_endpoint="https://mail.mci.ir/ews/exchange.asmx"
            )
            self.account = Account(
                primary_smtp_address=email,
                config=config,
                autodiscover=False,
                access_type=DELEGATE
            )
            self.account.root.refresh()

            if remember:
                SecureCredentials.save_credentials(email, password)
            self.login_success.emit(self.account)
            self.close()
        except Exception as e:
            if not auto:
                QMessageBox.critical(self, "Login Failed", str(e))


class EmailClient(QMainWindow):
    def __init__(self, account):
        super().__init__()
        self.account = account
        self.setWindowTitle("Email Client")
        self.setGeometry(100, 100, 800, 600)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        self.status_label = QLabel("Service is running")
        layout.addWidget(self.status_label)

        self.process_btn = QPushButton("Process Attachments")
        self.process_btn.clicked.connect(self.process_attachments)
        layout.addWidget(self.process_btn)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def process_attachments(self):
        try:
            if not self.account:
                QMessageBox.warning(self, "Warning", "No active account connection")
                return

            inbox = self.account.inbox.filter(
                datetime_received__gt=datetime.now() - timedelta(days=1)
            ).order_by('-datetime_received')

            for email in inbox:
                for attachment in email.attachments:
                    if isinstance(attachment, FileAttachment):
                        ext = os.path.splitext(attachment.name)[1].lower()
                        if ext in ALLOWED_EXTENSIONS:
                            file_path = os.path.join(TEMP_FOLDER, attachment.name)
                            with open(file_path, 'wb') as f:
                                f.write(attachment.content)
                            requests.post(UPLOAD_URL, files={'file': open(file_path, 'rb')})
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Processing failed: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    login_window = LoginWindow()
    client_window = None


    def handle_login_success(account):
        global client_window
        client_window = EmailClient(account)
        client_window.show()


    login_window.login_success.connect(handle_login_success)
    login_window.show()
    sys.exit(app.exec())