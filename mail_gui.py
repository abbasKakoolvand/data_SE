import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QListWidget,
    QMessageBox, QFileDialog, QStackedWidget
)
from PyQt6.QtCore import Qt, pyqtSignal
from exchangelib import Credentials, Account, DELEGATE, Configuration, Message, Mailbox, FileAttachment, Folder
from exchangelib.errors import UnauthorizedError


class LoginWindow(QMainWindow):
    login_success = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Login")
        self.setGeometry(100, 100, 400, 250)
        self.otp_required = False
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

    def attempt_basic_login(self):
        email = self.email_input.text()
        password = self.password_input.text()

        try:
            global account
            credentials = Credentials(username=email, password=password)
            config = Configuration(
                credentials=credentials,
                service_endpoint="https://mail.mci.ir/ews/exchange.asmx"  # Only service_endpoint
            )
            account = Account(
                primary_smtp_address=email,
                config=config,
                autodiscover=False,  # Disable autodiscovery
                access_type=DELEGATE
            )
            # Test connection
            account.root.refresh()
            self.login_success.emit()
            self.close()
        except UnauthorizedError as e:
            if "multi-factor authentication" in str(e).lower():
                self.stacked_widget.setCurrentWidget(self.otp_login_widget)
                QMessageBox.warning(self, "OTP Required",
                                    "Multi-factor authentication required. Please enter your OTP.")
            else:
                QMessageBox.critical(self, "Login Failed", str(e))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Login failed: {str(e)}")

    def attempt_otp_login(self):
        email = self.email_input.text()
        password = self.password_input.text()
        otp = self.otp_input.text()

        try:
            global account
            credentials = Credentials(username=email, password=f"{password}{otp}")
            config = Configuration(
                credentials=credentials,
                service_endpoint="https://mail.mci.ir/ews/exchange.asmx"  # Only service_endpoint
            )
            account = Account(
                primary_smtp_address=email,
                config=config,
                autodiscover=False,  # Disable autodiscovery
                access_type=DELEGATE
            )
            # Test connection
            account.root.refresh()
            self.login_success.emit()
            self.close()
        except Exception as e:
            QMessageBox.critical(self, "Login Failed", f"OTP login failed: {str(e)}")


class EmailClient(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Email Client")
        self.setGeometry(100, 100, 800, 600)
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        layout = QVBoxLayout()

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


if __name__ == "__main__":
    app = QApplication(sys.argv)

    login_window = LoginWindow()
    client_window = EmailClient()

    login_window.login_success.connect(client_window.show)

    login_window.show()
    sys.exit(app.exec())