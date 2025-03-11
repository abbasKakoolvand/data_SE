import sys
import os
import json
import time
from datetime import datetime, timedelta

import requests
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QListWidget,
    QMessageBox, QFileDialog, QStackedWidget
)
from PyQt6.QtCore import Qt, pyqtSignal, QTimer
from exchangelib import Credentials, Account, DELEGATE, Configuration, Message, Mailbox, FileAttachment, Folder
from exchangelib.errors import UnauthorizedError

# Global account variable
account = None


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
        print(f"Attempting basic login with email: {email}")
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
            print("Basic login successful!")
            self.login_success.emit()
            self.close()
        except UnauthorizedError as e:
            print(f"UnauthorizedError encountered: {str(e)}")
            if "multi-factor authentication" in str(e).lower():
                self.stacked_widget.setCurrentWidget(self.otp_login_widget)
                QMessageBox.warning(self, "OTP Required",
                                    "Multi-factor authentication required. Please enter your OTP.")
            else:
                QMessageBox.critical(self, "Login Failed", str(e))
        except Exception as e:
            print(f"Error during basic login: {str(e)}")
            QMessageBox.critical(self, "Error", f"Login failed: {str(e)}")

    def attempt_otp_login(self):
        email = self.email_input.text()
        password = self.password_input.text()
        otp = self.otp_input.text()
        print(f"Attempting OTP login with email: {email} and OTP provided.")
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
            print("OTP login successful!")
            self.login_success.emit()
            self.close()
        except Exception as e:
            print(f"Error during OTP login: {str(e)}")
            QMessageBox.critical(self, "Login Failed", f"OTP login failed: {str(e)}")


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


class EmailClient(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Email Client")
        self.setGeometry(100, 100, 800, 600)
        self.init_ui()
        self.emails_to_process = []  # To hold emails for attachment processing

    def init_ui(self):
        central_widget = QWidget()
        layout = QVBoxLayout()

        # Buttons layout
        btn_layout = QHBoxLayout()
        self.send_btn = QPushButton("Send Email")
        self.send_btn.clicked.connect(self.open_composer)
        btn_layout.addWidget(self.send_btn)

        self.inbox_btn = QPushButton("Check Inbox")
        self.inbox_btn.clicked.connect(self.open_inbox)
        btn_layout.addWidget(self.inbox_btn)

        # New button for processing attachments
        self.process_attachments_btn = QPushButton("Process Attachments")
        self.process_attachments_btn.clicked.connect(self.process_attachments)
        btn_layout.addWidget(self.process_attachments_btn)

        # New button for processing attachments
        self.upload_attachments_btn = QPushButton("list attachments")
        self.upload_attachments_btn.clicked.connect(self.upload_attachments)
        btn_layout.addWidget(self.upload_attachments_btn)

        layout.addLayout(btn_layout)
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def upload_attachments(self):
        with open("config.json", "r") as f:
            config = json.load(f)
        destination_folder = config.get("destination_folder")
        attachment_files = os.listdir(destination_folder)
        upload_Done = True
        for ind, file in enumerate(attachment_files):
            full_file_path = os.path.join(destination_folder, file)  # Get full path
            print(f"{ind}: {file} : {full_file_path}")

            # if ind == 0:
            if True:
                url = "https://bi.mci.ir/myflask/upload"

                # Use the full path to the file
                with open(full_file_path, 'rb') as f:
                    response = requests.post(url, files={'file': f}, verify=False)

                print(response.status_code)
                print(response.json())
                if response.status_code != 200:
                    upload_Done = False
                time.sleep(60)
        if upload_Done:
            update_time_config()
        else:
            pass

    def open_composer(self):
        print("Opening Email Composer.")
        self.composer = EmailComposer()
        self.composer.show()

    def open_inbox(self):
        print("Opening Inbox Viewer.")
        self.inbox = InboxViewer()
        self.inbox.show()

    def process_attachments(self):
        """
        Reads config from a JSON file and processes emails (one per minute)
        to save Excel/CSV attachments from emails received after a specified date.
        """
        print("Starting process_attachments()")
        # Load configuration from JSON file
        try:
            with open("config.json", "r") as f:
                config = json.load(f)
            start_date_str = config.get("start_date")
            destination_folder = config.get("destination_folder")
            print(f"Config loaded: start_date={start_date_str}, destination_folder={destination_folder}")
            if not start_date_str or not destination_folder:
                QMessageBox.critical(self, "Error", "Configuration must include 'start_date' and 'destination_folder'.")
                return

            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            if not os.path.exists(destination_folder):
                print(f"Destination folder {destination_folder} does not exist. Creating it.")
                os.makedirs(destination_folder)
        except Exception as e:
            print(f"Error reading config file: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to read config file: {str(e)}")
            return

        # Retrieve emails from the inbox received after the start_date
        try:
            print("Retrieving emails from inbox...")
            # Convert to list for further processing
            tic = time.time()
            all_emails = list(account.inbox.all().order_by('datetime_received'))
            toc = time.time()
            print(f"all mails are: {len(all_emails)} - in {int(toc - tic)} sec")
            self.emails_to_process = [email for email in all_emails if
                                      email.datetime_received.replace(tzinfo=None) > start_date]
            print(f"all mails after {start_date_str} are: {len(self.emails_to_process)}")
            self.emails_to_process.sort(key=lambda email: email.datetime_received)
            print(f"Found {len(self.emails_to_process)} email(s) after {start_date_str}.")
            if not self.emails_to_process:
                QMessageBox.information(self, "No Emails", "No emails found after the specified date.")
                return
        except Exception as e:
            print(f"Error retrieving emails: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to retrieve emails: {str(e)}")
            return

        QMessageBox.information(self, "Processing",
                                f"Found {len(self.emails_to_process)} email(s) to process. Processing will begin now.")
        # Store destination folder to use in processing
        self.destination_folder = destination_folder
        self.process_next_email()

    def process_next_email(self):
        """
        Processes the next email in the list by saving its Excel/CSV attachments.
        Waits one minute between each email.
        """
        if not self.emails_to_process:
            print("No more emails to process.")
            QMessageBox.information(self, "Done", "All emails have been processed.")
            return

        email = self.emails_to_process.pop(0)
        print(f"Processing email: {email.subject}, received at {email.datetime_received}")
        try:
            attachments_saved = False
            for ind, attachment in enumerate(email.attachments):
                print(f"attachment: {ind}, name: {attachment.name}")
                if isinstance(attachment, FileAttachment):
                    if attachment.name.lower().endswith(('.xlsx', '.xls', '.csv')):
                        file_path = os.path.join(self.destination_folder, attachment.name)
                        with open(file_path, 'wb') as f:
                            f.write(attachment.content)
                        print(f"Saved attachment: {attachment.name} to {file_path}")
                        attachments_saved = True
            if attachments_saved:
                print(f"Attachments from email '{email.subject}' saved successfully.")
            else:
                print(f"No valid attachments found in email '{email.subject}'.")
        except Exception as e:
            print(f"Error processing email '{email.subject}': {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to save attachments from email '{email.subject}': {str(e)}")

        # Schedule the processing of the next email after one minute (60000 ms)
        print("Waiting 1 minute before processing the next email...")
        QTimer.singleShot(5000, self.process_next_email)


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
            print(f"Attachment added: {os.path.basename(file_path)}")
            QMessageBox.information(self, "Attachment Added",
                                    f"Added attachment: {os.path.basename(file_path)}")

    def send_email(self):
        print("Sending email...")
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
            QMessageBox.information(self, "Success", "Email sent successfully!")
            self.close()
        except Exception as e:
            print(f"Error sending email: {str(e)}")
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
        print("Loading emails for Inbox Viewer...")
        try:
            inbox = list(account.inbox.all().order_by('-datetime_received')[:10])
            for email in inbox:
                self.email_list.addItem(f"{email.subject} - {email.sender.email_address}")
            self.emails = inbox
            print(f"Loaded {len(inbox)} email(s) into the inbox list.")
        except Exception as e:
            print(f"Error loading emails: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to load emails: {str(e)}")

    def show_email(self, item):
        try:
            index = self.email_list.row(item)
            email = self.emails[index]
            print(f"Displaying email: {email.subject}")
            content = f"From: {email.sender.email_address}\n"
            content += f"Date: {email.datetime_received}\n\n"
            content += email.text_body
            self.email_viewer.setText(content)
        except Exception as e:
            print(f"Error displaying email: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to display email: {str(e)}")


if __name__ == "__main__":
    print("Starting Email Client Application...")
    app = QApplication(sys.argv)

    login_window = LoginWindow()
    client_window = EmailClient()

    login_window.login_success.connect(client_window.show)

    login_window.show()
    sys.exit(app.exec())
