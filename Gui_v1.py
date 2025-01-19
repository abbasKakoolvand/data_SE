import sys
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QVBoxLayout,
    QWidget,
    QLineEdit,
    QComboBox,
    QPushButton,
    QTableView,
    QMessageBox,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QStandardItemModel, QStandardItem


class SearchApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # Window setup
        self.setWindowTitle("Excel Search Tool")
        self.setGeometry(100, 100, 800, 600)

        # Load the Excel file and sheet
        self.file_path = "warehouse SE/TablesDataEDO_2.xlsx"
        self.sheet_name = "TablesandColumnsinDW"
        self.df = pd.read_excel(self.file_path, sheet_name=self.sheet_name)

        # Initialize UI
        self.init_ui()

    def init_ui(self):
        # Main layout
        layout = QVBoxLayout()

        # Search field
        self.search_field = QLineEdit(self)
        self.search_field.setPlaceholderText("Enter text to search...")
        layout.addWidget(self.search_field)

        # Dropdown menu for column selection
        self.column_dropdown = QComboBox(self)
        self.column_dropdown.addItems(self.df.columns)  # Add all columns to the dropdown
        self.column_dropdown.setCurrentIndex(0)  # Set default selection
        layout.addWidget(self.column_dropdown)

        # Search button
        self.search_button = QPushButton("Search", self)
        self.search_button.clicked.connect(self.on_search)
        layout.addWidget(self.search_button)

        # Table view to display results
        self.table_view = QTableView(self)
        self.model = QStandardItemModel()
        self.table_view.setModel(self.model)
        layout.addWidget(self.table_view)

        # Set layout to central widget
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def on_search(self):
        # Get search text and selected column
        search_text = self.search_field.text().strip().lower()
        selected_column = self.column_dropdown.currentText()

        if not search_text:
            QMessageBox.warning(self, "Error", "Please enter a search term.")
            return

        # Filter the DataFrame based on the selected column and search text
        filtered_df = self.df[
            self.df[selected_column].astype(str).str.lower().str.contains(search_text)
        ]

        if filtered_df.empty:
            QMessageBox.information(self, "No Results", "No matching rows found.")
            return

        # Display the filtered results in the table view
        self.model.clear()
        self.model.setHorizontalHeaderLabels(filtered_df.columns)

        for _, row in filtered_df.iterrows():
            items = [QStandardItem(str(item)) for item in row]
            self.model.appendRow(items)


# Run the application
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SearchApp()
    window.show()
    sys.exit(app.exec())