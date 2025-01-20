import os
import sys
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QVBoxLayout,
    QWidget,
    QLineEdit,
    QPushButton,
    QTableView,
    QMessageBox,
    QListWidget,
    QLabel,
    QComboBox,
    QHBoxLayout,
    QListWidgetItem,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QStandardItemModel, QStandardItem, QColor


class SearchApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # Window setup
        self.setWindowTitle("Excel Search Tool")
        self.setGeometry(100, 100, 800, 600)

        # Load the Excel file
        self.file_path = "warehouse SE/TablesDataEDO_2.xlsx"
        self.sheets = pd.ExcelFile(self.file_path).sheet_names  # Get all sheet names
        self.current_sheet = self.sheets[0]  # Default to the first sheet
        self.df = pd.read_excel(self.file_path, sheet_name=self.current_sheet)

        # Initialize UI
        self.init_ui()

    def init_ui(self):
        # Main layout
        layout = QVBoxLayout()

        # Sheet selector
        self.sheet_selector = QComboBox(self)
        self.sheet_selector.addItems(self.sheets)
        self.sheet_selector.currentTextChanged.connect(self.on_sheet_change)
        layout.addWidget(QLabel("Select Sheet:"))
        layout.addWidget(self.sheet_selector)

        # Search field
        self.search_field = QLineEdit(self)
        self.search_field.setPlaceholderText("Enter text to search...")
        layout.addWidget(self.search_field)

        # Search button
        self.search_button = QPushButton("Search", self)
        self.search_button.clicked.connect(self.on_search)
        layout.addWidget(self.search_button)

        # Label for column selection
        layout.addWidget(QLabel("Select columns to search:"))

        # List widget with checkboxes for column selection
        self.column_list_widget = QListWidget(self)
        self.update_column_list()
        layout.addWidget(self.column_list_widget)

        # Table view to display results
        self.table_view = QTableView(self)
        self.model = QStandardItemModel()
        self.table_view.setModel(self.model)
        layout.addWidget(self.table_view)

        # Display all rows initially
        self.display_all_rows()

        # Set layout to central widget
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def update_column_list(self):
        """Update the column list widget based on the current sheet's columns."""
        self.column_list_widget.clear()
        for column in self.df.columns:
            item = QListWidgetItem(column)
            item.setCheckState(Qt.CheckState.Checked)  # Default checked
            self.column_list_widget.addItem(item)

    def display_all_rows(self):
        """Display all rows of the current sheet in the table view."""
        self.model.clear()
        self.model.setHorizontalHeaderLabels(self.df.columns)
        for _, row in self.df.iterrows():
            items = [QStandardItem(str(item)) for item in row]
            self.model.appendRow(items)
        self.norm_table_columns()

    def on_sheet_change(self):
        """Handle sheet selection change."""
        self.current_sheet = self.sheet_selector.currentText()
        self.df = pd.read_excel(self.file_path, sheet_name=self.current_sheet)
        self.update_column_list()
        self.display_all_rows()

    def on_search(self):
        """Handle search button click."""
        # Get search text
        search_text = self.search_field.text().strip().lower()

        if not search_text:
            QMessageBox.warning(self, "Error", "Please enter a search term.")
            return

        # Get selected columns (checked items)
        selected_columns = []
        for index in range(self.column_list_widget.count()):
            item = self.column_list_widget.item(index)
            if item.checkState() == Qt.CheckState.Checked:
                selected_columns.append(item.text())

        if not selected_columns:
            QMessageBox.warning(self, "Error", "Please select at least one column.")
            return

        # Filter the DataFrame based on the selected columns and search text
        filtered_df = self.df[
            self.df[selected_columns]
            .astype(str)
            .apply(lambda x: x.str.lower().str.contains(search_text))
            .any(axis=1)
        ]

        if filtered_df.empty:
            QMessageBox.information(self, "No Results", "No matching rows found.")
            return

        # Display the filtered results in the table view and highlight matching cells
        self.model.clear()
        self.model.setHorizontalHeaderLabels(filtered_df.columns)
        col_num = {}
        for _, row in filtered_df.iterrows():
            items = []
            for col in filtered_df.columns:
                cell_value = str(row[col])
                item = QStandardItem(cell_value)
                if search_text in cell_value.lower():
                    try:
                        col_num[col] += 1
                    except:
                        col_num[col] = 1
                    item.setBackground(QColor("yellow"))  # Highlight matching cells
                items.append(item)
            self.model.appendRow(items)

        # Update column headers with match counts
        columns = []
        for col in filtered_df.columns:
            try:
                columns.append(f"{col} ({col_num[col]})")
            except:
                columns.append(f"{col} (0)")
        self.model.setHorizontalHeaderLabels(columns)
        self.norm_table_columns()

    def norm_table_columns(self):
        # Set column widths based on content size and a minimum of 200 pixels
        for column in range(self.model.columnCount()):
            # Resize to fit content
            self.table_view.resizeColumnToContents(column)
            # Get the current column width
            current_width = self.table_view.columnWidth(column)
            # Set the column width to the maximum of 200 and the content width
            self.table_view.setColumnWidth(column, min(200, current_width))


# Run the application
if __name__ == "__main__":
    app = QApplication(sys.argv)
    try:
        script_dir = os.path.dirname(__file__)
        style_dir = os.path.join(script_dir, "files/MacOS.qss")
        with open(style_dir, "r") as file:
            app.setStyleSheet(file.read())
    except Exception as e:
        print(e)
    window = SearchApp()
    window.show()
    sys.exit(app.exec())
