import os
import re
import sys

import nltk
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
    QGroupBox, QCompleter,
)
from PyQt6.QtCore import Qt, QStringListModel
from PyQt6.QtGui import QStandardItemModel, QStandardItem, QColor
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
import string

nltk_data_dir = os.path.join(os.getcwd(), 'files/nltk_data')
nltk.data.path.append(nltk_data_dir)


def are_all_signs(word):
    # Check if the word contains only non-alphanumeric characters
    return bool(re.fullmatch(r'[^a-zA-Z0-9]+', word))


class SearchApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # Window setup
        self.word_list = None
        self.search_text = None
        self.schema_column_name = None
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
        self.scan_text_of_df()
        # Main layout
        layout = QVBoxLayout()

        # Create a group box for the top section
        group_box = QGroupBox("Search Options")  # Title for the group box
        self.result_group_box = QGroupBox("Search Result Filters")  # Title for the result filters
        layout_sheet_items = QHBoxLayout()
        layout_search_items = QVBoxLayout()
        layout_search_btn = QVBoxLayout()
        layout_list_column = QVBoxLayout()
        layout_schema_column = QVBoxLayout()
        layout_result_filter = QHBoxLayout()
        layout_list_search_aggr = QHBoxLayout()

        # Sheet selector
        maximum_size = 200
        self.sheet_selector = QComboBox(self)
        self.sheet_selector.addItems(self.sheets)
        self.sheet_selector.currentTextChanged.connect(self.on_sheet_change)
        sheet_label = QLabel("Select Sheet:")
        sheet_label.setMaximumWidth(maximum_size)
        self.sheet_selector.setMaximumWidth(maximum_size)

        layout_sheet_items.addWidget(sheet_label)
        layout_sheet_items.addWidget(self.sheet_selector)
        layout_sheet_items.addWidget(QLabel(""))

        # Search field
        self.search_field = QLineEdit(self)
        self.search_field.setPlaceholderText("Enter text to search...")
        self.search_field.setMaximumWidth(maximum_size)
        self.completer = QCompleter()
        self.completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.completer.setCompletionMode(QCompleter.CompletionMode.UnfilteredPopupCompletion)

        self.search_field.textChanged.connect(lambda text: self.suggest_words(text, self.search_field, self.completer))

        layout_search_items.addWidget(self.search_field)

        self.search_field.returnPressed.connect(self.on_search)

        # Search button
        self.search_button = QPushButton("Search", self)
        self.search_button.clicked.connect(self.on_search)
        self.search_button.setMaximumWidth(maximum_size)
        layout_search_btn.addWidget(self.search_button)

        # Add layouts to the main layout of the group box

        layout_list_search_aggr.addLayout(layout_sheet_items)
        layout_list_search_aggr.addWidget(QLabel(""))
        layout_list_search_aggr.addLayout(layout_search_items)
        layout_list_search_aggr.addLayout(layout_search_btn)
        layout_list_search_aggr.addWidget(QLabel(""))
        layout_list_search_aggr.addWidget(QLabel(""))
        layout_list_search_aggr.addWidget(QLabel(""))

        # Label for column selection
        list_column_label = QLabel("Selected columns to search:")
        list_column_label.setMaximumHeight(10)
        layout_list_column.addWidget(list_column_label)

        # List widget with checkboxes for column selection
        self.column_list_widget = QListWidget(self)
        self.column_list_widget.setMaximumHeight(100)  # Adjust height as needed
        self.column_list_widget.setMaximumWidth(maximum_size)
        self.column_list_widget.itemChanged.connect(self.on_search)

        self.update_column_list()

        # Add the column list widget to the column layout
        layout_list_column.addWidget(self.column_list_widget)
        # list_column_spacer = QLabel("")
        # layout_list_column.addWidget(list_column_spacer)

        self.schema_column_label = QLabel("Select schema to filter:")
        self.schema_column_label.setMaximumHeight(10)
        layout_schema_column.addWidget(self.schema_column_label)
        # Create a combo box for filtering by unique values
        self.filter_combo_box = QComboBox(self)
        self.filter_combo_box.currentIndexChanged.connect(self.filter_results)
        self.filter_combo_box.setMaximumWidth(maximum_size)
        self.filter_combo_box.hide()  # Hide initially
        layout_schema_column.addWidget(self.filter_combo_box)

        schema_column_spacer = QLabel("")
        layout_schema_column.addWidget(schema_column_spacer)

        layout_result_filter.addLayout(layout_list_column)
        layout_result_filter.addLayout(layout_schema_column)

        # Set the layout for the group box
        group_box.setLayout(layout_list_search_aggr)
        self.result_group_box.setLayout(layout_result_filter)
        self.result_group_box.setMaximumHeight(120)
        self.result_group_box.hide()

        # Add the group box to the main layout
        layout.addWidget(group_box)
        layout.addWidget(self.result_group_box)

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
        self.result_group_box.hide()

    def on_search(self):
        """Handle search button click."""
        # Get search text
        search_text = self.search_field.text().strip().lower()

        if not search_text:
            QMessageBox.warning(self, "Error", "Please enter a search term.")
            return

        # Get selected columns (checked items)
        self.selected_columns = []
        for index in range(self.column_list_widget.count()):
            item = self.column_list_widget.item(index)
            if item.checkState() == Qt.CheckState.Checked:
                self.selected_columns.append(item.text())

        if not self.selected_columns:
            QMessageBox.warning(self, "Error", "Please select at least one column.")
            return
        self.search_text = search_text
        # Filter the DataFrame based on the selected columns and search text
        filtered_df = self.df[
            self.df[self.selected_columns]
            .astype(str)
            .apply(lambda x: x.str.lower().str.contains(search_text))
            .any(axis=1)
        ]

        if filtered_df.empty:
            QMessageBox.information(self, "No Results", "No matching rows found.")
            self.filter_combo_box.hide()  # Hide combo box if no results
            return

        # Display the filtered results in the table view and highlight matching cells
        self.model.clear()
        self.model.setHorizontalHeaderLabels(filtered_df.columns)

        schema_columns = [col for col in filtered_df.columns if "schema" in col.lower()]

        if schema_columns:  # Check if there are any schema columns
            self.schema_column_name = schema_columns[0]
            # Prepare the combo box for filtering
            self.filter_combo_box.clear()

            self.filter_combo_box.addItem(f"All Items ({len(filtered_df)})")
            unique_values = filtered_df[
                self.schema_column_name].unique()  # Get unique values from the first schema column
            self.unique_scema_values = unique_values
            self.filter_combo_box.addItems(
                [f"{str(value)} ({len(filtered_df[filtered_df[self.schema_column_name].astype(str) == str(value)])})"
                 for value in unique_values])
            self.schema_column_label.show()
            self.filter_combo_box.show()  # Show the combo box
        else:
            self.filter_combo_box.addItem(f"All Items")
            self.filter_combo_box.hide()  # Hide the combo box if no schema columns
            self.schema_column_label.hide()

        self.filter_results()
        self.norm_table_columns()
        self.result_group_box.show()

    def filter_results(self):
        if self.search_text is not None:
            col_num = {}
            """Filter the displayed results based on the selected value in the combo box."""
            selected_value = self.filter_combo_box.currentText()
            selected_index = self.filter_combo_box.currentIndex()
            if selected_value:
                filtered_df_search = self.df[
                    self.df[self.selected_columns]
                    .astype(str)
                    .apply(lambda x: x.str.lower().str.contains(self.search_text))
                    .any(axis=1)
                ]
                if selected_index == 0:
                    # Filter the DataFrame based on the selected value
                    filtered_schema_df = filtered_df_search

                else:

                    # Filter the DataFrame based on the selected value
                    filtered_schema_df = filtered_df_search[
                        filtered_df_search[self.schema_column_name].astype(str) == self.unique_scema_values[
                            selected_index - 1]]
                self.model.clear()
                for _, row in filtered_schema_df.iterrows():
                    items = []
                    for col in filtered_schema_df.columns:
                        cell_value = str(row[col])
                        item = QStandardItem(cell_value)
                        if self.search_text in cell_value.lower():
                            try:
                                col_num[col] += 1
                            except:
                                col_num[col] = 1
                            item.setBackground(QColor("yellow"))  # Highlight matching cells
                        items.append(item)
                    self.model.appendRow(items)

                # Update column headers with match counts
                columns = []
                for col in filtered_schema_df.columns:
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

    def scan_text_of_df(self):

        # Convert each row to a string with values joined by a space
        row_strings = self.df.apply(lambda row: ' '.join(row.values.astype(str)), axis=1)

        # Join all row strings with a newline character
        final_text = '\n'.join(row_strings)

        word_frequencies = {}

        text = final_text.replace("&", "ANDAND")

        text = text.replace("/", " ")
        text = text.replace("'", " ")
        text = text.replace("+", " ")
        text = text.replace(".", " ")
        text = text.replace("_", " ")

        # Tokenize text
        tokens = word_tokenize(text)
        tokens = [token.replace("ANDAND", "&") if token.__contains__("ANDAND") else token for token in tokens]

        try:
            tokens.remove("ANDAND")
        except:
            pass
        # tokens = word_tokenize(text_slide)

        # Get stopwords and punctuation
        stop_words = set(stopwords.words('english'))
        punctuation = set(string.punctuation)
        for token in tokens:
            if len(token) < 2:
                continue
            if are_all_signs(token):
                # print(token)
                continue

            # Skip stopwords and punctuation
            if token.lower() in stop_words and (token.islower() or token.istitle()) or (token in punctuation):
                continue
                # Convert to lower case

            if token not in word_frequencies:
                word_frequencies[token] = 1
            else:
                word_frequencies[token] += 1
        self.word_list = word_frequencies.keys()

    def suggest_words(self, text, search_field, completer):
        # if len(text) > 1:
        if True:
            try:
                words = self.get_word_suggestions(text)
                suggestions = sorted(words, key=len)
                completer.setModel(QStringListModel(suggestions))
                self.search_field.setCompleter(completer)
            except Exception as e5:
                print(e5)

    def get_word_suggestions(self, prefix):
        # Replace this with your actual word list
        return [word for word in self.word_list if str(prefix).lower() in str(word).lower()]


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
