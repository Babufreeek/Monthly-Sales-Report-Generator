import sys
import openpyxl
import os
from monthly_sales_calculations import total_sales
import styles
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog, QVBoxLayout, QComboBox, QCheckBox, QMessageBox

class ExcelForm(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        # Set default width and height of window
        self.resize(styles.width, styles.height)

        # Set default font size and dimensions of buttons and text
        self.setStyleSheet(styles.style_sheet)

        # Select Excel File
        self.excel_file_label = QLabel('Select Excel File:')
        self.excel_file_edit = QLineEdit()
        self.excel_file_button = QPushButton('Browse')
        self.excel_file_button.clicked.connect(self.get_excel_file)

        # Select Worksheet
        self.worksheet_label = QLabel('Select Worksheet:')
        self.worksheet_combo = QComboBox()

        # Translation Source File
        self.translation_source_label = QLabel('Select Translation Source Excel File:')
        # Search for translation source and autofill if possible
        search_results = ExcelForm.autofill_translation_source()
        self.translation_source_edit = QLineEdit(search_results if search_results else '')
        self.translation_source_button = QPushButton('Browse')
        self.translation_source_button.clicked.connect(self.get_translation_source)

        # Option to save translated table as a new worksheet in the existing file
        self.save_translations_checkbox = QCheckBox('Save Translations')

        # Option to add the data into the existing spreadsheet
        self.add_to_existing_checkbox = QCheckBox('Add to Existing Spreadsheet')
        self.add_to_existing_checkbox.stateChanged.connect(self.toggle_add_to_existing)

        # Name of new worksheet to be added
        self.worksheet_to_add_label = QLabel('Worksheet to Add:')
        self.worksheet_to_add_edit = QLineEdit()
        self.worksheet_to_add_edit.setEnabled(False)

        # Option to output the data in the brand new spreadsheet (selected by default)
        self.create_new_spreadsheet_checkbox = QCheckBox('Create New Spreadsheet')
        self.create_new_spreadsheet_checkbox.setChecked(True)  # Checked by default
        self.create_new_spreadsheet_checkbox.stateChanged.connect(self.toggle_create_new_spreadsheet)

        # Filename of new spreadsheet
        self.new_filename_label = QLabel('New Filename:')
        self.new_filename_edit = QLineEdit('Result.xlsx')
        self.new_filename_edit.setEnabled(True)

        # Folder to save output in
        self.output_location_label = QLabel('Output Location:')
        self.output_location_edit = QLineEdit()
        self.output_location_edit.setEnabled(True)

        self.browse_location_button = QPushButton('Browse')
        self.browse_location_button.setEnabled(True)
        self.browse_location_button.clicked.connect(self.get_output_location)

        # Worksheet name in the new file
        self.new_worksheet_label = QLabel('New Worksheet Name:')
        self.new_worksheet_edit = QLineEdit('Sheet1')
        self.new_worksheet_edit.setEnabled(True)

        # Submit button
        self.submit_button = QPushButton('Submit')
        self.submit_button.clicked.connect(self.submit_form)

        # Set default value for Worksheet to Add
        self.worksheet_to_add_edit.setText('Monthly Sales Calculations')

        layout = QVBoxLayout()
        layout.addWidget(self.excel_file_label)
        layout.addWidget(self.excel_file_edit)
        layout.addWidget(self.excel_file_button)

        layout.addWidget(self.worksheet_label)
        layout.addWidget(self.worksheet_combo)

        layout.addWidget(self.translation_source_label)
        layout.addWidget(self.translation_source_edit)
        layout.addWidget(self.translation_source_button)

        layout.addWidget(self.save_translations_checkbox)
        layout.addWidget(self.add_to_existing_checkbox)

        layout.addWidget(self.worksheet_to_add_label)
        layout.addWidget(self.worksheet_to_add_edit)

        layout.addWidget(self.create_new_spreadsheet_checkbox)

        layout.addWidget(self.new_filename_label)
        layout.addWidget(self.new_filename_edit)

        layout.addWidget(self.output_location_label)
        layout.addWidget(self.output_location_edit)

        layout.addWidget(self.browse_location_button)

        layout.addWidget(self.new_worksheet_label)
        layout.addWidget(self.new_worksheet_edit)

        layout.addWidget(self.submit_button)

        self.setLayout(layout)
        

    def get_excel_file(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.AnyFile)
        file_dialog.setNameFilter('Excel Files (*.xlsx *.xls *.xlsm)')
        if file_dialog.exec_():
            selected_file = file_dialog.selectedFiles()[0]
            self.excel_file_edit.setText(selected_file)

            # Load worksheets into the combo box
            self.load_worksheets(selected_file)

    @classmethod
    def autofill_translation_source(cls):
        """
        Look in the present working directory for the translation file
        """
        # Get list of files in all the directory and subdirectory
        for root, dirs, files in os.walk(os.getcwd()):
            for file in files:
                # Return file path if a matching file is found
                if "language translation" in file.lower():
                    return os.path.join(root, file)
        # Else return None at the end
        return None
    
    def get_translation_source(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.AnyFile)
        file_dialog.setNameFilter('Excel Files (*.xlsx *.xls *.xlsm)')
        if file_dialog.exec_():
            selected_file = file_dialog.selectedFiles()[0]
            self.translation_source_edit.setText(selected_file)

    def load_worksheets(self, excel_file):
        self.worksheet_combo.clear()

        print("Loading worksheets...")

        try:
            workbook = openpyxl.load_workbook(excel_file, read_only=True)
            self.worksheet_combo.addItems(workbook.sheetnames)
        except Exception as e:
            print(f"Error loading worksheets: {e}")

        print("Worksheets loaded!")

    def toggle_add_to_existing(self, state):
        self.worksheet_to_add_edit.setEnabled(state == 2)  # 2 is checked, 0 is unchecked

    def toggle_create_new_spreadsheet(self, state):
        self.new_filename_edit.setEnabled(state == 2)  # 2 is checked, 0 is unchecked
        self.browse_location_button.setEnabled(state == 2)  # 2 is checked, 0 is unchecked
        self.output_location_edit.setEnabled(state == 2)  # 2 is checked, 0 is unchecked
        self.new_worksheet_edit.setEnabled(state == 2)  # 2 is checked, 0 is unchecked

    def get_output_location(self):
        folder_dialog = QFileDialog()
        folder_dialog.setFileMode(QFileDialog.Directory)
        if folder_dialog.exec_():
            selected_folder = folder_dialog.selectedFiles()[0]
            self.output_location_edit.setText(selected_folder)

    def submit_form(self):
        excel_file = self.excel_file_edit.text()
        worksheet_name = self.worksheet_combo.currentText()
        translation_source = self.translation_source_edit.text()

        save_translations = self.save_translations_checkbox.isChecked()
        add_to_existing = self.add_to_existing_checkbox.isChecked()
        worksheet_to_add = self.worksheet_to_add_edit.text() if add_to_existing else None

        create_new_spreadsheet = self.create_new_spreadsheet_checkbox.isChecked()
        new_filename = self.new_filename_edit.text() if create_new_spreadsheet else None
        output_location = self.output_location_edit.text() if create_new_spreadsheet else None
        new_worksheet_name = self.new_worksheet_edit.text() if create_new_spreadsheet else None

        # Check that all relevant fields are filled up and show error messages if needed
        if not excel_file or not worksheet_name or not translation_source:
            self.show_message("Missing Fields", "Excel file, worksheet, and translation file fields cannot be blank.")
            return

        if not add_to_existing and not create_new_spreadsheet:
            self.show_message("Missing Fields", "Either 'Add to Existing Spreadsheet' or 'Create New Spreadsheet' must be checked.")
            return

        if create_new_spreadsheet and (not new_filename or not output_location or not new_worksheet_name):
            self.show_message("Missing Fields", "If 'Create New Spreadsheet' is checked, filename, location, and worksheet fields cannot be blank.")
            return

        if add_to_existing and not worksheet_to_add:
            self.show_message("Missing Fields", "If 'Add to Existing Spreadsheet' is checked, 'Worksheet to Add' field cannot be blank.")
            return

        print("Processing files...")

        # Use the values to calculate total sales
        total_sales(
            file_path=excel_file,
            sheet_name=worksheet_name,
            translation_sheet=translation_source,
            output_translations=save_translations,
            add_to=add_to_existing,
            worksheet_to_add=worksheet_to_add,
            create_new_spreadsheet=create_new_spreadsheet,
            new_filename=os.path.join(output_location, new_filename),
            new_worksheet=new_worksheet_name,
        )

        print("Finished processing!")

        # Show a message after processing files
        self.show_processed_message()

    def show_message(self, title, text):
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Warning)
        msg_box.setWindowTitle(title)
        msg_box.setText(text)
        msg_box.exec_()

    def show_processed_message(self):
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setWindowTitle("Files Processed")
        msg_box.setText("Files have been processed!")
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.finished.connect(self.clear_fields)
        msg_box.exec_()

    def clear_fields(self):
        # Clear all fields
        self.excel_file_edit.clear()
        self.worksheet_combo.clear()
        self.translation_source_edit.clear()
        self.save_translations_checkbox.setChecked(False)
        self.add_to_existing_checkbox.setChecked(False)
        self.worksheet_to_add_edit.clear()
        self.worksheet_to_add_edit.setEnabled(False)
        self.create_new_spreadsheet_checkbox.setChecked(True)  # Set back to default state
        self.new_filename_edit.clear()
        self.output_location_edit.clear()
        self.new_worksheet_edit.clear()

        # Set default values for specific fields
        search_results = ExcelForm.autofill_translation_source()
        self.translation_source_edit.setText(search_results if search_results else '')
        self.new_filename_edit.setText('Result.xlsx')
        self.worksheet_to_add_edit.setText('Monthly Sales Calculation')
        self.new_worksheet_edit.setText('Sheet1')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    excel_form = ExcelForm()
    excel_form.show()
    sys.exit(app.exec_())
