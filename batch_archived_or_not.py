import sys
import os
import json
import pandas as pd
import requests.packages
import requests.api
from datetime import datetime
from PySide6 import QtGui
from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (QApplication, QTextEdit, QWidget, QVBoxLayout, QPushButton, QHBoxLayout, QLabel,
                               QFileDialog, QCheckBox, QLineEdit, QProgressBar, QComboBox)
from PySide6.QtCore import Qt
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from creds import APP_API_USERNAME, APP_API_PASSWORD

VERSION = "1.1.2"
URL_TEMPLATE = r"https://{}/api/archived_or_not?user={}&password={}"
# ADDRESS = r"localhost:5000" # for testing
ADDRESS = r"ppdo-dev-app-1.ucsc.edu"
basedir = os.path.dirname(__file__)

class HeavyLifter(QThread):
    progress = Signal(int)
    finished = Signal(str)
    error = Signal(str)

    def __init__(self, path, exclude_src, recursive, only_missing_files, output_type, custom_path, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.path = path
        self.exclude_src = exclude_src
        self.recursive = recursive
        self.only_missing_files = only_missing_files
        self.output_type = output_type
        self.custom_path = custom_path
        self.stop = False # new stop flag that resets progress

    def run(self):
        self.stop = False
        try:
            self.process_files()
        except Exception as e:
            self.error.emit(f"Error occurred: {str(e)}")

    def cancel(self):
        self.stop = True

    def process_files(self):
        file_ignore = [".DS_Store", "Thumbs.db"]
        results = {}
        progress_bar_counter = 0
        self.progress.emit(0)
        progress_bar_max = self.find_file_count()

        if progress_bar_max == 0:
            self.finished.emit("No files found.")
            return

        for root, dirs, files in os.walk(self.path):
            if self.stop:
                self.finished.emit("<br><b>Process canceled.</b>")
                self.progress.emit(100)
                return
            # iterate through files in directory
            for file in files:
                # skip hidden and temp files
                if self.stop:
                    self.finished.emit("<br><b>Process canceled.</b>")
                    self.progress.emit(100)
                    return
                if file in file_ignore or file.startswith("~$"):
                    continue

                # update progress bar
                progress_bar_counter += 1
                progress = int((progress_bar_counter * 100) // progress_bar_max)
                # avoid 0%
                if progress == 0:
                    progress = 1

                filepath = os.path.join(root, file)
                path_relative_to_files_location = os.path.relpath(filepath, self.path)
                request_url = URL_TEMPLATE.format(ADDRESS, APP_API_USERNAME, APP_API_PASSWORD)
                file_locations = []

                # open file and send to server endpoint
                try:
                    with open(filepath, 'rb') as f:
                        files = {'file': f}
                        response = requests.post(request_url, files=files, verify=False)
                        filepath = filepath.replace('/', '\\')

                        file_str = "Locations for {}".format(path_relative_to_files_location.replace('/', '\\'))
                        if not self.only_missing_files:
                            self.finished.emit("<br><b>{}</b>".format(file_str))
                        if response.status_code == 404:
                            if self.only_missing_files:
                                self.finished.emit("<br><b>{}</b>".format(file_str))
                            self.finished.emit("\n<pre>    None</pre>")
                            file_locations = "None"
                        else:
                            file_locations = json.loads(response.text)
                            if self.only_missing_files:
                                results[filepath] = file_locations
                            else:
                                for i in range(len(file_locations)):
                                    file_locations[i] = "N:\\PPDO\\Records\\{}".format(file_locations[i].replace('/', '\\'))
                                    self.finished.emit("<pre>    {}</pre>".format(file_locations[i]))
                                    if self.exclude_src and file_locations[i] == filepath:
                                        del file_locations[i]
                except Exception as e:
                    if 'response' in locals() and response.status_code in [404, 400, 500, 405]:
                        self.error.emit(f"Request Error:<br>{response.text}")
                        return
                self.progress.emit(progress)
                results[filepath] = file_locations

            if root == self.path and not self.recursive:
                break

        self.save_results(results)
        self.finished.emit("<br><b>Search complete.</b>")

    def find_file_count(self):
        file_count = 0
        file_ignore = {".DS_Store", "Thumbs.db"}

        self.finished.emit("<b>Calculating file count...</b>")
        for _, _, files in os.walk(self.path):
            for file in files:
                if file not in file_ignore and not file.startswith("~$"):
                    file_count += 1
            if not self.recursive:
                break
        self.finished.emit(f"<b>File count completed for {file_count} files.</b>")
        return file_count


    def save_results(self, results):
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        output_location = self.custom_path

        try:
            if self.output_type in ['json', 'json and excel']:
                if os.path.isdir(output_location):
                    path_name = json_export(results, timestamp, output_location)
                else:
                    path_name = json_export(results, timestamp, "default")
                self.finished.emit(f"<br><b>Results JSON file saved to:</b><br>{path_name}")

            if self.output_type in ['excel', 'json and excel']:
                if os.path.isdir(output_location):
                    path_name = excel_export(results, timestamp, output_location)
                else:
                    path_name = excel_export(results, timestamp, "default")
                self.finished.emit(f"<br><b>Results Excel file saved to:</b><br>{path_name}")

        except Exception as e:
            self.error.emit(f"Error: Can't export file to requested location. {str(e)}.")


class GuiHandler(QWidget):
    def __init__(self, app_version: str):
        super().__init__()
        self.layout = QVBoxLayout()
        self.app_version = app_version
        self.initUI()
        self.hl = None

    def initUI(self):
        self.setWindowTitle("Batch Archived or Not")
        self.layout.addWidget(QLabel("Version: " + self.app_version))

        # Input section with directory selection
        self.path_label_head = QLabel("Input a valid file path in box below. Copy and paste it from Windows File "
                                      "Explorer or use 'Browse' to locate a folder.", self)
        self.layout.addWidget(self.path_label_head)

        self.path_layout = QHBoxLayout()
        self.path_label = QLabel("Path to directory of files to check: ", self)
        self.path_layout.addWidget(self.path_label)

        self.path_line_edit = QLineEdit(self)
        self.path_layout.addWidget(self.path_line_edit)

        self.browse_button = QPushButton("Browse", self)
        self.browse_button.clicked.connect(self.browse_directory)
        self.path_layout.addWidget(self.browse_button)

        self.layout.addLayout(self.path_layout)  # Added path_layout to main layout

        # Optional custom save path section and dropdown
        self.custom_path_head = QLabel("Optional: Input an output path to save excel/json to or use 'Browse' to "
                                       "locate a folder, then select a format.", self)
        self.layout.addWidget(self.custom_path_head)

        self.custom_path_layout = QHBoxLayout()  # Added QHBoxLayout for custom path selection
        self.custom_path_label = QLabel("Path to directory to save output in:", self)
        self.custom_path_layout.addWidget(self.custom_path_label)

        self.custom_path_line_edit = QLineEdit(self)
        self.custom_path_layout.addWidget(self.custom_path_line_edit)

        self.custom_browse_button = QPushButton("Browse", self)
        self.custom_browse_button.clicked.connect(self.browse_custom_path)
        self.custom_path_layout.addWidget(self.custom_browse_button)

        self.save_combo_box = QComboBox(self)
        self.save_combo_box.addItem("none")
        self.save_combo_box.addItem("json")
        self.save_combo_box.addItem("excel")
        self.save_combo_box.addItem("json and excel")
        self.custom_path_layout.addWidget(self.save_combo_box)

        self.layout.addLayout(self.custom_path_layout)  # Added custom_path_layout to main layout

        # Checkboxes
        self.recursive_box = QCheckBox("Should file checking be recursive through nested sub-directories?", self)
        self.layout.addWidget(self.recursive_box)

        self.missing_box = QCheckBox("Only show files that are not found on the server? Useful for reducing the output from this tool (won't effect excel or json output)", self)
        self.layout.addWidget(self.missing_box)

        self.exclude_source_box = QCheckBox("ONLY FOR JSON/EXCEL OUTPUT: Exclude source path for each file? helpful when looking for files other occurences.", self)
        self.layout.addWidget(self.exclude_source_box)

        # Submit button
        self.submit_layout = QHBoxLayout()  # Added QHBoxLayout for submit button and progress bar
        self.submit_button = QPushButton("Submit", self)
        self.submit_button.clicked.connect(self.archived_or_not_call)
        self.submit_layout.addWidget(self.submit_button)

        # Cancel button
        self.cancel_button = QPushButton("Cancel", self)
        self.cancel_button.clicked.connect(self.cancel_heavylifter)
        self.cancel_button.setEnabled(False)
        self.submit_layout.addWidget(self.cancel_button)

        # Progress bar
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setTextVisible(True)  # Show percentage text
        self.submit_layout.addWidget(self.progress_bar)

        self.layout.addLayout(self.submit_layout)  # Added submit_layout to main layout

        # Output section
        self.output_text_edit = QTextEdit(self)
        self.output_text_edit.setReadOnly(True)
        self.layout.addWidget(self.output_text_edit)

        # Exit button
        self.exit_button = QPushButton("Exit", self)
        self.exit_button.clicked.connect(self.close)
        self.layout.addWidget(self.exit_button)

        self.setLayout(self.layout)

    def browse_directory(self):
        directory_path = QFileDialog.getExistingDirectory(self, "Select Directory")
        if directory_path:
            self.path_line_edit.setText(directory_path)

    def browse_custom_path(self):
        custom_directory_path = QFileDialog.getExistingDirectory(self, "Select Custom Directory")
        if custom_directory_path:
            self.custom_path_line_edit.setText(custom_directory_path)

    def archived_or_not_call(self):
        self.output_text_edit.clear()

        exclude_src = self.exclude_source_box.isChecked()
        recursive = self.recursive_box.isChecked()
        only_missing_files = self.missing_box.isChecked()
        output_type = self.save_combo_box.currentText()
        custom_path = self.custom_path_line_edit.text().strip()
        files_location = self.path_line_edit.text().strip()

        if not os.path.isdir(files_location):
            self.output_text_edit.append("Must input valid filepath")
            return

        self.hl = HeavyLifter(files_location, exclude_src, recursive, only_missing_files, output_type, custom_path)
        self.hl.progress.connect(self.progress_bar.setValue)
        self.hl.finished.connect(self.handle_finished)
        self.hl.error.connect(self.output_text_edit.append)
        self.cancel_button.setEnabled(True)
        self.hl.start()

    def cancel_heavylifter(self):
        if self.hl and self.hl.isRunning():
            self.hl.cancel()
            self.cancel_button.setEnabled(False)

    def handle_finished(self, message):
        self.output_text_edit.append(message)

def json_export(r, time, custom_directory_path):
    if custom_directory_path == "default":
        results_filepath = os.path.join(os.getcwd(), f'archived_or_not_results_{time}.json')
    else:
        results_filepath = os.path.join(custom_directory_path, f'archived_or_not_results_{time}.json')
    results_filepath = results_filepath.replace("/", "\\")
    with open(results_filepath, 'w') as f:
        json.dump(r, f, indent=4)
    return results_filepath

def excel_export(r, time, custom_directory_path):
    if custom_directory_path == "default":
        results_filepath = os.path.join(os.getcwd(), f'archived_or_not_results_{time}.xlsx')
    else:
        results_filepath = os.path.join(custom_directory_path, f'archived_or_not_results_{time}.xlsx')
    results_filepath = results_filepath.replace("/", "\\")
    df = pd.DataFrame(columns=["Source Path", "Found Locations"])
    for key, vals in r.items():
        if vals == "None":
            df.loc[len(df.index)] = [key, vals]
            continue
        for val in vals:
            df.loc[len(df.index)] = [key, val]
    with pd.ExcelWriter(results_filepath, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return results_filepath




def main():
    # Disable SSL warnings
    requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

    app = QApplication(sys.argv)
    app.setWindowIcon(QtGui.QIcon(os.path.join(basedir, 'app_icon_.ico')))
    gui = GuiHandler(app_version=VERSION)
    gui.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
