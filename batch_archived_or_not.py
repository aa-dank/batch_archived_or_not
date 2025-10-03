import sys
import os
import json
import pandas as pd
import httpx
from datetime import datetime
from PySide6 import QtGui
from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (QApplication, QTextEdit, QWidget, QVBoxLayout, QPushButton, QHBoxLayout, QLabel,
                               QFileDialog, QCheckBox, QLineEdit, QProgressBar, QComboBox)
from PySide6.QtCore import Qt
from creds import APP_API_USERNAME, APP_API_PASSWORD

VERSION = "1.1.4"
URL_TEMPLATE = r"https://{}/api/archived_or_not"
# ADDRESS = r"localhost:5000" # for testing
ADDRESS = r"ppdo-dev-app-1.ucsc.edu"
basedir = os.path.dirname(__file__)

headers = {"user": APP_API_USERNAME, "password": APP_API_PASSWORD}

class HeavyLifter(QThread):
    """
    A QThread subclass that handles the heavy lifting of file processing and HTTP API calls.
    
    This class runs in a separate thread to prevent the GUI from freezing while processing
    files. It walks through directory structures, sends files to a remote API via HTTPX
    to check if they are archived, and reports progress and results back to the main thread.
    
    Signals:
        progress (int): Emitted to update the progress bar (0-100)
        finished (str): Emitted with status messages for the UI
        error (str): Emitted when errors occur during processing
        
    Attributes:
        files_to_ignore (list): List of system files to skip during processing
    """
    progress = Signal(int)
    finished = Signal(str)
    error = Signal(str)
    files_to_ignore = [".DS_Store", "Thumbs.db"]

    def __init__(self, path, exclude_src, recursive, only_missing_files, output_type, custom_path, *args, **kwargs):
        """
        Initialize the HeavyLifter thread with processing parameters.
        
        Args:
            path (str): The directory path to process files from
            exclude_src (bool): Whether to exclude source paths from results
            recursive (bool): Whether to search subdirectories recursively
            only_missing_files (bool): Whether to show only files not found on server
            output_type (str): Output format ("none", "json", "excel", or "json and excel")
            custom_path (str): Custom path for saving output files
            *args: Additional positional arguments for QThread
            **kwargs: Additional keyword arguments for QThread
        """
        super().__init__(*args, **kwargs)
        self.path = path
        self.exclude_src = exclude_src
        self.recursive = recursive
        self.only_missing_files = only_missing_files
        self.output_type = output_type
        self.custom_path = custom_path
        self.stop = False # new stop flag that resets progress

    def run(self):
        """
        Main thread execution method that runs the file processing workflow.
        
        This method is called when the thread starts. It resets the stop flag,
        calls process_files() to handle the actual work, and emits error signals
        if any exceptions occur during processing.
        
        Raises:
            Exception: Any exception during file processing is caught and emitted as an error signal
        """
        self.stop = False
        try:
            self.process_files()
        except Exception as e:
            self.error.emit(f"Error occurred: {str(e)}")

    def cancel(self):
        """
        Cancel the current file processing operation.
        
        Sets the stop flag to True, which will cause the processing loop
        to exit gracefully on the next iteration.
        """
        self.stop = True

    def ignore_file(self, filename):
        """
        Check if a file should be ignored based on filename.
        
        Args:
            filename (str): The name of the file to check
            
        Returns:
            bool: True if the file should be ignored, False otherwise
            
        Note:
            Files are ignored if they are in the files_to_ignore list
            or if they start with "~$" (temporary files)
        """
        return filename in self.files_to_ignore or filename.startswith("~$")

    def update_progress(self, current_count, total_count):
        """
        Update progress bar with current file processing count.
        
        Args:
            current_count (int): Number of files processed so far
            total_count (int): Total number of files to process
            
        Note:
            Calculates percentage and ensures minimum 1% to avoid showing 0%
            Emits progress signal to update the GUI progress bar
        """
        if total_count == 0:
            progress = 0
        else:
            progress = int((current_count * 100) // total_count)
            # avoid 0%
            if progress == 0:
                progress = 1
        self.progress.emit(progress)

    def process_files(self):
        """
        Process all files in the specified directory and check if they exist on the remote server.
        
        This method walks through the directory structure, filters out ignored files,
        sends each valid file to the remote API endpoint to check if it's archived,
        and collects the results. Progress is updated throughout the process.
        
        The method handles API responses and formats the results based on the
        configured options (only_missing_files, exclude_src, etc.).
        
        Side Effects:
            - Updates progress bar through signals
            - Emits status messages to the GUI
            - Saves results when processing is complete
        """
        results = {}
        progress_bar_counter = 0
        self.progress.emit(0)
        progress_bar_max = self.find_file_count()

        if progress_bar_max == 0:
            self.finished.emit("No files found.")
            return

        for root, _, files in os.walk(self.path):
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
                
                # update progress bar
                progress_bar_counter += 1
                self.update_progress(progress_bar_counter, progress_bar_max)

                if self.ignore_file(file):
                    continue

                filepath = os.path.join(root, file)
                path_relative_to_files_location = os.path.relpath(filepath, self.path)
                request_url = URL_TEMPLATE.format(ADDRESS)
                file_locations = []

                # open file and send to server endpoint
                try:
                    with open(filepath, 'rb') as f:
                        files = {'file': f}
                        with httpx.Client(verify=False) as client:
                            response = client.post(request_url, headers=headers, files=files)
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
                            if not self.only_missing_files:
                                for i in range(len(file_locations)):
                                    file_locations[i] = "N:\\PPDO\\Records\\{}".format(file_locations[i].replace('/', '\\'))
                                    self.finished.emit("<pre>    {}</pre>".format(file_locations[i]))
                                    if self.exclude_src and file_locations[i] == filepath:
                                        del file_locations[i]
                except Exception as e:
                    if 'response' in locals() and response.status_code in [404, 400, 500, 405]:
                        self.error.emit(f"Request Error:<br>{response.text}")
                    else:
                        self.error.emit(f"Error processing file {filepath}: {str(e)}")
                    results[filepath] = "Error"
                    continue
                
                results[filepath] = file_locations

            if root == self.path and not self.recursive:
                break

        self.save_results(results)
        self.finished.emit("<br><b>Search complete.</b>")

    def find_file_count(self):
        """
        Count the total number of valid files in the directory to be processed.
        
        This method walks through the directory structure and counts all files
        that are not in the ignore list. The count is used to calculate progress
        percentage during file processing.
        
        Returns:
            int: Total number of files that will be processed (excluding ignored files)
            
        Side Effects:
            - Emits status messages to the GUI about the counting process
            - Respects the recursive setting when walking directories
        """
        file_count = 0

        self.finished.emit("<b>Calculating file count...</b>")
        for _, _, files in os.walk(self.path):
            for file in files:
                if not self.ignore_file(file):
                    file_count += 1
            if not self.recursive:
                break
        self.finished.emit(f"<b>File count completed for {file_count} files.</b>")
        return file_count


    def save_results(self, results):
        """
        Save the processing results to files in the specified output format(s).
        
        This method handles saving results in JSON and/or Excel formats based on
        the output_type configuration. It uses a timestamp to create unique filenames
        and respects the custom output path setting.
        
        Args:
            results (dict): Dictionary containing file paths as keys and their
                          archive locations (or "None"/"Error") as values
                          
        Side Effects:
            - Creates output files in the specified directory
            - Emits status messages to the GUI about save locations
            - Emits error messages if file saving fails
        """
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
    """
    A QWidget subclass that provides the graphical user interface for the application.
    
    This class creates and manages the GUI components including input fields, checkboxes,
    buttons, progress bar, and output display. It handles user interactions and coordinates
    with the HeavyLifter thread to process files while keeping the interface responsive.
    
    Attributes:
        layout (QVBoxLayout): Main vertical layout container for all GUI elements
        app_version (str): Version string displayed in the GUI
        hl (HeavyLifter): Reference to the background processing thread
        
    GUI Components:
        - Directory path input with browse button
        - Custom output path selection
        - Output format dropdown (none/json/excel/both)
        - Recursive search checkbox
        - Show only missing files checkbox
        - Exclude source paths checkbox (for exports)
        - Submit and Cancel buttons
        - Progress bar
        - Output text display area
    """
    def __init__(self, app_version: str):
        """
        Initialize the GuiHandler with the application version.
        
        Sets up the main layout, stores the version string, initializes the UI,
        and prepares the HeavyLifter thread reference.
        
        Args:
            app_version (str): Version string to display in the GUI
        """
        super().__init__()
        self.layout = QVBoxLayout()
        self.app_version = app_version
        self.initUI()
        self.hl = None

    def initUI(self):
        """
        Initialize and configure all GUI components.
        
        Creates and arranges all the user interface elements including:
        - Window title and version display
        - Directory path input with browse button
        - Custom output path selection
        - Output format dropdown
        - Configuration checkboxes (recursive, missing files only, exclude source)
        - Submit and cancel buttons with progress bar
        - Output text display area
        - Exit button
        
        Side Effects:
            - Sets up all widget layouts and connections
            - Configures event handlers for buttons
            - Sets initial widget states and properties
        """
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

        self.exclude_source_box = QCheckBox("ONLY FOR JSON/EXCEL: Exclude the source path for each file. Helpful when looking for files that are already on the archives file server other occurences.", self)
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
        """
        Open a directory selection dialog for choosing the input directory.
        
        Opens a file dialog allowing the user to select a directory containing
        files to be processed. Updates the path input field with the selected directory.
        
        Side Effects:
            - Opens a QFileDialog for directory selection
            - Updates the path_line_edit widget with the selected directory path
        """
        directory_path = QFileDialog.getExistingDirectory(self, "Select Directory")
        if directory_path:
            self.path_line_edit.setText(directory_path)

    def browse_custom_path(self):
        """
        Open a directory selection dialog for choosing the custom output directory.
        
        Opens a file dialog allowing the user to select a directory where output
        files (JSON/Excel) should be saved. Updates the custom path input field.
        
        Side Effects:
            - Opens a QFileDialog for directory selection
            - Updates the custom_path_line_edit widget with the selected directory path
        """
        custom_directory_path = QFileDialog.getExistingDirectory(self, "Select Custom Directory")
        if custom_directory_path:
            self.custom_path_line_edit.setText(custom_directory_path)

    def archived_or_not_call(self):
        """
        Initiate the file processing operation based on current GUI settings.
        
        Collects all configuration options from the GUI, validates the input directory,
        creates and configures a HeavyLifter thread, and starts the background processing.
        
        Configuration collected:
            - exclude_src: Whether to exclude source paths from results
            - recursive: Whether to search subdirectories
            - only_missing_files: Whether to show only missing files
            - output_type: Format for saving results (none/json/excel/both)
            - custom_path: Directory for saving output files
            - files_location: Source directory to process
            
        Side Effects:
            - Clears the output text display
            - Validates the input directory path
            - Creates and starts a HeavyLifter thread
            - Connects thread signals to GUI update methods
            - Enables the cancel button
        """
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
        """
        Cancel the currently running file processing operation.
        
        Checks if a HeavyLifter thread is running and calls its cancel method
        to gracefully stop the processing. Disables the cancel button after use.
        
        Side Effects:
            - Calls cancel() on the HeavyLifter thread if it's running
            - Disables the cancel button
        """
        if self.hl and self.hl.isRunning():
            self.hl.cancel()
            self.cancel_button.setEnabled(False)

    def handle_finished(self, message):
        """
        Handle status messages from the HeavyLifter thread.
        
        This slot is connected to the HeavyLifter's finished signal and receives
        status messages throughout the processing operation. Messages are displayed
        in the output text area to keep the user informed of progress.
        
        Args:
            message (str): Status message from the HeavyLifter thread, typically
                         containing HTML formatting for display in the text widget
        """
        self.output_text_edit.append(message)

def json_export(r, time, custom_directory_path):
    """
    Export processing results to a JSON file.
    
    Creates a JSON file containing the file processing results with a timestamp
    in the filename. The file is saved either in the specified custom directory
    or in the current working directory.
    
    Args:
        r (dict): Results dictionary with file paths as keys and archive locations as values
        time (str): Timestamp string for unique filename generation (format: YYYY-MM-DD_HH-MM-SS)
        custom_directory_path (str): Directory path for saving the file, or "default" for current directory
        
    Returns:
        str: Full file path where the JSON results were saved
        
    Side Effects:
        - Creates a JSON file on disk
        - Normalizes path separators to Windows format
    """
    if custom_directory_path == "default":
        results_filepath = os.path.join(os.getcwd(), f'archived_or_not_results_{time}.json')
    else:
        results_filepath = os.path.join(custom_directory_path, f'archived_or_not_results_{time}.json')
    results_filepath = results_filepath.replace("/", "\\")
    with open(results_filepath, 'w') as f:
        json.dump(r, f, indent=4)
    return results_filepath

def excel_export(r, time, custom_directory_path):
    """
    Export processing results to an Excel file.
    
    Creates an Excel file containing the file processing results with a timestamp
    in the filename. The results are formatted in a two-column table with source
    paths and their corresponding found locations.
    
    Args:
        r (dict): Results dictionary with file paths as keys and archive locations as values.
                 Values can be lists of locations, "None" for missing files, or "Error" for failed processing.
        time (str): Timestamp string for unique filename generation (format: YYYY-MM-DD_HH-MM-SS)
        custom_directory_path (str): Directory path for saving the file, or "default" for current directory
        
    Returns:
        str: Full file path where the Excel results were saved
        
    Side Effects:
        - Creates an Excel file on disk using pandas and openpyxl
        - Normalizes path separators to Windows format
        - Each source file gets multiple rows if it has multiple found locations
    """
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
    """
    Main application entry point that initializes and runs the GUI application.
    
    Sets up the QApplication, configures the window icon, creates the GUI handler,
    and starts the main event loop. The application will run until the user exits.
    
    Side Effects:
        - Creates QApplication instance
        - Sets window icon from app_icon_.ico file
        - Creates and displays the main GUI window
        - Starts the Qt event loop
        - Exits the application when loop ends
    """
    app = QApplication(sys.argv)
    app.setWindowIcon(QtGui.QIcon(os.path.join(basedir, 'app_icon_.ico')))
    gui = GuiHandler(app_version=VERSION)
    gui.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
