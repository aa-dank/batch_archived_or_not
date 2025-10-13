# Batch Archived or Not - Version 1.1.5
Stand-alone app for checking files to see if they have been archived in UCSC PPDO Construction Archives

## Overview
**Batch Archived or Not** is a Python-based desktop application designed to help users check whether files stored in a local directory are archived on a remote server. The app offers both recursive file search and selective reporting features, and allows users to save results in JSON or Excel formats.

## Features
- **File Checking**: Verifies whether files in a specified directory (and optionally sub-directories) exist on a remote server using an API.
- **File Size Limit**: Automatically skips files larger than 650MB to prevent timeout issues and excessive network usage.
- **Adaptive Timeouts**: Intelligent timeout handling that adjusts based on file size for optimal performance on network-mounted drives.
- **Recursive Search**: Option to enable searching through sub-directories.
- **Filtered Output**: Option to show only missing files in the output.
- **Custom Output Formats**: Allows exporting results in JSON, Excel, or both formats.
- **Progress Bar**: Visual feedback through a progress bar during processing.
- **GUI-based Interaction**: Simple file selection and output handling through a PySide6-based graphical interface.
- **Error Handling**: Displays detailed error messages if something goes wrong with file processing or API calls.

## Prerequisites
Before running the application, ensure you have the following installed:

- **Python**: Version 3.8 or higher.
- **Required Python Packages**:
  - `pandas`
  - `httpx`
  - `PySide6`
  
To install the necessary packages, run:
```bash
pip install pandas httpx PySide6
```

## How to Use
1. **File Path**: Input a valid directory path containing the files to check or use the Browse button to select a folder.
2. **Custom Output Path (Optional)**: Specify a custom path where the output files will be saved (or leave blank to save in the current directory).
3. **Output Format**: Select the desired output format (None, JSON, Excel, or both) from the dropdown menu.
4. **Recursive Search**: Check the "Recursive" option if you want to search within sub-directories.
5. **Only Missing Files**: Select the "Only show missing files" option to only display files that are not found on the server.
6. **Submit**: Press the Submit button to start the process.
7. **Progress**: The progress bar will update as files are processed, and detailed results or errors will be shown in the output text area.
8. **Save Output**: The results will be saved in the specified format, and the application will display the path to the saved files.

## File Export Options
- **JSON Export**: Results are saved in a structured JSON file showing the file paths and their locations on the server (if found).
- **Excel Export**: Results are saved in an Excel file with columns for the source file paths and their corresponding locations.
- **Both**: You can select both JSON and Excel output formats.

## Example Directory Structure
For example, if you are checking files in the directory `C:\MyFiles`, and the option to search recursively is enabled, the application will search through all sub-directories within `C:\MyFiles`.

## Application GUI
The application provides a simple, intuitive graphical interface where users can:

- Browse directories for input and output.
- Select options for recursive file search and selective reporting.
- Monitor progress and view results in real-time.
- Save the output to the desired location in JSON, Excel, or both formats.

## Installation and Setup
1. Clone the repository or download the project files.
2. Install the required dependencies:
```bash
pip install pandas httpx PySide6
```
3. Fill in the `creds.py` file in the project directory with the following content:
```python
APP_API_USERNAME = 'your_username'
APP_API_PASSWORD = 'your_password'
```
4. Run the application:
```bash
python app.py
```

## Notes
- **File Size Limit**: Files larger than 650MB are automatically skipped to prevent timeouts and network issues. These files are recorded in the results with a "Skipped" status indicating they exceeded the size limit. The limit can be adjusted by modifying the `MAX_FILE_SIZE_MB` constant in the source code.
- **Adaptive Timeouts**: The application uses intelligent timeout settings that scale with file size. Small files timeout faster, while larger files get more time to upload from network-mounted drives. Connection timeouts are kept short (10 seconds) while file upload timeouts adapt from 60 seconds up to 10 minutes for very large files.
- **SSL Certificate Handling**: The application uses httpx with SSL verification disabled for convenience. If you need to enforce certificate validation, modify the httpx.Client call in the code to verify=True.
- **Icon**: The application includes a window icon, which you can replace by updating the app_icon_.ico file in the root directory. This can be used when packaging the project into an application.
