"# ALHashGUI" 
Excel Sheet Protection Data Extractor
Overview
This tool provides a simple GUI to extract and convert sheet protection data from Excel files (.xlsx). It is designed to assist in the analysis and handling of sheet protection data by converting it into a standardized format that can be easily used for further processing or auditing purposes.

Features
Excel File Selection: Users can select an Excel file from which to extract sheet protection data.
Output Directory Selection: Users can specify an output directory where the extracted data will be saved.
Data Conversion: Extracts sheet protection data and converts it into a standardized format.
GUI Interface: Easy to use graphical user interface for all operations.
Installation
No installation is required. The tool is a standalone Python script that utilizes tkinter for the GUI. However, ensure that Python is installed on your system.

Prerequisites
Python 3.x
tkinter (should be included with Python 3.x)
Usage
Run the script using Python:
bash
Copy code
python ALHashGUI.py
Click on the "Select Excel File" button to choose an Excel file (.xlsx) from which to extract sheet protection data.
Choose an output directory where the extracted and converted sheet protection data will be saved.
The tool will process the file and save the data in the specified output directory. A success message will be displayed upon completion.
How It Works
The script reads an Excel file (.xlsx), identifies the sheet protection data within each worksheet, and converts this data into a standardized format. This data includes the hash value, salt value, and spin count, among other attributes. The converted data is then saved to text files, one for each sheet, in a directory named after the Excel file.

Limitations
The tool currently supports Excel files (.xlsx) only.
It is designed to extract data specifically formatted for sheet protection; it does not extract other types of data from Excel files.
Contributing
Contributions to improve the tool or extend its functionality are welcome. Please feel free to fork the repository and submit pull requests.

License
This project is open source and available under the MIT License.
