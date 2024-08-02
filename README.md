# XML to Excel Converter

A simple and intuitive GUI application to convert XML files into Excel format. This tool is designed to help users easily select multiple XML files and convert each into a separate Excel file.

## Features

- Convert multiple XML files to Excel files in one go.
- User-friendly graphical interface built with Tkinter.
- Error handling to provide feedback in case of conversion issues.
- Modern and attractive design.

## Prerequisites

Make sure you have the following Python libraries installed:

- pandas
- openpyxl
- tkinter (comes with Python standard library)
- Pillow

You can install the required libraries using pip:

```bash
pip install pandas openpyxl pillow

Installation
1. Clone the repository:

```bash git clone https://github.com/your-username/xml-to-excel-converter.git

```bash
cd xml-to-excel-converter

```bash
python xml_to_excel_converter.py

Creating an Executable
To create an executable file for easy distribution, you can use PyInstaller:

Install PyInstaller:

```bash
pip install pyinstaller

Create the executable:

```bash
pyinstaller --onefile --windowed xml_to_excel_converter.py
