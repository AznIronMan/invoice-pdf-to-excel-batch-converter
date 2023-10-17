# PDF to Excel Batch Converter
- **Version**: v1.10
- **Date**: 10.17.2023 @ 10:00 PST
- **Developed for**: Casey Ryan (cryan@swfirefightingfoam.com) of SW Firefighting Foam & Equipment, LLC

## Overview
This Python script is designed to batch-convert PDF files to Excel spreadsheets. The code is well-structured and modular, with specific functionalities encapsulated in specialized functions. It was written specifically for the structued PDFs provided by SW Firefighting Foam & Equipment, LLC, but can be easily adapted for other PDFs.

## Modules

filer.py

- **check_function**: Checks if a file or directory exists. Can create directories.
- **dir_check**: Wrapper around check_function for directories.
- **file_exists**: Wrapper around check_function for files.
- **is_gui_available**: Checks if a GUI environment is available.
- **path_to_module**: Converts a file path to a Python module path.
- **select_folder**: Opens a file dialog for folder selection or takes user input.

logger.py

- **check_function**: Checks if a file or directory exists. Can create directories. (Duplicated from filer.py)
- **dir_check**: Wrapper around check_function for directories. (Duplicated from filer.py)
- **fix_datetime**: Formats datetime objects or timestamps.
- **format_log_date**: Formats date for log files.
- **log**: Writes log messages to a file.
- **now**: Gets the current time in a formatted string.
- **today**: Gets the current date in a formatted string.
- **zlog**: Enhanced logging function with console output option.

process.py

- **batch_convert**: Batch converts PDFs to Excel in a target directory.
- **clean_currency**: Cleans currency strings.
- **find_and_parse_date**: Finds and parses dates in text.
- **format_excel**: Formats the generated Excel file.
- **get_years_to_search**: Returns a list of years to search for in text.
- **map_text_to_excel_columns**: Maps extracted text to Excel columns.
- **pdf_to_excel**: Core function that manages the PDF to Excel conversion.
- **_Various parse_ functions**: Extract specific information from text.

## Requirements
- Python 3.8+
- openpyxl
- pdfplumber
- pandas
- tkinter (for GUI dialogs)
- re (Regular expressions)
- datetime
- dateutil
- dotenv (For environment variables)

## Usage
- Ensure all dependencies are installed. Run `pip install -r requirements.txt` to install all dependencies.
- Run `python \.` from the root of the project where __main__.py is located.
- If a GUI environment is available, a file dialog will open for folder selection. Otherwise, the user will be prompted to enter a folder path.
- All processed PDFs will output as Excel files in a new 'processed' directory within the same directory as the PDFs.
- Subdirectories PDF files will be converted to Excel files within the same subdirectory in a new 'processed' subdirectory.

## Areas for Improvement
- **Exception Handling**: Could be improved for more specific error messages.

## Author Information
- **Author**: [Geoff Clark of ClarkTribeGames, LLC](https://clarktribegames.com)
- **Email**:  [geoff@clarktribegames.com](mailto:geoff@clarktribegames.com)
- **Socials**:
    [Github @aznironman](https://github.com/aznironman)
    [IG: @aznironman](https://instagram.com/aznironman)
    [X: @aznironman](https://www.twitter.com/aznironman)
