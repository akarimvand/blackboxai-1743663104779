# VCF to Excel Converter

A Flask web application that converts VCF (vCard) contact files to Excel format.

## Features
- Upload VCF files and convert them to Excel spreadsheets
- Preserves all contact fields including:
  - Names, organizations, titles
  - Multiple phone numbers (each in separate columns)
  - Multiple email addresses  
  - Addresses, notes, URLs, birthdays
- Clean responsive UI with Tailwind CSS
- Error handling for invalid files

## Requirements
- Python 3.7+
- Flask
- pandas
- vobject
- openpyxl

## Installation
1. Clone the repository
2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage
1. Run the application:
```bash
python app.py
```
2. Open http://localhost:8000 in your browser
3. Upload a VCF file
4. Download the converted Excel file

## Screenshot
![Screenshot of the converter interface](screenshot.png)

## License
MIT
