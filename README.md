# Excel Parser

This Python script parses a Microsoft Excel (.xlsx) file and prints the contents of each sheet.

## Requirements

- Python 3
- openpyxl (available in Debian 12 apt repository)

## Installation

Install the required package using apt:

```bash
sudo apt update
sudo apt install python3-openpyxl
```

## Usage

Run the script with the path to your Excel file:

```bash
python3 parse_excel.py your_file.xlsx
```

The script will print the contents of each sheet in the Excel file. 