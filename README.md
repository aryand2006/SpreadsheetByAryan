# PySpreadsheet

A powerful spreadsheet application with Excel and Google Sheets-like features, built using Python and PyQt5.

## Features

- Modern and intuitive user interface
- Multi-sheet workbook support
- Formula calculation engine
- Cell formatting options
- Data visualization with interactive charts
- Conditional formatting
- Import and export to various file formats
- Find and replace functionality
- Customizable preferences
- Dark and light theme support

## Requirements

- Python 3.6+
- PyQt5
- openpyxl (for Excel file support)

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/aryand2006/PySpreadsheet.git
   ```

2. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

3. Run the application:
   ```
   python src/main.py
   ```

## Usage Guide

### Basic Operations

- **Creating a New Spreadsheet**: File > New or Ctrl+N
- **Opening a Spreadsheet**: File > Open or Ctrl+O
- **Saving a Spreadsheet**: File > Save or Ctrl+S
- **Exporting to PDF**: File > Export to PDF

### Working with Sheets

- **Add Sheet**: Sheet > Insert Sheet
- **Rename Sheet**: Sheet > Rename Sheet
- **Delete Sheet**: Sheet > Delete Sheet
- **Switch between Sheets**: Click on the sheet tabs at the bottom

### Cell Editing

- **Select Cell**: Click on the cell
- **Edit Cell**: Double-click or press F2
- **Enter Formula**: Start with = (e.g., =SUM(A1:A5))
- **Cut/Copy/Paste**: Edit menu or Ctrl+X/C/V

### Formatting

- **Change Font**: Format > Font
- **Change Cell Color**: Format > Cell Color
- **Change Text Color**: Format > Text Color
- **Apply Conditional Formatting**: Format > Conditional Formatting

### Data Operations

- **Sort Data**: Select data, then Data > Sort Ascending/Descending
- **Filter Data**: Select data, then Data > Filter
- **Find and Replace**: Edit > Find or Edit > Replace

### Charts

- **Create Chart**: Select data, then Insert > Chart
- **Customize Chart**: Use the chart options panel
- **Save Chart as Image**: Use the "Save Chart as Image" button

### Customizing the Application

- **Change Theme**: Tools > Preferences > Appearance tab
- **Set Default Font**: Tools > Preferences > Appearance tab
- **Auto-save Settings**: Tools > Preferences > General tab

## Supported Formulas

- SUM(range): Sum of values in the range
- AVERAGE(range): Average of values in the range
- COUNT(range): Count of items in the range
- MIN(range): Minimum value in the range
- MAX(range): Maximum value in the range
- PRODUCT(range): Product of all values in the range

## Credits

- Created by Aryan Daga
- Built using PyQt5
- Documentation and feature suggestions by GitHub Copilot