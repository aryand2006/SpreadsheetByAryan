# SpreadsheetByAryan

A powerful spreadsheet application with Excel and Google Sheets-like features, built using Python and PyQt5.

## Features

### Core Functionality
- Modern and intuitive user interface with tab-based document management
- Multi-sheet workbook support with easy navigation
- Comprehensive formula calculation engine with real-time evaluation
- Advanced cell formatting options (fonts, colors, alignments)
- Data visualization with interactive charts (bar, line, pie, scatter)
- Conditional formatting with multiple rule types and customizable styles
- Import and export to various file formats (CSV, XLSX, PDF)

### Data Management
- Cell merging and splitting
- Row and column operations (insert, delete, resize)
- Cell comments and annotations
- Named ranges for easier formula references
- Data sorting and filtering capabilities
- Find and replace with advanced options
- Auto-fill and pattern recognition for sequences

### User Experience
- Undo/redo functionality with comprehensive history
- Customizable keyboard shortcuts
- Zoom in/out for better visibility
- Auto-saving and recovery
- Dark and light theme support
- Customizable UI preferences
- Cell protection and worksheet locking

## Requirements

- Python 3.6+
- PyQt5
- openpyxl (for Excel file support)
- matplotlib (for chart visualization)
- numpy (for numerical operations)
- pandas (for data analysis functions)

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/aryand2006/SpreadsheetByAryan.git
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
- **Saving As**: File > Save As or Ctrl+Shift+S
- **Exporting to PDF**: File > Export to PDF
- **Print**: File > Print or Ctrl+P
- **Exit Application**: File > Exit or Alt+F4

### Working with Sheets

- **Add Sheet**: Sheet > Insert Sheet or + button on tab bar
- **Rename Sheet**: Double-click on sheet tab or Sheet > Rename Sheet
- **Delete Sheet**: Sheet > Delete Sheet or right-click on tab > Delete
- **Switch between Sheets**: Click on the sheet tabs at the bottom
- **Reorder Sheets**: Drag and drop sheet tabs
- **Copy Sheet**: Sheet > Copy Sheet

### Cell Editing

- **Select Cell**: Click on the cell
- **Select Range**: Click and drag across cells
- **Select Row/Column**: Click on row/column header
- **Select All**: Click top-left corner or Ctrl+A
- **Edit Cell**: Double-click or press F2
- **Enter Formula**: Start with = (e.g., =SUM(A1:A5))
- **Complete Formula**: Tab key when typing formula
- **Cut/Copy/Paste**: Edit menu, context menu, or Ctrl+X/C/V
- **Fill Down/Right**: Ctrl+D/Ctrl+R or drag fill handle

### Navigation

- **Move Selection**: Arrow keys
- **Move to Beginning of Row**: Home
- **Move to End of Row**: End
- **Move to Top of Column**: Ctrl+Home
- **Move to Bottom of Column**: Ctrl+End
- **Next Sheet**: Ctrl+Page Down
- **Previous Sheet**: Ctrl+Page Up
- **Go to Cell**: Ctrl+G

### Formatting

- **Bold/Italic/Underline**: Format menu or Ctrl+B/I/U
- **Change Font**: Format > Font or context menu
- **Change Cell Color**: Format > Cell Color
- **Change Text Color**: Format > Text Color
- **Borders**: Format > Borders
- **Number Formats**: Format > Number Format (currency, percentage, date, etc.)
- **Alignment**: Format > Alignment (left, center, right, top, middle, bottom)
- **Conditional Formatting**: Format > Conditional Formatting
- **Merge Cells**: Format > Merge Cells
- **Cell Styles**: Format > Cell Styles for predefined combinations

### Data Operations

- **Sort Data**: Select data, then Data > Sort Ascending/Descending
- **Filter Data**: Select data, then Data > Filter
- **Remove Duplicates**: Data > Remove Duplicates
- **Data Validation**: Data > Validation
- **Text to Columns**: Data > Text to Columns
- **Group/Ungroup**: Data > Group/Ungroup
- **Find and Replace**: Edit > Find/Replace or Ctrl+F/H
- **Auto-fill Series**: Drag fill handle to create patterns (numbers, dates, etc.)

### Charts and Visualizations

- **Create Chart**: Select data, then Insert > Chart
- **Chart Types**: Bar, Line, Pie, Scatter, Area, Doughnut, Radar
- **Customize Chart**: Use the chart options panel for titles, legends, axes
- **Chart Styling**: Change colors, patterns, borders
- **Trendlines**: Add statistical trendlines to data series
- **Save Chart as Image**: Use the "Save Chart as Image" button
- **Move/Resize Chart**: Drag chart or resize handles
- **Update Chart Data**: Automatically updates when source data changes

### Advanced Features

- **Freeze Panes**: View > Freeze Panes to lock rows/columns
- **Split View**: View > Split to create multiple window panes
- **Custom Number Formats**: Create your own number display formats
- **Named Ranges**: Define and use named cell ranges in formulas
- **Comments**: Insert > Comment to add notes to cells
- **Hyperlinks**: Insert > Hyperlink to add links
- **Headers and Footers**: Insert > Header & Footer for printed pages
- **Page Layout**: View > Page Layout for print preview

### Customizing the Application

- **Change Theme**: Tools > Preferences > Appearance tab
- **Set Default Font**: Tools > Preferences > Appearance tab
- **Auto-save Settings**: Tools > Preferences > General tab
- **Customize Ribbon**: Tools > Customize Ribbon
- **Keyboard Shortcuts**: Tools > Customize Shortcuts
- **Language Settings**: Tools > Language

## Supported Formulas

### Mathematical Functions
- `SUM(range)`: Sum of values in the range
- `AVERAGE(range)`: Average of values in the range
- `COUNT(range)`: Count of items in the range
- `MIN(range)`: Minimum value in the range
- `MAX(range)`: Maximum value in the range
- `PRODUCT(range)`: Product of all values in the range
- `SQRT(value)`: Square root of value
- `ABS(value)`: Absolute value
- `ROUND(value, decimals)`: Round to specified decimal places
- `ROUNDUP(value, decimals)`: Round up to specified decimal places
- `ROUNDDOWN(value, decimals)`: Round down to specified decimal places
- `INT(value)`: Integer part of value
- `MOD(number, divisor)`: Remainder after division

### Statistical Functions
- `STDEV(range)`: Standard deviation of a sample
- `STDEVP(range)`: Standard deviation of a population
- `VAR(range)`: Variance of a sample
- `VARP(range)`: Variance of a population
- `MEDIAN(range)`: Median value in range
- `MODE(range)`: Most common value in range
- `PERCENTILE(range, k)`: k-th percentile of values in range
- `CORREL(range1, range2)`: Correlation coefficient

### Logical Functions
- `IF(condition, value_if_true, value_if_false)`: Conditional evaluation
- `AND(condition1, condition2, ...)`: True if all conditions are true
- `OR(condition1, condition2, ...)`: True if any condition is true
- `NOT(condition)`: Reverses logical value
- `ISBLANK(value)`: True if value is blank
- `ISNUMBER(value)`: True if value is a number
- `ISTEXT(value)`: True if value is text

### Text Functions
- `LEFT(text, num_chars)`: Extract characters from left
- `RIGHT(text, num_chars)`: Extract characters from right
- `MID(text, start_num, num_chars)`: Extract characters from middle
- `LEN(text)`: Number of characters in text
- `CONCATENATE(text1, text2, ...)`: Join text strings
- `LOWER(text)`: Convert to lowercase
- `UPPER(text)`: Convert to uppercase
- `PROPER(text)`: Capitalize first letter of each word
- `TRIM(text)`: Remove extra spaces

### Date and Time Functions
- `NOW()`: Current date and time
- `TODAY()`: Current date
- `DATE(year, month, day)`: Create a date
- `DAY(date)`: Extract day from date
- `MONTH(date)`: Extract month from date
- `YEAR(date)`: Extract year from date
- `WEEKDAY(date)`: Day of week (1-7)
- `DATEDIF(start_date, end_date, unit)`: Difference between dates

### Lookup and Reference Functions
- `VLOOKUP(lookup_value, table_array, col_index, [range_lookup])`: Vertical lookup
- `HLOOKUP(lookup_value, table_array, row_index, [range_lookup])`: Horizontal lookup
- `INDEX(array, row_num, [column_num])`: Return value at position in array
- `MATCH(lookup_value, lookup_array, [match_type])`: Position of value in array
- `CHOOSE(index_num, value1, value2, ...)`: Choose value from list based on index

### Financial Functions
- `PMT(rate, nper, pv)`: Calculate payment for loan
- `FV(rate, nper, pmt, [pv], [type])`: Future value of investment
- `PV(rate, nper, pmt, [fv], [type])`: Present value of investment
- `NPV(rate, value1, value2, ...)`: Net present value
- `IRR(values, [guess])`: Internal rate of return

## Keyboard Shortcuts

### General
- **Ctrl+N**: New workbook
- **Ctrl+O**: Open workbook
- **Ctrl+S**: Save workbook
- **Ctrl+P**: Print
- **Ctrl+Z**: Undo
- **Ctrl+Y**: Redo
- **Ctrl+F**: Find
- **Ctrl+H**: Replace
- **F1**: Help

### Navigation
- **Arrow keys**: Move one cell
- **Tab**: Move one cell right
- **Shift+Tab**: Move one cell left
- **Enter**: Move one cell down
- **Shift+Enter**: Move one cell up
- **Ctrl+Arrow**: Move to edge of data region
- **Home**: Go to beginning of row
- **Ctrl+Home**: Go to beginning of sheet
- **Ctrl+End**: Go to last used cell

### Selection
- **Shift+Arrow**: Extend selection
- **Ctrl+Space**: Select entire column
- **Shift+Space**: Select entire row
- **Ctrl+A**: Select all cells
- **Ctrl+Shift+Arrow**: Extend selection to last non-empty cell

### Formatting
- **Ctrl+B**: Bold
- **Ctrl+I**: Italic
- **Ctrl+U**: Underline
- **Ctrl+1**: Format cells dialog
- **Ctrl+5**: Strikethrough

### Editing
- **F2**: Edit cell
- **Alt+Enter**: New line in cell
- **Ctrl+C**: Copy
- **Ctrl+X**: Cut
- **Ctrl+V**: Paste
- **Delete**: Clear cell content
- **Ctrl+D**: Fill down
- **Ctrl+R**: Fill right

## Credits

- Created by Aryan Daga
- Built using PyQt5
- Formula engine based on custom parser and evaluator
- Documentation and feature suggestions by GitHub Copilot

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/amazing-feature`
3. Commit your changes: `git commit -m 'Add some amazing feature'`
4. Push to the branch: `git push origin feature/amazing-feature`
5. Open a Pull Request

## Contact

Aryan Daga - [aryandaga00@gmail.com](mailto:aryandaga00@gmail.com)

Project Link: [https://github.com/aryand2006/SpreadsheetByAryan](https://github.com/aryand2006SpreadsheetByAryan)
