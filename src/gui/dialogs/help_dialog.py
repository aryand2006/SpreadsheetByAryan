from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTabWidget, 
    QTextBrowser, QScrollArea, QWidget, QTableWidget, QTableWidgetItem, 
    QHeaderView, QDialogButtonBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QPixmap, QIcon

class HelpDialog(QDialog):
    """Base class for help dialogs"""
    def __init__(self, parent=None, title="Help"):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumSize(700, 500)
        self.setup_ui()
        
    def setup_ui(self):
        """Set up the basic UI for the help dialog"""
        layout = QVBoxLayout(self)
        
        # Content area (to be implemented by subclasses)
        self.content_widget = QWidget()
        layout.addWidget(self.content_widget)
        
        # Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok)
        button_box.accepted.connect(self.accept)
        layout.addWidget(button_box)

class QuickStartDialog(HelpDialog):
    """Quick Start Guide dialog"""
    def __init__(self, parent=None):
        super().__init__(parent, title="Quick Start Guide")
        
    def setup_ui(self):
        """Set up UI with quick start information"""
        super().setup_ui()
        layout = QVBoxLayout(self.content_widget)
        
        # Create tabs for different sections
        tab_widget = QTabWidget()
        
        # Getting Started tab
        getting_started = QTextBrowser()
        getting_started.setOpenExternalLinks(True)
        getting_started.setHtml("""
        <h1>Getting Started with PySpreadsheet</h1>
        <p>Welcome to PySpreadsheet! This guide will help you get started with the basic features.</p>
        
        <h2>Creating a New Spreadsheet</h2>
        <p>To create a new spreadsheet:</p>
        <ol>
            <li>Click <b>File > New</b> or press <b>Ctrl+N</b></li>
            <li>A new empty spreadsheet with one sheet will be created</li>
        </ol>
        
        <h2>Opening and Saving</h2>
        <p>To open an existing spreadsheet:</p>
        <ol>
            <li>Click <b>File > Open</b> or press <b>Ctrl+O</b></li>
            <li>Browse to the file location and select it</li>
            <li>Click "Open"</li>
        </ol>
        
        <p>To save your work:</p>
        <ol>
            <li>Click <b>File > Save</b> or press <b>Ctrl+S</b></li>
            <li>If this is the first time saving, enter a name and location</li>
            <li>Click "Save"</li>
        </ol>
        
        <h2>Adding Data</h2>
        <p>To add data to cells:</p>
        <ol>
            <li>Click on a cell to select it</li>
            <li>Type your data</li>
            <li>Press Enter or click on another cell to confirm</li>
        </ol>
        """)
        
        # Working with Sheets tab
        sheets_tab = QTextBrowser()
        sheets_tab.setOpenExternalLinks(True)
        sheets_tab.setHtml("""
        <h1>Working with Sheets</h1>
        
        <h2>Managing Sheets</h2>
        <p>PySpreadsheet supports multiple sheets in a workbook:</p>
        <ul>
            <li>To <b>add a new sheet</b>, click the + button on the sheet tab bar or go to <b>Sheet > Insert Sheet</b></li>
            <li>To <b>rename a sheet</b>, double-click on the sheet tab or go to <b>Sheet > Rename Sheet</b></li>
            <li>To <b>delete a sheet</b>, right-click on the sheet tab and select "Delete" or go to <b>Sheet > Delete Sheet</b></li>
        </ul>
        
        <h2>Navigating Between Sheets</h2>
        <p>Click on the sheet tabs at the bottom of the window to switch between sheets.</p>
        
        <h2>Moving and Copying Sheets</h2>
        <p>To rearrange sheets, drag and drop the sheet tabs to the desired position.</p>
        """)
        
        # Basic Formatting tab
        formatting_tab = QTextBrowser()
        formatting_tab.setOpenExternalLinks(True)
        formatting_tab.setHtml("""
        <h1>Basic Formatting</h1>
        
        <h2>Cell Formatting</h2>
        <p>To format cells:</p>
        <ol>
            <li>Select the cells you want to format</li>
            <li>Use the formatting tools in the toolbar or go to the <b>Format</b> menu</li>
            <li>Apply formatting such as:
                <ul>
                    <li>Font styles (bold, italic, underline)</li>
                    <li>Font size and type</li>
                    <li>Text alignment</li>
                    <li>Cell colors and borders</li>
                </ul>
            </li>
        </ol>
        
        <h2>Number Formats</h2>
        <p>To change how numbers are displayed:</p>
        <ol>
            <li>Select the cells with numbers</li>
            <li>Go to <b>Format > Number Format</b></li>
            <li>Choose a format like:
                <ul>
                    <li>Currency</li>
                    <li>Percentage</li>
                    <li>Date</li>
                    <li>Custom formats</li>
                </ul>
            </li>
        </ol>
        """)
        
        # Add tabs to the widget
        tab_widget.addTab(getting_started, "Getting Started")
        tab_widget.addTab(sheets_tab, "Working with Sheets")
        tab_widget.addTab(formatting_tab, "Basic Formatting")
        
        layout.addWidget(tab_widget)

class FormulasHelpDialog(HelpDialog):
    """Formulas and Functions help dialog"""
    def __init__(self, parent=None):
        super().__init__(parent, title="Formulas and Functions Help")
        
    def setup_ui(self):
        """Set up UI with formulas information"""
        super().setup_ui()
        layout = QVBoxLayout(self.content_widget)
        
        # Create tabs for different categories
        tab_widget = QTabWidget()
        
        # Intro to Formulas tab
        intro_tab = QTextBrowser()
        intro_tab.setOpenExternalLinks(True)
        intro_tab.setHtml("""
        <h1>Introduction to Formulas</h1>
        
        <h2>Creating Formulas</h2>
        <p>All formulas in PySpreadsheet start with an equals sign (=). Here's how to create them:</p>
        <ol>
            <li>Select a cell where you want the formula result</li>
            <li>Type "=" followed by your formula</li>
            <li>Press Enter to calculate the result</li>
        </ol>
        
        <h2>Formula Components</h2>
        <p>Formulas can include:</p>
        <ul>
            <li>Cell references (e.g., A1, B2)</li>
            <li>Ranges (e.g., A1:B5)</li>
            <li>Functions (e.g., SUM, AVERAGE)</li>
            <li>Operators (+, -, *, /, ^)</li>
            <li>Constants (numbers, text)</li>
        </ul>
        
        <h2>Example Formulas</h2>
        <ul>
            <li><code>=A1+B1</code> - Adds the values in cells A1 and B1</li>
            <li><code>=SUM(A1:A10)</code> - Adds all values from A1 to A10</li>
            <li><code>=AVERAGE(B1:B20)*2</code> - Calculates average then multiplies by 2</li>
        </ul>
        """)
        
        # Math Functions tab
        math_tab = QTextBrowser()
        math_tab.setOpenExternalLinks(True)
        math_tab.setHtml("""
        <h1>Math Functions</h1>
        
        <table border="1" cellspacing="0" cellpadding="5" width="100%">
        <tr style="background-color: #f0f0f0;">
            <th>Function</th>
            <th>Description</th>
            <th>Example</th>
        </tr>
        <tr>
            <td><b>SUM(range)</b></td>
            <td>Adds all values in the range</td>
            <td>=SUM(A1:A10)</td>
        </tr>
        <tr>
            <td><b>AVERAGE(range)</b></td>
            <td>Calculates the average of values</td>
            <td>=AVERAGE(B1:B20)</td>
        </tr>
        <tr>
            <td><b>MIN(range)</b></td>
            <td>Finds the minimum value</td>
            <td>=MIN(C5:C15)</td>
        </tr>
        <tr>
            <td><b>MAX(range)</b></td>
            <td>Finds the maximum value</td>
            <td>=MAX(D10:D20)</td>
        </tr>
        <tr>
            <td><b>ROUND(number, decimal_places)</b></td>
            <td>Rounds a number to specified decimal places</td>
            <td>=ROUND(A1, 2)</td>
        </tr>
        <tr>
            <td><b>ABS(number)</b></td>
            <td>Returns the absolute value</td>
            <td>=ABS(B5)</td>
        </tr>
        <tr>
            <td><b>SQRT(number)</b></td>
            <td>Calculates the square root</td>
            <td>=SQRT(C3)</td>
        </tr>
        <tr>
            <td><b>POWER(number, power)</b></td>
            <td>Raises number to given power</td>
            <td>=POWER(5, 2)</td>
        </tr>
        </table>
        """)
        
        # Text Functions tab
        text_tab = QTextBrowser()
        text_tab.setOpenExternalLinks(True)
        text_tab.setHtml("""
        <h1>Text Functions</h1>
        
        <table border="1" cellspacing="0" cellpadding="5" width="100%">
        <tr style="background-color: #f0f0f0;">
            <th>Function</th>
            <th>Description</th>
            <th>Example</th>
        </tr>
        <tr>
            <td><b>CONCATENATE(text1, text2, ...)</b></td>
            <td>Joins text strings together</td>
            <td>=CONCATENATE(A1, " ", B1)</td>
        </tr>
        <tr>
            <td><b>LEFT(text, num_chars)</b></td>
            <td>Gets the leftmost characters from text</td>
            <td>=LEFT(A2, 3)</td>
        </tr>
        <tr>
            <td><b>RIGHT(text, num_chars)</b></td>
            <td>Gets the rightmost characters from text</td>
            <td>=RIGHT(B5, 4)</td>
        </tr>
        <tr>
            <td><b>MID(text, start_pos, num_chars)</b></td>
            <td>Gets characters from the middle of text</td>
            <td>=MID(C1, 2, 5)</td>
        </tr>
        <tr>
            <td><b>LEN(text)</b></td>
            <td>Returns the length of a text string</td>
            <td>=LEN(A1)</td>
        </tr>
        <tr>
            <td><b>LOWER(text)</b></td>
            <td>Converts text to lowercase</td>
            <td>=LOWER(B2)</td>
        </tr>
        <tr>
            <td><b>UPPER(text)</b></td>
            <td>Converts text to uppercase</td>
            <td>=UPPER(C3)</td>
        </tr>
        <tr>
            <td><b>PROPER(text)</b></td>
            <td>Capitalizes the first letter of each word</td>
            <td>=PROPER(D4)</td>
        </tr>
        </table>
        """)
        
        # Date & Time Functions tab
        date_tab = QTextBrowser()
        date_tab.setOpenExternalLinks(True)
        date_tab.setHtml("""
        <h1>Date and Time Functions</h1>
        
        <table border="1" cellspacing="0" cellpadding="5" width="100%">
        <tr style="background-color: #f0f0f0;">
            <th>Function</th>
            <th>Description</th>
            <th>Example</th>
        </tr>
        <tr>
            <td><b>TODAY()</b></td>
            <td>Returns the current date</td>
            <td>=TODAY()</td>
        </tr>
        <tr>
            <td><b>NOW()</b></td>
            <td>Returns the current date and time</td>
            <td>=NOW()</td>
        </tr>
        <tr>
            <td><b>DATE(year, month, day)</b></td>
            <td>Creates a date from year, month, day</td>
            <td>=DATE(2023, 12, 31)</td>
        </tr>
        <tr>
            <td><b>DAY(date)</b></td>
            <td>Returns the day component of a date</td>
            <td>=DAY(A1)</td>
        </tr>
        <tr>
            <td><b>MONTH(date)</b></td>
            <td>Returns the month component of a date</td>
            <td>=MONTH(B5)</td>
        </tr>
        <tr>
            <td><b>YEAR(date)</b></td>
            <td>Returns the year component of a date</td>
            <td>=YEAR(C10)</td>
        </tr>
        <tr>
            <td><b>WEEKDAY(date)</b></td>
            <td>Returns the day of the week (1-7)</td>
            <td>=WEEKDAY(A3)</td>
        </tr>
        <tr>
            <td><b>NETWORKDAYS(start_date, end_date)</b></td>
            <td>Returns number of workdays between two dates</td>
            <td>=NETWORKDAYS(A1, B1)</td>
        </tr>
        </table>
        """)
        
        # Logical Functions tab
        logical_tab = QTextBrowser()
        logical_tab.setOpenExternalLinks(True)
        logical_tab.setHtml("""
        <h1>Logical Functions</h1>
        
        <table border="1" cellspacing="0" cellpadding="5" width="100%">
        <tr style="background-color: #f0f0f0;">
            <th>Function</th>
            <th>Description</th>
            <th>Example</th>
        </tr>
        <tr>
            <td><b>IF(condition, value_if_true, value_if_false)</b></td>
            <td>Returns different values based on a condition</td>
            <td>=IF(A1>10, "High", "Low")</td>
        </tr>
        <tr>
            <td><b>AND(condition1, condition2, ...)</b></td>
            <td>Returns TRUE if all conditions are true</td>
            <td>=AND(A1>10, B1<20)</td>
        </tr>
        <tr>
            <td><b>OR(condition1, condition2, ...)</b></td>
            <td>Returns TRUE if any condition is true</td>
            <td>=OR(A1>10, A1<5)</td>
        </tr>
        <tr>
            <td><b>NOT(condition)</b></td>
            <td>Reverses a logical value</td>
            <td>=NOT(A1=10)</td>
        </tr>
        <tr>
            <td><b>ISBLANK(value)</b></td>
            <td>Returns TRUE if the value is blank</td>
            <td>=ISBLANK(A1)</td>
        </tr>
        <tr>
            <td><b>ISNUMBER(value)</b></td>
            <td>Returns TRUE if the value is a number</td>
            <td>=ISNUMBER(B2)</td>
        </tr>
        <tr>
            <td><b>ISTEXT(value)</b></td>
            <td>Returns TRUE if the value is text</td>
            <td>=ISTEXT(C3)</td>
        </tr>
        </table>
        """)

        # Lookup Functions tab
        lookup_tab = QTextBrowser()
        lookup_tab.setOpenExternalLinks(True)
        lookup_tab.setHtml("""
        <h1>Lookup & Reference Functions</h1>
        
        <table border="1" cellspacing="0" cellpadding="5" width="100%">
        <tr style="background-color: #f0f0f0;">
            <th>Function</th>
            <th>Description</th>
            <th>Example</th>
        </tr>
        <tr>
            <td><b>VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])</b></td>
            <td>Vertical lookup. Searches for a value in the leftmost column and returns a value in the same row from a column you specify.</td>
            <td>=VLOOKUP("John", A1:C10, 3, FALSE)</td>
        </tr>
        <tr>
            <td><b>HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])</b></td>
            <td>Horizontal lookup. Searches for a value in the top row and returns a value in the same column from a row you specify.</td>
            <td>=HLOOKUP(100, A1:J2, 2, FALSE)</td>
        </tr>
        <tr>
            <td><b>INDEX(array, row_num, [column_num])</b></td>
            <td>Returns the value at a given position in a range or array</td>
            <td>=INDEX(A1:C10, 5, 2)</td>
        </tr>
        <tr>
            <td><b>MATCH(lookup_value, lookup_array, [match_type])</b></td>
            <td>Returns the relative position of an item in an array that matches a specified value</td>
            <td>=MATCH("Apple", A1:A10, 0)</td>
        </tr>
        </table>
        """)
        
        # Add all tabs to the tab widget
        tab_widget.addTab(intro_tab, "Intro to Formulas")
        tab_widget.addTab(math_tab, "Math Functions")
        tab_widget.addTab(text_tab, "Text Functions")
        tab_widget.addTab(date_tab, "Date & Time")
        tab_widget.addTab(logical_tab, "Logical")
        tab_widget.addTab(lookup_tab, "Lookup & Reference")
        
        layout.addWidget(tab_widget)

class ShortcutsDialog(HelpDialog):
    """Keyboard Shortcuts dialog"""
    def __init__(self, parent=None):
        super().__init__(parent, title="Keyboard Shortcuts")
        
    def setup_ui(self):
        """Set up UI with keyboard shortcuts"""
        super().setup_ui()
        layout = QVBoxLayout(self.content_widget)
        
        # Table of keyboard shortcuts
        shortcuts_table = QTableWidget()
        shortcuts_table.setRowCount(25)
        shortcuts_table.setColumnCount(2)
        shortcuts_table.setHorizontalHeaderLabels(["Shortcut", "Action"])
        
        # Resize columns to content
        shortcuts_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        shortcuts_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        
        # Populate shortcuts
        self.add_shortcut(shortcuts_table, 0, "Ctrl+N", "New spreadsheet")
        self.add_shortcut(shortcuts_table, 1, "Ctrl+O", "Open spreadsheet")
        self.add_shortcut(shortcuts_table, 2, "Ctrl+S", "Save spreadsheet")
        self.add_shortcut(shortcuts_table, 3, "Ctrl+Shift+S", "Save spreadsheet as")
        self.add_shortcut(shortcuts_table, 4, "Ctrl+Z", "Undo")
        self.add_shortcut(shortcuts_table, 5, "Ctrl+Y", "Redo")
        self.add_shortcut(shortcuts_table, 6, "Ctrl+X", "Cut")
        self.add_shortcut(shortcuts_table, 7, "Ctrl+C", "Copy")
        self.add_shortcut(shortcuts_table, 8, "Ctrl+V", "Paste")
        self.add_shortcut(shortcuts_table, 9, "Ctrl+B", "Bold")
        self.add_shortcut(shortcuts_table, 10, "Ctrl+I", "Italic")
        self.add_shortcut(shortcuts_table, 11, "Ctrl+U", "Underline")
        self.add_shortcut(shortcuts_table, 12, "Ctrl+F", "Find")
        self.add_shortcut(shortcuts_table, 13, "Ctrl+H", "Replace")
        self.add_shortcut(shortcuts_table, 14, "Ctrl+Home", "Go to beginning of sheet")
        self.add_shortcut(shortcuts_table, 15, "Ctrl+End", "Go to end of sheet")
        self.add_shortcut(shortcuts_table, 16, "F2", "Edit cell")
        self.add_shortcut(shortcuts_table, 17, "Enter", "Complete cell entry")
        self.add_shortcut(shortcuts_table, 18, "Ctrl++", "Zoom in")
        self.add_shortcut(shortcuts_table, 19, "Ctrl+-", "Zoom out")
        self.add_shortcut(shortcuts_table, 20, "F7", "Spell check")
        self.add_shortcut(shortcuts_table, 21, "F11", "Create chart")
        self.add_shortcut(shortcuts_table, 22, "F12", "Save as")
        self.add_shortcut(shortcuts_table, 23, "Alt+Enter", "Start new line in the same cell")
        self.add_shortcut(shortcuts_table, 24, "Ctrl+Q", "Exit application")
        
        layout.addWidget(shortcuts_table)

    def add_shortcut(self, table, row, shortcut, description):
        """Add a shortcut to the table"""
        table.setItem(row, 0, QTableWidgetItem(shortcut))
        table.setItem(row, 1, QTableWidgetItem(description))
