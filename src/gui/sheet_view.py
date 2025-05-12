from PyQt5.QtWidgets import (
    QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, QHeaderView, 
    QAbstractItemView, QMenu, QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
    QComboBox, QLineEdit, QPushButton, QCheckBox, QGridLayout, QMessageBox,
    QSpinBox, QColorDialog, QDialogButtonBox, QTabWidget, QGroupBox, QApplication, QFontDialog
)
from PyQt5.QtCore import Qt, QEvent, QMimeData, QBuffer, QIODevice, QByteArray, QSize, QPoint, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QBrush, QPixmap, QIcon, QPainter
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
import json
import csv
import io
import re

class SheetView(QWidget):
    # Add signal to forward the table's currentCellChanged signal
    currentCellChanged = pyqtSignal(int, int, int, int)
    # Add signal for formula evaluation
    formulaEvaluationNeeded = pyqtSignal(str, int, int)  # formula, row, col
    # Add signal for notifying when data changes that might affect formulas
    dataChanged = pyqtSignal()
    
    def __init__(self, parent=None, rows=100, columns=26):
        super().__init__(parent)
        self.setWindowTitle("Spreadsheet")
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(0)
        
        # Create table widget
        self.table = QTableWidget(rows, columns)
        self.layout.addWidget(self.table)
        self.setLayout(self.layout)
        
        # Connect internal table signals to our forwarded signals
        self.table.currentCellChanged.connect(self.currentCellChanged)
        
        # Use formula evaluation handler
        self.formulaEvaluationNeeded.connect(self.evaluateFormula)
        
        # Configure table
        self.configure_table()
        self.initialize_table()
        
        # Track changes for undo/redo
        self.history = []
        self.redo_stack = []
        
        # Cell display values (for formulas)
        self.cell_display_values = {}
        
        # Conditional formatting rules
        self.conditional_formatting_rules = []
        
        # Remember current selection for copy/paste
        self.copied_cells = []
        self.cut_mode = False
        
        # Zoom level (100% = 1.0)
        self.zoom_level = 1.0
        
        # Track if filtering is active
        self.filtering_active = False
        self.hidden_rows = set()
        
        # Find/replace context
        self.find_text = ""
        self.last_found_cell = (-1, -1)
        
        # Number formatting
        self.number_formats = {}

        # Add event filter to handle Enter key
        self.table.installEventFilter(self)

    def configure_table(self):
        # Set column headers (A, B, C...)
        self.update_column_headers()
        
        # Set row headers (1, 2, 3...)
        self.update_row_headers()
        
        # Configure header behavior
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.Interactive)
        
        # Set selection behavior
        self.table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        
        # Enable context menu
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        
        # Connect item changed signal
        self.table.itemChanged.connect(self.on_item_changed)
        
        # Make the grid lines more visible
        self.table.setShowGrid(True)
        self.table.setGridStyle(Qt.SolidLine)
        
        # Enable sorting
        self.table.setSortingEnabled(False)  # We'll handle sorting manually
        
        # Set default row heights and column widths
        for col in range(self.table.columnCount()):
            self.table.setColumnWidth(col, 100)
        
        for row in range(self.table.rowCount()):
            self.table.setRowHeight(row, 22)
            
        # Add event filter to handle Enter key
        self.table.installEventFilter(self)

    def update_column_headers(self):
        """Update column headers based on the current number of columns"""
        column_headers = []
        for i in range(self.table.columnCount()):
            if i < 26:
                # A-Z
                column_headers.append(chr(65 + i))
            else:
                # AA, AB, etc.
                first_char = chr(64 + int(i / 26))
                second_char = chr(65 + (i % 26))
                column_headers.append(first_char + second_char)
                
        self.table.setHorizontalHeaderLabels(column_headers)

    def update_row_headers(self):
        """Update row headers based on the current number of rows"""
        row_headers = [str(i+1) for i in range(self.table.rowCount())]
        self.table.setVerticalHeaderLabels(row_headers)

    def initialize_table(self):
        for row in range(self.table.rowCount()):
            for column in range(self.table.columnCount()):
                item = QTableWidgetItem("")
                self.table.setItem(row, column, item)

    def get_cell_value(self, row, column):
        """Get the raw value (formula or text) from a cell"""
        item = self.table.item(row, column)
        if not item:
            return None
        
        # Check if this is a formula cell (stored in UserRole)
        formula = item.data(Qt.UserRole)
        if formula and isinstance(formula, str) and formula.startswith('='):
            # Return the display value instead of the formula
            cell_key = f"{row},{column}"
            if cell_key in self.cell_display_values:
                # Try to convert to number if possible
                display_value = self.cell_display_values[cell_key]
                try:
                    return float(display_value)
                except (ValueError, TypeError):
                    return display_value
            return formula
        
        # Otherwise, return the normal text
        return item.text()

    def get_cell_display_value(self, row, column):
        """Get the display value (evaluated result) for a cell"""
        cell_key = f"{row},{column}"
        if cell_key in self.cell_display_values:
            return self.cell_display_values[cell_key]
        else:
            return self.get_cell_value(row, column)

    def set_cell_value(self, row, column, value):
        """Set the raw value (formula or text) of a cell"""
        # Store previous value for undo
        prev_value = self.get_cell_value(row, column)
        if prev_value != value:
            self.history.append({
                'type': 'cell_edit',
                'row': row,
                'column': column,
                'old_value': prev_value,
                'new_value': value
            })
            
            # Clear redo stack when a new change is made
            self.redo_stack = []
        
        # Set the new value
        item = self.table.item(row, column)
        if not item:
            item = QTableWidgetItem(value)
            self.table.setItem(row, column, item)
        else:
            # If it's a formula, handle it specially
            if value and isinstance(value, str) and value.startswith('='):
                item.setText(value)  # This will trigger on_item_changed
            else:
                # Not a formula, just set text normally
                item.setData(Qt.UserRole, None)  # Clear any formula stored
                item.setText(value)
            
        # Apply any conditional formatting
        self.apply_conditional_formatting_to_cell(row, column)

    def set_cell_display_value(self, row, column, display_value):
        """Set the display value for a cell with a formula"""
        cell_key = f"{row},{column}"
        self.cell_display_values[cell_key] = display_value
        
        # Update the cell's display if it's a formula result
        item = self.table.item(row, column)
        if item and item.text().startswith('='):
            # Store formula as user data and show result as display text
            formula = item.text()
            
            # The actual result is shown in the tooltip
            item.setToolTip(f"Formula: {formula}\nResult: {display_value}")
            
            # Also update the display text
            self.table.blockSignals(True)
            if isinstance(display_value, (int, float, str)):
                item.setData(Qt.DisplayRole, str(display_value))
                # Store the original formula
                item.setData(Qt.UserRole, formula)
            self.table.blockSignals(False)

    def clear_cell(self, row, column):
        """Clear the content of a cell"""
        self.set_cell_value(row, column, "")
        cell_key = f"{row},{column}"
        if cell_key in self.cell_display_values:
            del self.cell_display_values[cell_key]

    def clear_selected_cells(self):
        """Clear content from all selected cells"""
        for item in self.table.selectedItems():
            row = item.row()
            column = item.column()
            self.clear_cell(row, column)

    def undo(self):
        """Undo the last action"""
        if not self.history:
            return
            
        action = self.history.pop()
        self.redo_stack.append(action)
        
        if action['type'] == 'cell_edit':
            row = action['row']
            column = action['column']
            old_value = action['old_value']
            
            # Temporarily disconnect the itemChanged signal
            self.table.itemChanged.disconnect(self.on_item_changed)
            
            # Restore old value
            item = self.table.item(row, column)
            if item:
                item.setText(old_value if old_value else "")
                
            # Reconnect the signal
            self.table.itemChanged.connect(self.on_item_changed)

    def redo(self):
        """Redo the last undone action"""
        if not self.redo_stack:
            return
            
        action = self.redo_stack.pop()
        self.history.append(action)
        
        if action['type'] == 'cell_edit':
            row = action['row']
            column = action['column']
            new_value = action['new_value']
            
            # Temporarily disconnect the itemChanged signal
            self.table.itemChanged.disconnect(self.on_item_changed)
            
            # Apply new value
            item = self.table.item(row, column)
            if item:
                item.setText(new_value if new_value else "")
                
            # Reconnect the signal
            self.table.itemChanged.connect(self.on_item_changed)

    def load_data(self, data):
        """Load data from a list of lists into the sheet"""
        # Clear existing data and undo history
        self.table.clearContents()
        self.history = []
        self.redo_stack = []
        self.cell_display_values = {}
        if hasattr(self, 'cell_formulas'):
            self.cell_formulas = {}
        
        # Ensure table is large enough
        max_rows = max(len(data), self.table.rowCount())
        max_cols = max(len(data[0]) if data else 0, self.table.columnCount())
        
        if max_rows > self.table.rowCount():
            self.table.setRowCount(max_rows)
            self.update_row_headers()
        
        if max_cols > self.table.columnCount():
            self.table.setColumnCount(max_cols)
            self.update_column_headers()
        
        # Temporarily disconnect the itemChanged signal
        self.table.itemChanged.disconnect(self.on_item_changed)
        
        # Load data into cells
        for row_idx, row_data in enumerate(data):
            for col_idx, cell_value in enumerate(row_data):
                if cell_value is not None:  # Skip None values
                    item = QTableWidgetItem(str(cell_value))
                    self.table.setItem(row_idx, col_idx, item)
                    
                    # If it's a formula, store it properly
                    if str(cell_value).startswith('='):
                        cell_key = f"{row_idx},{col_idx}"
                        self.cell_display_values[cell_key] = "Formula"  # Will be evaluated later
        
        # Reconnect the signal
        self.table.itemChanged.connect(self.on_item_changed)
        
        # Apply conditional formatting to all cells
        self.apply_conditional_formatting_to_all_cells()

    def get_all_data(self):
        """Extract all data from the sheet as a list of lists"""
        rows = self.table.rowCount()
        cols = self.table.columnCount()
        data = []
        
        for row in range(rows):
            if row in self.hidden_rows and self.filtering_active:
                continue
                
            row_data = []
            for col in range(cols):
                item = self.table.item(row, col)
                cell_value = item.text() if item else ""
                row_data.append(cell_value)
            data.append(row_data)
        
        return data

    def apply_font_to_selected_cells(self, font):
        """Apply a font to selected cells"""
        for item in self.table.selectedItems():
            item.setFont(font)

    def apply_background_color_to_selected_cells(self, color):
        """Apply background color to selected cells"""
        for item in self.table.selectedItems():
            item.setBackground(QBrush(color))

    def apply_text_color_to_selected_cells(self, color):
        """Apply text color to selected cells"""
        for item in self.table.selectedItems():
            item.setForeground(QBrush(color))

    def clear_all_cells(self):
        """Clear content from all cells"""
        self.table.clearContents()
        self.initialize_table()
        self.history = []
        self.redo_stack = []
        self.cell_display_values = {}
        self.conditional_formatting_rules = []
        self.hidden_rows = set()
        self.filtering_active = False

    def zoom_in(self):
        """Increase the zoom level of the sheet"""
        if self.zoom_level < 3.0:  # Limit maximum zoom
            self.zoom_level += 0.1
            self.apply_zoom()

    def zoom_out(self):
        """Decrease the zoom level of the sheet"""
        if self.zoom_level > 0.5:  # Limit minimum zoom
            self.zoom_level -= 0.1
            self.apply_zoom()

    def apply_zoom(self):
        """Apply the current zoom level to the table"""
        base_font_size = 10
        base_row_height = 22
        base_col_width = 100
        
        # Adjust font size based on zoom
        font = self.table.font()
        new_size = int(base_font_size * self.zoom_level)
        font.setPointSize(max(1, new_size))  # Ensure font size is at least 1
        self.table.setFont(font)
        
        # Adjust row heights and column widths
        for row in range(self.table.rowCount()):
            self.table.setRowHeight(row, int(base_row_height * self.zoom_level))
            
        for col in range(self.table.columnCount()):
            self.table.setColumnWidth(col, int(base_col_width * self.zoom_level))
            
    def export_to_pdf(self, file_path):
        """Export the current sheet to a PDF file"""
        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(file_path)
        
        # Create a painter to paint on the printer
        painter = QPainter()
        if not painter.begin(printer):
            return False  # Failed to open printer
        
        # Paint the table
        self.table.render(painter)
        
        # End painting
        painter.end()
        return True

    def show_context_menu(self, position):
        """Show the context menu with cell operations"""
        menu = QMenu()
        
        # Add actions
        cut_action = menu.addAction(QIcon(), "Cut")
        copy_action = menu.addAction(QIcon(), "Copy")
        paste_action = menu.addAction(QIcon(), "Paste")
        menu.addSeparator()
        clear_action = menu.addAction(QIcon(), "Clear")
        menu.addSeparator()
        
        format_menu = menu.addMenu("Format")
        font_action = format_menu.addAction("Font...")
        cell_color_action = format_menu.addAction("Cell Color...")
        text_color_action = format_menu.addAction("Text Color...")
        conditional_format_action = format_menu.addAction("Conditional Formatting...")
        number_format_menu = format_menu.addMenu("Number Format")
        currency_action = number_format_menu.addAction("Currency")
        percentage_action = number_format_menu.addAction("Percentage")
        comma_action = number_format_menu.addAction("Comma")
        decimal_2_action = number_format_menu.addAction("2 Decimal Places")
        decimal_4_action = number_format_menu.addAction("4 Decimal Places")
        scientific_action = number_format_menu.addAction("Scientific")
        
        menu.addSeparator()
        insert_row_action = menu.addAction(QIcon(), "Insert Row")
        insert_col_action = menu.addAction(QIcon(), "Insert Column")
        delete_row_action = menu.addAction(QIcon(), "Delete Row")
        delete_col_action = menu.addAction(QIcon(), "Delete Column")
        
        menu.addSeparator()
        sort_menu = menu.addMenu("Sort")
        sort_asc_action = sort_menu.addAction("Sort Ascending")
        sort_desc_action = sort_menu.addAction("Sort Descending")
        
        filter_menu = menu.addMenu("Filter")
        filter_action = filter_menu.addAction("Filter Data...")
        remove_filter_action = filter_menu.addAction("Remove Filter")
        
        # Connect actions
        cut_action.triggered.connect(self.cut_cells)
        copy_action.triggered.connect(self.copy_cells)
        paste_action.triggered.connect(self.paste_cells)
        clear_action.triggered.connect(self.clear_selected_cells)
        
        font_action.triggered.connect(self.show_font_dialog)
        cell_color_action.triggered.connect(self.show_cell_color_dialog)
        text_color_action.triggered.connect(self.show_text_color_dialog)
        conditional_format_action.triggered.connect(self.show_conditional_format_dialog)
        currency_action.triggered.connect(lambda: self.apply_number_format("currency"))
        percentage_action.triggered.connect(lambda: self.apply_number_format("percentage"))
        comma_action.triggered.connect(lambda: self.apply_number_format("comma"))
        decimal_2_action.triggered.connect(lambda: self.apply_number_format("decimal_2"))
        decimal_4_action.triggered.connect(lambda: self.apply_number_format("decimal_4"))
        scientific_action.triggered.connect(lambda: self.apply_number_format("scientific"))
        
        insert_row_action.triggered.connect(self.insert_row)
        insert_col_action.triggered.connect(self.insert_column)
        delete_row_action.triggered.connect(self.delete_row)
        delete_col_action.triggered.connect(self.delete_column)
        
        sort_asc_action.triggered.connect(lambda: self.sort_selected_data(True))
        sort_desc_action.triggered.connect(lambda: self.sort_selected_data(False))
        filter_action.triggered.connect(self.show_filter_dialog)
        remove_filter_action.triggered.connect(self.remove_filter)
        
        # Show the menu
        menu.exec_(self.table.viewport().mapToGlobal(position))

    def on_item_changed(self, item):
        """Handle changes to cell items"""
        row = item.row()
        column = item.column()
        value = item.text()
        
        # Check if it's a formula
        if value and value.startswith('='):
            # Emit signal for formula evaluation
            self.formulaEvaluationNeeded.emit(value, row, column)
            
            # Let the parent window know data has changed (affects other formulas)
            self.dataChanged.emit()
        else:
            # Non-formula data changed, may affect formulas
            self.dataChanged.emit()

    def evaluateFormula(self, formula, row, column):
        """Evaluate a formula using the calculator from parent window"""
        parent = self.window()
        result = None
        try:
            if hasattr(parent, 'calculator'):
                # The calculator's evaluate method expects only 3 parameters: formula, row, column
                # We need to set up a way for the calculator to access cell values
                if not hasattr(parent.calculator, 'sheet_view'):
                    parent.calculator.sheet_view = self
                result = parent.calculator.evaluate(formula, row, column)
            else:
                result = "No calculator"
        except Exception as e:
            result = f"Error: {str(e)}"
            
        # Update the cell with the result
        self.set_cell_display_value(row, column, result)
        
        # Set tooltip and formula storage
        item = self.table.item(row, column)
        if item:
            # Set tooltip to show the formula and result
            item.setToolTip(f"Formula: {formula}\nResult: {result}")
            
            # Display the result value instead of the formula
            # Temporarily disconnect signals to avoid recursion
            self.table.blockSignals(True)
            item.setData(Qt.DisplayRole, str(result) if result is not None else "")
            # Store the original formula for future reference
            item.setData(Qt.UserRole, formula)
            self.table.blockSignals(False)
            
    def getCellValueForFormula(self, cell_ref):
        """Get a cell value by reference (like A1, B2) for formula evaluation"""
        # Convert cell reference (e.g. 'A1') to row/column indices
        if not cell_ref or len(cell_ref) < 2:
            return 0
            
        # Parse column letter (A, B, C..., AA, AB...)
        col_part = ""
        row_part = ""
        
        for char in cell_ref:
            if char.isalpha():
                col_part += char
            elif char.isdigit():
                row_part += char
                
        if not col_part or not row_part:
            return 0
            
        # Convert column letters to column index (A=0, B=1, ..., Z=25, AA=26, ...)
        col = 0
        for i, char in enumerate(reversed(col_part.upper())):
            col += (ord(char) - ord('A') + 1) * (26 ** i)
        col -= 1  # Adjust to 0-based indexing
        
        # Convert row number to row index (1-based to 0-based)
        try:
            row = int(row_part) - 1
        except ValueError:
            return 0
            
        # Get cell value
        value = self.get_cell_value(row, col)
        
        # Convert to number if possible
        try:
            return float(value) if value else 0
        except (ValueError, TypeError):
            # Return as string if not numeric
            return value if value else 0
        
    def show_font_dialog(self):
        """Show the font dialog"""
        font, ok = QFontDialog.getFont()
        if ok:
            self.apply_font_to_selected_cells(font)

    def show_cell_color_dialog(self):
        """Show the cell color dialog"""
        color = QColorDialog.getColor()
        if color.isValid():
            self.apply_background_color_to_selected_cells(color)

    def show_text_color_dialog(self):
        """Show the text color dialog"""
        color = QColorDialog.getColor()
        if color.isValid():
            self.apply_text_color_to_selected_cells(color)

    def show_conditional_format_dialog(self):
        """Show the conditional formatting dialog"""
        dialog = ConditionalFormatDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            self.apply_conditional_formatting_to_all_cells()

    def apply_conditional_formatting(self):
        """Apply conditional formatting options from the dialog"""
        self.show_conditional_format_dialog()
    
    def apply_conditional_formatting_to_cell(self, row, column):
        """Apply all conditional formatting rules to a specific cell"""
        if not self.conditional_formatting_rules:
            return
            
        cell_value = self.get_cell_value(row, column)
        if not cell_value:
            return
            
        item = self.table.item(row, column)
        if not item:
            return
            
        # Check each rule
        for rule in self.conditional_formatting_rules:
            condition = rule['condition']
            value = rule['value']
            bg_color = QColor(rule['bg_color']) if rule['bg_color'] else None
            text_color = QColor(rule['text_color']) if rule['text_color'] else None
            
            # Try to convert cell value to a number for numeric comparisons
            try:
                numeric_cell_value = float(cell_value)
                numeric_comparison = True
            except ValueError:
                numeric_cell_value = cell_value
                numeric_comparison = False
            
            # Apply the condition
            if condition == "equal_to" and cell_value == value:
                self.apply_formatting(item, bg_color, text_color)
                
            elif condition == "not_equal_to" and cell_value != value:
                self.apply_formatting(item, bg_color, text_color)
                
            elif condition == "greater_than" and numeric_comparison:
                try:
                    if numeric_cell_value > float(value):
                        self.apply_formatting(item, bg_color, text_color)
                except ValueError:
                    pass
                    
            elif condition == "less_than" and numeric_comparison:
                try:
                    if numeric_cell_value < float(value):
                        self.apply_formatting(item, bg_color, text_color)
                except ValueError:
                    pass
                    
            elif condition == "contains" and isinstance(cell_value, str) and value in cell_value:
                self.apply_formatting(item, bg_color, text_color)
                
            elif condition == "not_contains" and isinstance(cell_value, str) and value not in cell_value:
                self.apply_formatting(item, bg_color, text_color)

    def apply_formatting(self, item, bg_color=None, text_color=None):
        """Apply formatting to a table item"""
        if bg_color:
            item.setBackground(QBrush(bg_color))
            
        if text_color:
            item.setForeground(QBrush(text_color))

    def apply_conditional_formatting_to_all_cells(self):
        """Apply all conditional formatting rules to all cells"""
        if not self.conditional_formatting_rules:
            return
            
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                self.apply_conditional_formatting_to_cell(row, col)

    def cut_cells(self):
        """Cut selected cells"""
        self.copy_cells()
        self.cut_mode = True

    def copy_cells(self):
        """Copy selected cells to clipboard"""
        self.copied_cells = []
        selected_items = self.table.selectedItems()
        
        if not selected_items:
            return
            
        # Find the top-left cell in the selection
        min_row = min(item.row() for item in selected_items)
        min_col = min(item.column() for item in selected_items)
        
        # Store the copied data with relative positions
        for item in selected_items:
            rel_row = item.row() - min_row
            rel_col = item.column() - min_col
            value = item.text()
            font = item.font()
            bg_color = item.background().color()
            fg_color = item.foreground().color()
            
            self.copied_cells.append({
                'rel_row': rel_row,
                'rel_col': rel_col,
                'value': value,
                'font_family': font.family(),
                'font_size': font.pointSize(),
                'font_bold': font.bold(),
                'font_italic': font.italic(),
                'font_underline': font.underline(),
                'bg_color': bg_color.name(),
                'fg_color': fg_color.name()
            })
            
        # Also copy to system clipboard as CSV
        self.copy_to_system_clipboard(selected_items, min_row, min_col)
        self.cut_mode = False

    def copy_to_system_clipboard(self, selected_items, min_row, min_col):
        """Copy data to system clipboard in CSV format"""
        # Find the bounds of the selection
        max_row = max(item.row() for item in selected_items)
        max_col = max(item.column() for item in selected_items)
        
        # Create a 2D list to hold the cell contents
        data = [['' for _ in range(max_col - min_col + 1)] for _ in range(max_row - min_row + 1)]
        
        # Fill in the data
        for item in selected_items:
            rel_row = item.row() - min_row
            rel_col = item.column() - min_col
            data[rel_row][rel_col] = item.text()
        
        # Create CSV string
        output = io.StringIO()  # Updated to use renamed import
        csv_writer = csv.writer(output, quoting=csv.QUOTE_MINIMAL)
        csv_writer.writerows(data)
        csv_text = output.getvalue()
        
        # Create HTML representation for rich clipboard
        html = "<table>"
        for row in data:
            html += "<tr>"
            for cell in row:
                html += f"<td>{cell}</td>"
            html += "</tr>"
        html += "</table>"
        
        # Create a mime data object
        mime_data = QMimeData()
        mime_data.setText(csv_text)
        mime_data.setHtml(html)
        
        # Set the mime data to the clipboard
        clipboard = QApplication.clipboard()
        clipboard.setMimeData(mime_data)

    def paste_cells(self):
        """Paste copied cells at the current selection"""
        if not self.copied_cells:
            # Try to paste from system clipboard if we don't have internal copy data
            self.paste_from_system_clipboard()
            return
            
        # Get the target cell (current cell)
        current_row = self.table.currentRow()
        current_col = self.table.currentColumn()
        
        if current_row < 0 or current_col < 0:
            return
            
        # Temporarily disconnect item changed signal
        self.table.itemChanged.disconnect(self.on_item_changed)
        
        # Paste all copied cells
        for cell_data in self.copied_cells:
            target_row = current_row + cell_data['rel_row']
            target_col = current_col + cell_data['rel_col']
            
            # Ensure the target cell is within the table
            if target_row >= 0 and target_row < self.table.rowCount() and \
               target_col >= 0 and target_col < self.table.columnCount():
                
                # Get or create the item
                item = self.table.item(target_row, target_col)
                if not item:
                    item = QTableWidgetItem("")
                    self.table.setItem(target_row, target_col, item)
                
                # Record for undo
                old_value = item.text()
                self.history.append({
                    'type': 'cell_edit',
                    'row': target_row,
                    'column': target_col,
                    'old_value': old_value,
                    'new_value': cell_data['value']
                })
                
                # Apply value and formatting
                item.setText(cell_data['value'])
                
                font = QFont(cell_data['font_family'], cell_data['font_size'])
                font.setBold(cell_data['font_bold'])
                font.setItalic(cell_data['font_italic'])
                font.setUnderline(cell_data['font_underline'])
                item.setFont(font)
                
                item.setBackground(QBrush(QColor(cell_data['bg_color'])))
                item.setForeground(QBrush(QColor(cell_data['fg_color'])))
                
                # Clear the source cell if in cut mode
                if self.cut_mode:
                    source_row = target_row - current_row + cell_data['rel_row']
                    source_col = target_col - current_col + cell_data['rel_col']
                    source_item = self.table.item(source_row, source_col)
                    if source_item:
                        source_item.setText("")
        
        # Reconnect signal
        self.table.itemChanged.connect(self.on_item_changed)
        
        # Clear cut mode after paste
        self.cut_mode = False

    def paste_from_system_clipboard(self):
        """Paste data from system clipboard"""
        clipboard = QApplication.clipboard()
        mime_data = clipboard.mimeData()
        
        if mime_data.hasText():
            text = mime_data.text()
            current_row = self.table.currentRow()
            current_col = self.table.currentColumn()
            
            # Check for CSV format (contains commas or tabs)
            if ',' in text or '\t' in text:
                # Parse as CSV
                delimiter = '\t' if '\t' in text else ','
                reader = csv.reader(text.split('\n'), delimiter=delimiter)
                rows = list(reader)
                
                # Temporarily disconnect item changed signal
                self.table.itemChanged.disconnect(self.on_item_changed)
                
                # Paste the CSV data
                for i, row_data in enumerate(rows):
                    for j, cell_value in enumerate(row_data):
                        target_row = current_row + i
                        target_col = current_col + j
                        
                        if target_row < self.table.rowCount() and target_col < self.table.columnCount():
                            # Get or create the item
                            item = self.table.item(target_row, target_col)
                            if not item:
                                item = QTableWidgetItem("")
                                self.table.setItem(target_row, target_col, item)
                                
                            # Record for undo
                            old_value = item.text()
                            self.history.append({
                                'type': 'cell_edit',
                                'row': target_row,
                                'column': target_col,
                                'old_value': old_value,
                                'new_value': cell_value
                            })
                            
                            # Set cell value
                            item.setText(cell_value)
                
                # Reconnect signal
                self.table.itemChanged.connect(self.on_item_changed)
                
            else:
                # Just paste as a single cell
                item = self.table.item(current_row, current_col)
                if item:
                    old_value = item.text()
                    self.history.append({
                        'type': 'cell_edit',
                        'row': current_row,
                        'column': current_col,
                        'old_value': old_value,
                        'new_value': text
                    })
                    item.setText(text)
                else:
                    item = QTableWidgetItem(text)
                    self.table.setItem(current_row, current_col, item)
                    self.history.append({
                        'type': 'cell_edit',
                        'row': current_row,
                        'column': current_col,
                        'old_value': "",
                        'new_value': text
                    })

    def insert_row(self):
        """Insert a new row above the current row"""
        current_row = self.table.currentRow()
        if current_row >= 0:
            self.table.insertRow(current_row)
            
            # Initialize the cells in the new row
            for col in range(self.table.columnCount()):
                item = QTableWidgetItem("")
                self.table.setItem(current_row, col, item)
                
            # Update row headers
            self.update_row_headers()

    def delete_row(self):
        """Delete the current row"""
        current_row = self.table.currentRow()
        if current_row >= 0:
            self.table.removeRow(current_row)
            
            # Update row headers
            self.update_row_headers()

    def insert_column(self):
        """Insert a new column to the left of the current column"""
        current_column = self.table.currentColumn()
        if current_column >= 0:
            self.table.insertColumn(current_column)
            
            # Initialize the cells in the new column
            for row in range(self.table.rowCount()):
                item = QTableWidgetItem("")
                self.table.setItem(row, current_column, item)
                
            # Update column headers
            self.update_column_headers()

    def delete_column(self):
        """Delete the current column"""
        current_column = self.table.currentColumn()
        if current_column >= 0:
            self.table.removeColumn(current_column)
            
            # Update column headers
            self.update_column_headers()

    def sort_selected_data(self, ascending=True):
        """Sort the selected data"""
        selected_ranges = self.table.selectedRanges()
        if not selected_ranges:
            return
            
        range_ = selected_ranges[0]  # Use the first selected range
        
        # Get the column to sort by (use the first column in the selection)
        sort_col = range_.leftColumn()
        
        # Extract the data and associated rows
        data_with_indices = []
        for row in range(range_.topRow(), range_.bottomRow() + 1):
            row_data = []
            for col in range(range_.leftColumn(), range_.rightColumn() + 1):
                item = self.table.item(row, col)
                value = item.text() if item else ""
                row_data.append(value)
            data_with_indices.append((row, row_data))
            
        # Sort by the first column in the selection
        data_with_indices.sort(key=lambda x: self._sort_key(x[1][0]), reverse=not ascending)
        
        # Temporarily disconnect item changed signal
        self.table.itemChanged.disconnect(self.on_item_changed)
        
        # Update the table with sorted data
        for i, (original_row, row_data) in enumerate(data_with_indices):
            target_row = range_.topRow() + i
            for j, value in enumerate(row_data):
                col = range_.leftColumn() + j
                item = self.table.item(target_row, col)
                if item:
                    item.setText(value)
        
        # Reconnect signal
        self.table.itemChanged.connect(self.on_item_changed)

    def _sort_key(self, value):
        """Create a sort key that properly handles numbers and text"""
        try:
            return (0, float(value))  # Numbers first
        except ValueError:
            return (1, value)  # Then text

    def show_filter_dialog(self):
        """Show the filter dialog"""
        dialog = FilterDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            self.apply_filter(dialog.get_filter_options())

    def apply_filter(self, options):
        """Apply filter to the data based on the options"""
        col_index = options['column']
        condition = options['condition']
        filter_value = options['value']
        
        self.hidden_rows = set()
        self.filtering_active = True
        
        for row in range(self.table.rowCount()):
            item = self.table.item(row, col_index)
            cell_value = item.text() if item else ""
            
            show_row = True  # Default to showing the row
            
            # Apply filter condition
            if condition == "equal_to":
                show_row = (cell_value == filter_value)
            elif condition == "not_equal_to":
                show_row = (cell_value != filter_value)
            elif condition == "contains":
                show_row = (filter_value in cell_value)
            elif condition == "not_contains":
                show_row = (filter_value not in cell_value)
            elif condition == "greater_than":
                try:
                    show_row = (float(cell_value) > float(filter_value))
                except ValueError:
                    show_row = False
            elif condition == "less_than":
                try:
                    show_row = (float(cell_value) < float(filter_value))
                except ValueError:
                    show_row = False
            
            # Hide or show the row
            if not show_row:
                self.table.hideRow(row)
                self.hidden_rows.add(row)
            else:
                self.table.showRow(row)

    def remove_filter(self):
        """Remove all filters and show all rows"""
        if self.filtering_active:
            for row in self.hidden_rows:
                self.table.showRow(row)
                
            self.filtering_active = False
            self.hidden_rows = set()

    def show_find_dialog(self):
        """Show the find dialog"""
        dialog = FindDialog(self)
        dialog.exec_()

    def show_replace_dialog(self):
        """Show the replace dialog"""
        dialog = ReplaceDialog(self)
        dialog.exec_()

    def find_next(self, text, match_case=False, match_whole_word=False, search_in_formulas=True):
        """Find the next occurrence of the text"""
        if not text:
            return False
            
        # Save search text for next find
        self.find_text = text
        
        rows = self.table.rowCount()
        cols = self.table.columnCount()
        
        # Start from the cell after the current one
        start_row, start_col = self.last_found_cell
        if start_row == -1:  # First search
            start_row = 0
            start_col = 0
        else:
            # Move to next cell
            start_col += 1
            if start_col >= cols:
                start_col = 0
                start_row += 1
        
        # Search through cells
        for row in range(start_row, rows):
            for col in range(start_col if row == start_row else 0, cols):
                item = self.table.item(row, col)
                cell_text = item.text() if item else ""
                
                if not search_in_formulas and cell_text.startswith('='):
                    continue
                    
                # Apply search options
                if not match_case:
                    cell_text = cell_text.lower()
                    search_text = text.lower()
                else:
                    search_text = text
                    
                if match_whole_word:
                    # Simple whole word search
                    words = re.findall(r'\b\w+\b', cell_text)
                    found = search_text in words
                else:
                    found = search_text in cell_text
                    
                if found:
                    # Select and scroll to the cell
                    self.table.setCurrentCell(row, col)
                    self.last_found_cell = (row, col)
                    return True
        
        # If we've searched the entire sheet without finding anything
        self.last_found_cell = (-1, -1)  # Reset for next search
        return False
        
    def replace_current(self, find_text, replace_text, match_case=False, match_whole_word=False, search_in_formulas=True):
        """Replace the current occurrence of the find_text with replace_text"""
        row, col = self.last_found_cell
        if row == -1:  # No current match
            return False
            
        item = self.table.item(row, col)
        if not item:
            return False
            
        cell_text = item.text()
        
        # Apply replacement
        if not match_case:
            new_text = re.sub(re.escape(find_text), replace_text, cell_text, flags=re.IGNORECASE)
        else:
            new_text = cell_text.replace(find_text, replace_text)
            
        # Update the cell
        self.set_cell_value(row, col, new_text)
        return True
        
    def replace_all(self, find_text, replace_text, match_case=False, match_whole_word=False, search_in_formulas=True):
        """Replace all occurrences of the find_text with replace_text"""
        rows = self.table.rowCount()
        cols = self.table.columnCount()
        count = 0
        
        # Reset search position
        self.last_found_cell = (-1, -1)
        
        # Temporarily disconnect item changed signal
        self.table.itemChanged.disconnect(self.on_item_changed)
        
        for row in range(rows):
            for col in range(cols):
                item = self.table.item(row, col)
                if not item:
                    continue
                    
                cell_text = item.text()
                
                if not search_in_formulas and cell_text.startswith('='):
                    continue
                    
                # Apply search options
                if not match_case:
                    search_pattern = re.compile(re.escape(find_text), re.IGNORECASE)
                else:
                    search_pattern = re.compile(re.escape(find_text))
                    
                if match_whole_word:
                    # For whole word, use word boundaries in regex
                    search_pattern = re.compile(r'\b' + re.escape(find_text) + r'\b', 
                                               re.IGNORECASE if not match_case else 0)
                
                # Check if the text is found
                if re.search(search_pattern, cell_text):
                    # Record for undo
                    old_value = cell_text
                    new_value = re.sub(search_pattern, replace_text)
                    
                    self.history.append({
                        'type': 'cell_edit',
                        'row': row,
                        'column': col,
                        'old_value': old_value,
                        'new_value': new_value
                    })
                    
                    # Update cell text
                    item.setText(new_value)
                    count += 1
        
        # Reconnect signal
        self.table.itemChanged.connect(self.on_item_changed)
        
        return count

    def apply_number_format(self, format_type):
        """Apply number format to selected cells"""
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        
        for item in selected_items:
            # Store original value for undo
            row = item.row()
            col = item.column()
            text = item.text()
            
            # Try to format the number
            try:
                value = float(text)
                formatted_text = text  # Default to original text
                
                if format_type == "currency":
                    formatted_text = "${:,.2f}".format(value)
                elif format_type == "percentage":
                    formatted_text = "{:.2f}%".format(value * 100)
                elif format_type == "comma":
                    formatted_text = "{:,.0f}".format(value)
                elif format_type == "decimal_2":
                    formatted_text = "{:.2f}".format(value)
                elif format_type == "decimal_4":
                    formatted_text = "{:.4f}".format(value)
                elif format_type == "scientific":
                    formatted_text = "{:.2e}".format(value)
                
                # Store the format for future use
                format_key = f"{row},{col}"
                self.number_formats[format_key] = format_type
                
                # Update the cell display but keep the original value
                item.setData(Qt.DisplayRole, formatted_text)
                item.setData(Qt.UserRole, text)  # Store original value
                
            except ValueError:
                # Not a number, ignore formatting
                pass

    def merge_cells(self):
        """Merge selected cells"""
        ranges = self.table.selectedRanges()
        if not ranges:
            return
        
        # We'll use the first selected range
        range_ = ranges[0]
        top_row = range_.topRow()
        left_col = range_.leftColumn()
        bottom_row = range_.bottomRow()
        right_col = range_.rightColumn()
        
        # Get content from the top-left cell
        top_left = self.table.item(top_row, left_col)
        content = top_left.text() if top_left else ""
        
        # Store merge data for undo
        self.history.append({
            'type': 'merge_cells',
            'range': (top_row, left_col, bottom_row, right_col),
            'previous_state': self.get_range_data(top_row, left_col, bottom_row, right_col)
        })
        
        # Clear content in the range
        for row in range(top_row, bottom_row + 1):
            for col in range(left_col, right_col + 1):
                if row == top_row and col == left_col:
                    continue  # Skip the top-left cell
                item = self.table.item(row, col)
                if item:
                    item.setText("")
        
        # Span the cells
        self.table.setSpan(top_row, left_col, 
                          bottom_row - top_row + 1, 
                          right_col - left_col + 1)

    def split_cells(self):
        """Split previously merged cells"""
        row = self.table.currentRow()
        col = self.table.currentColumn()
        
        # Check if this cell is part of a span
        rowSpan = self.table.rowSpan(row, col)
        colSpan = self.table.columnSpan(row, col)
        
        if rowSpan <= 1 and colSpan <= 1:
            return  # Not a merged cell
        
        # Store the current content and spans for undo
        content = self.table.item(row, col).text() if self.table.item(row, col) else ""
        
        self.history.append({
            'type': 'split_cells',
            'cell': (row, col),
            'rowSpan': rowSpan,
            'colSpan': colSpan,
            'content': content
        })
        
        # Remove the span
        self.table.setSpan(row, col, 1, 1)
        
        # Initialize cells in the range
        for r in range(row, row + rowSpan):
            for c in range(col, col + colSpan):
                if r == row and c == col:
                    continue  # Skip the top-left cell
                item = QTableWidgetItem("")
                self.table.setItem(r, c, item)

    def get_range_data(self, top, left, bottom, right):
        """Get all data in the specified range for undo/redo operations"""
        data = []
        for row in range(top, bottom + 1):
            row_data = []
            for col in range(left, right + 1):
                item = self.table.item(row, col)
                if item:
                    cell_data = {
                        'text': item.text(),
                        'font': item.font().toString(),
                        'bg_color': item.background().color().name(),
                        'fg_color': item.foreground().color().name(),
                        'alignment': item.textAlignment(),
                    }
                    row_data.append(cell_data)
                else:
                    row_data.append(None)
            data.append(row_data)
        return data

    def add_comment_to_cell(self):
        """Add a comment to the currently selected cell"""
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        
        item = selected_items[0]  # Get the first selected cell
        row = item.row()
        col = item.column()
        
        # Get existing comment if any
        existing_comment = ""
        if hasattr(item, "comment"):
            existing_comment = item.comment
        
        # Show dialog to get comment text
        from PyQt5.QtWidgets import QInputDialog
        comment, ok = QInputDialog.getMultiLineText(
            self, "Cell Comment", "Enter comment:", existing_comment)
        
        if ok:
            # Store the comment on the item
            item.comment = comment
            
            # Add visual indicator that the cell has a comment
            if comment:
                # Add a small triangle to the top-right corner of the cell
                indicator = QPixmap(10, 10)
                indicator.fill(QColor(0, 0, 0, 0))  # Transparent background
                painter = QPainter(indicator)
                painter.setBrush(QColor(255, 0, 0))  # Red triangle
                painter.drawPolygon([
                    QPoint(0, 0),
                    QPoint(10, 0),
                    QPoint(10, 10)
                ])
                painter.end()
                
                item.setIcon(QIcon(indicator))
                
                # Set tooltip to show the comment
                item.setToolTip(comment)
            else:
                # Remove indicator and tooltip if comment is empty
                item.setIcon(QIcon())
                item.setToolTip("")
                
            # Add to history for undo/redo
            self.history.append({
                'type': 'cell_comment',
                'row': row,
                'column': col,
                'old_comment': existing_comment,
                'new_comment': comment
            })
            
    def freeze_panes(self, row, column):
        """Freeze panes at the specified row and column"""
        # This would require implementing a custom view or using QSplitter
        # For now, we'll just remember the frozen position
        self.frozen_row = row
        self.frozen_column = column
        self.statusMessage.emit(f"Panes frozen at row {row+1}, column {column+1}")
        
    def create_named_range(self, name, range_str):
        """Create a named range for the selected cells"""
        if not hasattr(self, 'named_ranges'):
            self.named_ranges = {}
            
        # Store the range
        self.named_ranges[name] = range_str
        
    def autofit_columns(self):
        """Auto-fit column widths based on content"""
        for col in range(self.table.columnCount()):
            max_width = 0
            for row in range(self.table.rowCount()):
                item = self.table.item(row, col)
                if item:
                    width = self.table.fontMetrics().width(item.text()) + 10  # Add some padding
                    max_width = max(max_width, width)
            
            if max_width > 0:
                self.table.setColumnWidth(col, max_width)
                
    def autofit_rows(self):
        """Auto-fit row heights based on content"""
        for row in range(self.table.rowCount()):
            max_height = 0
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item:
                    text_lines = item.text().split('\n')
                    height = len(text_lines) * self.table.fontMetrics().height() + 5
                    max_height = max(max_height, height)
                    
            if max_height > 0:
                self.table.setRowHeight(row, max_height)

    def detect_and_apply_pattern_fill(self):
        """Detect patterns and fill the selected range automatically"""
        selected_ranges = self.table.selectedRanges()
        if not selected_ranges:
            return
        
        range_ = selected_ranges[0]
        
        # Check if we have enough data to detect a pattern
        if range_.rowCount() <= 1 and range_.columnCount() <= 1:
            return
        
        # Check if this is a vertical or horizontal fill
        if range_.rowCount() > 1 and range_.columnCount() == 1:
            direction = "vertical"
            source_data = self.get_column_data(range_.leftColumn(), range_.topRow(), range_.topRow() + 1)
            target_cells = [(row, range_.leftColumn()) for row in range(range_.topRow() + 2, range_.bottomRow() + 1)]
        elif range_.rowCount() == 1 and range_.columnCount() > 1:
            direction = "horizontal"
            source_data = self.get_row_data(range_.topRow(), range_.leftColumn(), range_.leftColumn() + 1)
            target_cells = [(range_.topRow(), col) for col in range(range_.leftColumn() + 2, range_.rightColumn() + 1)]
        else:
            # Try to determine if the pattern is more likely vertical or horizontal
            # For simplicity, we'll use the first two cells in each direction
            top_left = self.table.item(range_.topRow(), range_.leftColumn())
            below = self.table.item(range_.topRow() + 1, range_.leftColumn())
            right = self.table.item(range_.topRow(), range_.leftColumn() + 1)
            
            if below and right:
                # Check which has a clearer pattern - simple check for numeric progression
                try:
                    top_left_val = float(top_left.text())
                    below_val = float(below.text())
                    right_val = float(right.text())
                    
                    # Check if horizontal or vertical increment is clearer
                    vertical_diff = abs(below_val - top_left_val)
                    horizontal_diff = abs(right_val - top_left_val)
                    
                    if vertical_diff < horizontal_diff:
                        direction = "vertical"
                    else:
                        direction = "horizontal"
                except (ValueError, AttributeError):
                    # If not numeric, default to horizontal fill
                    direction = "horizontal"
            else:
                direction = "horizontal"  # Default
                
            if direction == "vertical":
                source_data = self.get_column_data(range_.leftColumn(), range_.topRow(), range_.topRow() + 1)
                target_cells = [(row, col) for row in range(range_.topRow() + 1, range_.bottomRow() + 1) 
                              for col in range(range_.leftColumn(), range_.rightColumn() + 1)]
            else:
                source_data = self.get_row_data(range_.topRow(), range_.leftColumn(), range_.leftColumn() + 1)
                target_cells = [(row, col) for row in range(range_.topRow(), range_.bottomRow() + 1) 
                              for col in range(range_.leftColumn() + 1, range_.rightColumn() + 1)]
        
        # Try to detect pattern
        pattern = self.detect_pattern(source_data)
        if not pattern:
            return
        
        # Apply the pattern to target cells
        self.apply_pattern_to_cells(pattern, target_cells)
        
    def get_column_data(self, column, start_row, end_row):
        """Get data from a column range"""
        data = []
        for row in range(start_row, end_row):
            item = self.table.item(row, column)
            if item:
                data.append(item.text())
            else:
                data.append("")
        return data
        
    def get_row_data(self, row, start_col, end_col):
        """Get data from a row range"""
        data = []
        for col in range(start_col, end_col):
            item = self.table.item(row, col)
            if item:
                data.append(item.text())
            else:
                data.append("")
        return data
        
    def detect_pattern(self, data):
        """Detect pattern in a data series"""
        # Basic pattern detections
        patterns = []
        
        # Try numeric patterns first
        try:
            numeric_data = [float(value) for value in data if value]
            if len(numeric_data) >= 2:
                # Check for arithmetic progression (addition/subtraction)
                differences = [numeric_data[i] - numeric_data[i-1] for i in range(1, len(numeric_data))]
                if all(abs(diff - differences[0]) < 0.0001 for diff in differences):
                    patterns.append({
                        'type': 'arithmetic',
                        'difference': differences[0],
                        'last_value': numeric_data[-1]
                    })
                
                # Check for geometric progression (multiplication)
                ratios = [numeric_data[i] / numeric_data[i-1] for i in range(1, len(numeric_data)) 
                         if abs(numeric_data[i-1]) > 0.0001]
                if ratios and all(abs(ratio - ratios[0]) < 0.0001 for ratio in ratios):
                    patterns.append({
                        'type': 'geometric',
                        'ratio': ratios[0],
                        'last_value': numeric_data[-1]
                    })
        except (ValueError, ZeroDivisionError):
            pass
        
        # Try date patterns
        try:
            from datetime import datetime
            date_formats = ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%d-%m-%Y']
            
            for fmt in date_formats:
                try:
                    date_data = [datetime.strptime(value, fmt) for value in data if value]
                    if len(date_data) >= 2:
                        # Check for consistent day/month/year increments
                        deltas = [(date_data[i] - date_data[i-1]).days for i in range(1, len(date_data))]
                        if all(delta == deltas[0] for delta in deltas):
                            patterns.append({
                                'type': 'date',
                                'format': fmt,
                                'delta_days': deltas[0],
                                'last_date': date_data[-1]
                            })
                            break  # Found a working date format
                except ValueError:
                    continue  # Try next format
        except ImportError:
            pass
        
        # Try text patterns (like days of week, months)
        days = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']
        months = ['january', 'february', 'march', 'april', 'may', 'june', 
                 'july', 'august', 'september', 'october', 'november', 'december']
        
        # Check for days of week
        lower_data = [value.lower() for value in data if value]
        if all(value in days for value in lower_data):
            days_indices = [days.index(value.lower()) for value in lower_data]
            delta = (days_indices[1] - days_indices[0]) % 7
            if len(days_indices) >= 2 and all((days_indices[i] - days_indices[i-1]) % 7 == delta for i in range(1, len(days_indices))):
                patterns.append({
                    'type': 'day_of_week',
                    'values': days,
                    'delta': delta,
                    'last_index': days_indices[-1]
                })
        
        # Check for months
        if all(value.lower() in months for value in lower_data):
            months_indices = [months.index(value.lower()) for value in lower_data]
            delta = (months_indices[1] - months_indices[0]) % 12
            if len(months_indices) >= 2 and all((months_indices[i] - months_indices[i-1]) % 12 == delta for i in range(1, len(months_indices))):
                patterns.append({
                    'type': 'month',
                    'values': months,
                    'delta': delta,
                    'last_index': months_indices[-1]
                })
        
        # Return the most likely pattern (for simplicity, the first one found)
        return patterns[0] if patterns else None
        
    def apply_pattern_to_cells(self, pattern, target_cells):
        """Apply detected pattern to target cells"""
        from datetime import datetime, timedelta
        
        # Block signals during update
        self.table.blockSignals(True)
        
        # Apply pattern based on type
        if pattern['type'] == 'arithmetic':
            next_value = pattern['last_value']
            for row, col in target_cells:
                next_value += pattern['difference']
                item = self.table.item(row, col)
                if not item:
                    item = QTableWidgetItem()
                    self.table.setItem(row, col, item)
                item.setText(str(next_value))
        
        elif pattern['type'] == 'geometric':
            next_value = pattern['last_value']
            for row, col in target_cells:
                next_value *= pattern['ratio']
                item = self.table.item(row, col)
                if not item:
                    item = QTableWidgetItem()
                    self.table.setItem(row, col, item)
                item.setText(str(next_value))
        
        elif pattern['type'] == 'date':
            next_date = pattern['last_date']
            for row, col in target_cells:
                next_date += timedelta(days=pattern['delta_days'])
                item = self.table.item(row, col)
                if not item:
                    item = QTableWidgetItem()
                    self.table.setItem(row, col, item)
                item.setText(next_date.strftime(pattern['format']))
        
        elif pattern['type'] in ('day_of_week', 'month'):
            values = pattern['values']
            last_index = pattern['last_index']
            delta = pattern['delta']
            
            for row, col in target_cells:
                next_index = (last_index + delta) % len(values)
                last_index = next_index
                
                item = self.table.item(row, col)
                if not item:
                    item = QTableWidgetItem()
                    self.table.setItem(row, col, item)
                
                # Preserve case from the original
                original_value = self.table.item(row-1, col).text() if row > 0 else ""
                if original_value and original_value[0].isupper():
                    item.setText(values[next_index].capitalize())
                else:
                    item.setText(values[next_index])
        
        # Restore signals
        self.table.blockSignals(False)

    def get_cell_raw_value(self, row, column):
        """Get the raw value (formula) from a cell, not the displayed result"""
        cell_key = f"{row},{column}"
        item = self.table.item(row, column)
        if not item:
            return None
            
        # Check if this cell has a formula stored
        if hasattr(self, 'cell_formulas') and cell_key in self.cell_formulas:
            return self.cell_formulas[cell_key]
        else:
            # No formula, return the displayed text
            return item.text()

    def set_formula_and_display(self, row, column, formula, display_value):
        """Set both the formula and display value for a cell"""
        if not hasattr(self, 'cell_formulas'):
            self.cell_formulas = {}
            
        cell_key = f"{row},{column}"
        self.cell_formulas[cell_key] = formula
        self.cell_display_values[cell_key] = display_value
        
        # Update the display in the table
        item = self.table.item(row, column)
        if not item:
            item = QTableWidgetItem()
            self.table.setItem(row, column, item)
        
        # Temporarily disconnect itemChanged signal to avoid recursion
        self.table.blockSignals(True)
        item.setText(str(display_value))
        self.table.blockSignals(False)
        
        # Set tooltip to show formula
        item.setToolTip(f"Formula: {formula}\nResult: {display_value}")
        
        # Add this to history for undo/redo
        self.history.append({
            'type': 'cell_edit',
            'row': row,
            'column': column,
            'old_value': item.text() if item else "",
            'new_value': formula,
            'display_value': display_value
        })

    def eventFilter(self, source, event):
        """Handle special key events like Enter"""
        if source is self.table and event.type() == QEvent.KeyPress:
            key = event.key()
            if key == Qt.Key_Return or key == Qt.Key_Enter:
                # Move to the cell below on Enter press
                current_row = self.table.currentRow()
                current_col = self.table.currentColumn()
                
                # If we're at the last row, optionally add new rows
                if current_row >= self.table.rowCount() - 1:
                    self.table.setRowCount(self.table.rowCount() + 5)  # Add more rows dynamically
                    
                # Move to the cell below
                self.table.setCurrentCell(current_row + 1, current_col)
                return True  # Event handled
        
        # Default event handling
        return super().eventFilter(source, event)

    def item(self, row, column):
        """Return the item at the specified row and column.
        This forwards the call to the internal QTableWidget."""
        return self.table.item(row, column)

    def currentRow(self):
        """Return the current row.
        This forwards the call to the internal QTableWidget."""
        return self.table.currentRow()
        
    def currentColumn(self):
        """Return the current column.
        This forwards the call to the internal QTableWidget."""
        return self.table.currentColumn()
        
    def setCurrentCell(self, row, column):
        """Set the current cell.
        This forwards the call to the internal QTableWidget."""
        self.table.setCurrentCell(row, column)
        
    def rowCount(self):
        """Return the number of rows.
        This forwards the call to the internal QTableWidget."""
        return self.table.rowCount()
        
    def columnCount(self):
        """Return the number of columns.
        This forwards the call to the internal QTableWidget."""
        return self.table.columnCount()
        
    def selectedRanges(self):
        """Return the selected ranges.
        This forwards the call to the internal QTableWidget."""
        return self.table.selectedRanges()
        
    def selectedItems(self):
        """Return the selected items.
        This forwards the call to the internal QTableWidget."""
        return self.table.selectedItems()
        
    def clearSelection(self):
        """Clear the current selection.
        This forwards the call to the internal QTableWidget."""
        self.table.clearSelection()
        
    def selectAll(self):
        """Select all cells.
        This forwards the call to the internal QTableWidget."""
        self.table.selectAll()
        
    def selectRow(self, row):
        """Select an entire row.
        This forwards the call to the internal QTableWidget."""
        self.table.selectRow(row)
        
    def selectColumn(self, column):
        """Select an entire column.
        This forwards the call to the internal QTableWidget."""
        self.table.selectColumn(column)
        
    def hideRow(self, row):
        """Hide a row.
        This forwards the call to the internal QTableWidget."""
        self.table.hideRow(row)
        
    def showRow(self, row):
        """Show a row.
        This forwards the call to the internal QTableWidget."""
        self.table.showRow(row)
        
    def hideColumn(self, column):
        """Hide a column.
        This forwards the call to the internal QTableWidget."""
        self.table.hideColumn(column)
        
    def showColumn(self, column):
        """Show a column.
        This forwards the call to the internal QTableWidget."""
        self.table.showColumn(column)


class ConditionalFormatDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.sheet_view = parent
        self.setWindowTitle("Conditional Formatting")
        self.setMinimumWidth(500)
        
        self.layout = QVBoxLayout(self)
        
        # Create rule editor area
        self.create_rule_editor()
        
        # Create button box
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.layout.addWidget(self.button_box)
        
        # Load existing rules
        self.load_existing_rules()

    def create_rule_editor(self):
        group_box = QGroupBox("Formatting Rules")
        layout = QVBoxLayout()
        
        # Rule type selector
        rule_type_layout = QHBoxLayout()
        rule_type_label = QLabel("Rule Type:")
        self.rule_type_combo = QComboBox()
        self.rule_type_combo.addItems(["Cell Value", "Text Contains"])
        rule_type_layout.addWidget(rule_type_label)
        rule_type_layout.addWidget(self.rule_type_combo)
        layout.addLayout(rule_type_layout)
        
        # Condition type selector
        condition_layout = QHBoxLayout()
        condition_label = QLabel("Condition:")
        self.condition_combo = QComboBox()
        self.condition_combo.addItems(["Equal to", "Not equal to", "Greater than", "Less than", "Contains", "Does not contain"])
        condition_layout.addWidget(condition_label)
        condition_layout.addWidget(self.condition_combo)
        layout.addLayout(condition_layout)
        
        # Value input
        value_layout = QHBoxLayout()
        value_label = QLabel("Value:")
        self.value_edit = QLineEdit()
        value_layout.addWidget(value_label)
        value_layout.addWidget(self.value_edit)
        layout.addLayout(value_layout)
        
        # Format options
        format_options_layout = QHBoxLayout()
        
        # Background color
        bg_color_label = QLabel("Background Color:")
        self.bg_color_button = QPushButton()
        self.bg_color_button.setFixedSize(30, 20)
        self.bg_color = QColor(255, 255, 255)  # Default white
        self.update_bg_color_button()
        self.bg_color_button.clicked.connect(self.select_bg_color)
        format_options_layout.addWidget(bg_color_label)
        format_options_layout.addWidget(self.bg_color_button)
        
        format_options_layout.addSpacing(20)
        
        # Text color
        text_color_label = QLabel("Text Color:")
        self.text_color_button = QPushButton()
        self.text_color_button.setFixedSize(30, 20)
        self.text_color = QColor(0, 0, 0)  # Default black
        self.update_text_color_button()
        self.text_color_button.clicked.connect(self.select_text_color)
        format_options_layout.addWidget(text_color_label)
        format_options_layout.addWidget(self.text_color_button)
        
        layout.addLayout(format_options_layout)
        
        # Add/Delete rule buttons
        rule_buttons_layout = QHBoxLayout()
        self.add_rule_button = QPushButton("Add Rule")
        self.add_rule_button.clicked.connect(self.add_rule)
        rule_buttons_layout.addWidget(self.add_rule_button)
        
        self.delete_rule_button = QPushButton("Delete Selected Rule")
        self.delete_rule_button.clicked.connect(self.delete_rule)
        rule_buttons_layout.addWidget(self.delete_rule_button)
        
        layout.addLayout(rule_buttons_layout)
        
        # Rule list (shows existing rules)
        self.rule_list = QTableWidget(0, 4)
        self.rule_list.setHorizontalHeaderLabels(["Condition", "Value", "Background", "Text Color"])
        self.rule_list.horizontalHeader().setStretchLastSection(True)
        self.rule_list.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.rule_list.setEditTriggers(QAbstractItemView.NoEditTriggers)
        layout.addWidget(self.rule_list)
        
        group_box.setLayout(layout)
        self.layout.addWidget(group_box)

    def update_bg_color_button(self):
        """Update the background color button appearance"""
        style = f"background-color: {self.bg_color.name()};"
        self.bg_color_button.setStyleSheet(style)

    def update_text_color_button(self):
        """Update the text color button appearance"""
        style = f"background-color: {self.text_color.name()};"
        self.text_color_button.setStyleSheet(style)

    def select_bg_color(self):
        """Open color picker for background color"""
        color = QColorDialog.getColor(self.bg_color, self)
        if color.isValid():
            self.bg_color = color
            self.update_bg_color_button()

    def select_text_color(self):
        """Open color picker for text color"""
        color = QColorDialog.getColor(self.text_color, self)
        if color.isValid():
            self.text_color = color
            self.update_text_color_button()

    def add_rule(self):
        """Add a new formatting rule"""
        # Get the condition information
        condition_type = self.condition_combo.currentText().lower().replace(' ', '_')
        value = self.value_edit.text()
        
        if not value:
            QMessageBox.warning(self, "Missing Value", "Please enter a value for the condition.")
            return
            
        # Create rule dictionary
        rule = {
            'condition': condition_type,
            'value': value,
            'bg_color': self.bg_color.name() if self.bg_color.isValid() else None,
            'text_color': self.text_color.name() if self.text_color.isValid() else None
        }
        
        # Add rule to sheet view
        self.sheet_view.conditional_formatting_rules.append(rule)
        
        # Update the rule list display
        self.update_rule_list()
        
        # Clear inputs for next rule
        self.value_edit.clear()
        
    def delete_rule(self):
        """Delete the selected rule"""
        selected_rows = self.rule_list.selectedIndexes()
        if not selected_rows:
            return
            
        row = selected_rows[0].row()
        if row >= 0 and row < len(self.sheet_view.conditional_formatting_rules):
            del self.sheet_view.conditional_formatting_rules[row]
            self.update_rule_list()

    def update_rule_list(self):
        """Update the displayed rule list"""
        rules = self.sheet_view.conditional_formatting_rules
        self.rule_list.setRowCount(len(rules))
        
        for i, rule in enumerate(rules):
            # Condition
            condition_text = rule['condition'].replace('_', ' ').title()
            self.rule_list.setItem(i, 0, QTableWidgetItem(condition_text))
            
            # Value
            self.rule_list.setItem(i, 1, QTableWidgetItem(rule['value']))
            
            # Background color
            bg_item = QTableWidgetItem()
            if rule['bg_color']:
                bg_item.setBackground(QColor(rule['bg_color']))
            self.rule_list.setItem(i, 2, bg_item)
            
            # Text color
            text_item = QTableWidgetItem()
            if rule['text_color']:
                text_item.setBackground(QColor(rule['text_color']))
            self.rule_list.setItem(i, 3, text_item)

    def load_existing_rules(self):
        """Load and display existing formatting rules"""
        self.update_rule_list()
        
    def accept(self):
        """Apply the formatting rules and close the dialog"""
        self.sheet_view.apply_conditional_formatting_to_all_cells()
        super().accept()


class FindDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.sheet_view = parent
        self.setWindowTitle("Find")
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout(self)
        
        # Find what
        find_layout = QHBoxLayout()
        find_label = QLabel("Find what:")
        self.find_edit = QLineEdit()
        self.find_edit.setMinimumWidth(200)
        find_layout.addWidget(find_label)
        find_layout.addWidget(self.find_edit)
        layout.addLayout(find_layout)
        
        # Options
        options_group = QGroupBox("Options")
        options_layout = QVBoxLayout()
        
        self.match_case = QCheckBox("Match case")
        options_layout.addWidget(self.match_case)
        
        self.match_whole_word = QCheckBox("Match whole word")
        options_layout.addWidget(self.match_whole_word)
        
        self.search_formulas = QCheckBox("Search in formulas")
        self.search_formulas.setChecked(True)
        options_layout.addWidget(self.search_formulas)
        
        options_group.setLayout(options_layout)
        layout.addWidget(options_group)
        
        # Buttons
        button_layout = QHBoxLayout()
        self.find_button = QPushButton("Find Next")
        self.find_button.clicked.connect(self.find_next)
        button_layout.addWidget(self.find_button)
        
        self.close_button = QPushButton("Close")
        self.close_button.clicked.connect(self.reject)
        button_layout.addWidget(self.close_button)
        
        layout.addLayout(button_layout)

    def find_next(self):
        """Find the next occurrence of the text"""
        text = self.find_edit.text()
        if not text:
            return
            
        found = self.sheet_view.find_next(
            text,
            match_case=self.match_case.isChecked(),
            match_whole_word=self.match_whole_word.isChecked(),
            search_in_formulas=self.search_formulas.isChecked()
        )
        
        if not found:
            QMessageBox.information(self, "Find", f"Cannot find '{text}'")


class ReplaceDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.sheet_view = parent
        self.setWindowTitle("Replace")
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout(self)
        
        # Find what
        find_layout = QHBoxLayout()
        find_label = QLabel("Find what:")
        self.find_edit = QLineEdit()
        self.find_edit.setMinimumWidth(200)
        find_layout.addWidget(find_label)
        find_layout.addWidget(self.find_edit)
        layout.addLayout(find_layout)
        
        # Replace with
        replace_layout = QHBoxLayout()
        replace_label = QLabel("Replace with:")
        self.replace_edit = QLineEdit()
        self.replace_edit.setMinimumWidth(200)
        replace_layout.addWidget(replace_label)
        replace_layout.addWidget(self.replace_edit)
        layout.addLayout(replace_layout)
        
        # Options
        options_group = QGroupBox("Options")
        options_layout = QVBoxLayout()
        
        self.match_case = QCheckBox("Match case")
        options_layout.addWidget(self.match_case)
        
        self.match_whole_word = QCheckBox("Match whole word")
        options_layout.addWidget(self.match_whole_word)
        
        self.search_formulas = QCheckBox("Search in formulas")
        self.search_formulas.setChecked(True)
        options_layout.addWidget(self.search_formulas)
        
        options_group.setLayout(options_layout)
        layout.addWidget(options_group)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        self.find_button = QPushButton("Find Next")
        self.find_button.clicked.connect(self.find_next)
        button_layout.addWidget(self.find_button)
        
        self.replace_button = QPushButton("Replace")
        self.replace_button.clicked.connect(self.replace)
        button_layout.addWidget(self.replace_button)
        
        self.replace_all_button = QPushButton("Replace All")
        self.replace_all_button.clicked.connect(self.replace_all)
        button_layout.addWidget(self.replace_all_button)
        
        self.close_button = QPushButton("Close")
        self.close_button.clicked.connect(self.reject)
        button_layout.addWidget(self.close_button)
        
        layout.addLayout(button_layout)

    def find_next(self):
        """Find the next occurrence of the text"""
        text = self.find_edit.text()
        if not text:
            return False
            
        found = self.sheet_view.find_next(
            text,
            match_case=self.match_case.isChecked(),
            match_whole_word=self.match_whole_word.isChecked(),
            search_in_formulas=self.search_formulas.isChecked()
        )
        
        if not found:
            QMessageBox.information(self, "Find", f"Cannot find '{text}'")
            
        return found

    def replace(self):
        """Replace the current found item and find the next"""
        find_text = self.find_edit.text()
        replace_text = self.replace_edit.text()
        
        if not find_text:
            return
            
        # Replace current
        replaced = self.sheet_view.replace_current(
            find_text,
            replace_text,
            match_case=self.match_case.isChecked(),
            match_whole_word=self.match_whole_word.isChecked(),
            search_in_formulas=self.search_formulas.isChecked()
        )
        
        # Find next
        self.find_next()

    def replace_all(self):
        """Replace all occurrences"""
        find_text = self.find_edit.text()
        replace_text = self.replace_edit.text()
        
        if not find_text:
            return
            
        count = self.sheet_view.replace_all(
            find_text,
            replace_text,
            match_case=self.match_case.isChecked(),
            match_whole_word=self.match_whole_word.isChecked(),
            search_in_formulas=self.search_formulas.isChecked()
        )
        
        QMessageBox.information(self, "Replace", f"Replaced {count} occurrences of '{find_text}'")


class FilterDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.sheet_view = parent
        self.setWindowTitle("Filter Data")
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout(self)
        
        # Column selection
        column_layout = QHBoxLayout()
        column_label = QLabel("Column:")
        self.column_combo = QComboBox()
        
        # Add column options (A, B, C, etc.)
        for col in range(self.sheet_view.table.columnCount()):
            col_name = chr(65 + col) if col < 26 else chr(64 + int(col/26)) + chr(65 + (col % 26))
            self.column_combo.addItem(col_name, col)  # Store column index as user data
            
        column_layout.addWidget(column_label)
        column_layout.addWidget(self.column_combo)
        layout.addLayout(column_layout)
        
        # Condition
        condition_layout = QHBoxLayout()
        condition_label = QLabel("Condition:")
        self.condition_combo = QComboBox()
        self.condition_combo.addItems([
            "Equal to", "Not equal to", "Contains", "Does not contain",
            "Greater than", "Less than"
        ])
        condition_layout.addWidget(condition_label)
        condition_layout.addWidget(self.condition_combo)
        layout.addLayout(condition_layout)
        
        # Filter value
        value_layout = QHBoxLayout()
        value_label = QLabel("Value:")
        self.value_edit = QLineEdit()
        value_layout.addWidget(value_label)
        value_layout.addWidget(self.value_edit)
        layout.addLayout(value_layout)
        
        # Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def get_filter_options(self):
        """Get the selected filter options"""
        column_index = self.column_combo.currentData()
        condition = self.condition_combo.currentText().lower().replace(' ', '_')
        value = self.value_edit.text()
        
        return {
            'column': column_index,
            'condition': condition,
            'value': value
        }