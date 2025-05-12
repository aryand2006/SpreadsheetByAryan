from PyQt5.QtWidgets import QPushButton, QLineEdit, QLabel, QComboBox, QWidget, QVBoxLayout, QCompleter, QToolTip
from PyQt5.QtCore import Qt, QPoint, QEvent

class CustomButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setStyleSheet("background-color: #4CAF50; color: white; font-size: 16px;")

class CustomLineEdit(QLineEdit):
    """Base custom line edit widget"""
    def __init__(self, parent=None):
        super().__init__(parent)

class CustomLabel(QLabel):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setStyleSheet("font-size: 14px; font-weight: bold;")

class CustomComboBox(QComboBox):
    def __init__(self, items, parent=None):
        super().__init__(parent)
        self.addItems(items)
        self.setStyleSheet("font-size: 14px;")

class FormulaLineEdit(QLineEdit):
    """Enhanced line edit for formula input with autocomplete and help"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.calculator = None
        self.completer = None
        self.function_list = []
        self.setup_completer()
        
        # Connect signals
        self.textChanged.connect(self.check_for_function_help)
    
    def set_calculator(self, calculator):
        """Set the calculator engine to access function definitions"""
        self.calculator = calculator
        if calculator and hasattr(calculator, 'functions'):
            self.function_list = list(calculator.functions.keys())
            self.setup_completer()
    
    def setup_completer(self):
        """Set up function name autocomplete"""
        # Common spreadsheet functions if calculator not set yet
        if not self.function_list:
            self.function_list = ["SUM", "AVERAGE", "COUNT", "MAX", "MIN", "IF", "VLOOKUP", 
                                 "HLOOKUP", "DATE", "NOW", "TODAY", "ROUND", "ABS", "SQRT"]
        
        self.completer = QCompleter(self.function_list)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.completer.setFilterMode(Qt.MatchContains)
        self.setCompleter(self.completer)
    
    def check_for_function_help(self, text):
        """Show help tooltip when typing a function"""
        if text.startswith('='):
            # Check for function name being typed
            formula = text[1:].strip()
            if '(' in formula:
                func_name = formula[:formula.find('(')].upper()
                if self.calculator and func_name in self.calculator.functions:
                    # Show tooltip with function help
                    QToolTip.showText(self.mapToGlobal(QPoint(0, -30)), 
                                    f"{func_name}: {self.get_function_help(func_name)}")

    def get_function_help(self, func_name):
        """Get help text for a function"""
        # Basic help for common functions
        help_dict = {
            "SUM": "SUM(range) - Adds all numbers in the range",
            "AVERAGE": "AVERAGE(range) - Calculates the average of numbers",
            "COUNT": "COUNT(range) - Counts non-empty cells in the range",
            "MAX": "MAX(range) - Finds the maximum value",
            "MIN": "MIN(range) - Finds the minimum value",
            "IF": "IF(condition, value_if_true, value_if_false) - Conditional logic",
            "VLOOKUP": "VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup]) - Vertical lookup",
            "NOW": "NOW() - Returns current date and time",
            "TODAY": "TODAY() - Returns current date"
        }
        
        return help_dict.get(func_name, f"{func_name}() - No help available")

    def event(self, event):
        """Override event handler to add custom keyboard shortcuts"""
        if event.type() == QEvent.KeyPress and event.modifiers() == Qt.ControlModifier:
            if event.key() == Qt.Key_Space:
                # Ctrl+Space shows function list
                if self.text().startswith('='):
                    self.completer.setCompletionPrefix("")
                    self.completer.complete()
                    return True
                    
        return super().event(event)