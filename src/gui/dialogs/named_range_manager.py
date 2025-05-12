from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QLabel, QPushButton, QLineEdit, QMessageBox, QDialogButtonBox, QHeaderView
)
from PyQt5.QtCore import Qt, QSize

class NamedRangeManager(QDialog):
    """Dialog for managing named ranges in the spreadsheet"""
    def __init__(self, parent=None, named_ranges=None):
        super().__init__(parent)
        self.named_ranges = named_ranges or {}
        self.setWindowTitle("Named Range Manager")
        self.setMinimumSize(500, 400)
        self.setup_ui()
        
    def setup_ui(self):
        """Set up the dialog UI components"""
        layout = QVBoxLayout(self)
        
        # Instructions
        instructions = QLabel(
            "Named ranges let you refer to cells by a meaningful name in formulas.\n"
            "Example: Use 'Sales_Total' instead of 'B15:B28' in formulas."
        )
        layout.addWidget(instructions)
        
        # Named ranges table
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Name", "Range", "Scope"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        layout.addWidget(self.table)
        
        # Add/edit controls
        form_layout = QHBoxLayout()
        
        form_layout.addWidget(QLabel("Name:"))
        self.name_input = QLineEdit()
        form_layout.addWidget(self.name_input)
        
        form_layout.addWidget(QLabel("Range:"))
        self.range_input = QLineEdit()
        form_layout.addWidget(self.range_input)
        
        self.add_button = QPushButton("Add")
        self.add_button.clicked.connect(self.add_range)
        form_layout.addWidget(self.add_button)
        
        layout.addLayout(form_layout)
        
        # Button row
        button_layout = QHBoxLayout()
        self.edit_button = QPushButton("Edit")
        self.edit_button.clicked.connect(self.edit_range)
        button_layout.addWidget(self.edit_button)
        
        self.delete_button = QPushButton("Delete")
        self.delete_button.clicked.connect(self.delete_range)
        button_layout.addWidget(self.delete_button)
        
        layout.addLayout(button_layout)
        
        # Standard dialog buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        # Load existing named ranges
        self.load_named_ranges()
        
    def load_named_ranges(self):
        """Load existing named ranges into the table"""
        self.table.setRowCount(0)
        
        for i, (name, range_data) in enumerate(self.named_ranges.items()):
            self.table.insertRow(i)
            self.table.setItem(i, 0, QTableWidgetItem(name))
            self.table.setItem(i, 1, QTableWidgetItem(range_data.get('range', '')))
            self.table.setItem(i, 2, QTableWidgetItem(range_data.get('scope', 'Workbook')))
    
    def add_range(self):
        """Add a new named range"""
        name = self.name_input.text().strip()
        range_text = self.range_input.text().strip()
        
        if not name or not range_text:
            QMessageBox.warning(self, "Missing Information", 
                              "Please provide both name and range.")
            return
            
        # Validate name (no spaces, starts with letter)
        if not name[0].isalpha() or ' ' in name:
            QMessageBox.warning(self, "Invalid Name", 
                              "Name must start with a letter and contain no spaces.")
            return
            
        # Add to named ranges
        self.named_ranges[name] = {
            'range': range_text,
            'scope': 'Workbook'
        }
        
        # Update table
        self.load_named_ranges()
        
        # Clear inputs
        self.name_input.clear()
        self.range_input.clear()
    
    def edit_range(self):
        """Edit the selected named range"""
        selected_rows = self.table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "No Selection", "Please select a range to edit.")
            return
            
        row = selected_rows[0].row()
        name = self.table.item(row, 0).text()
        
        # Fill edit fields
        self.name_input.setText(name)
        self.range_input.setText(self.named_ranges[name]['range'])
        
        # Remove old entry
        del self.named_ranges[name]
        self.load_named_ranges()
    
    def delete_range(self):
        """Delete the selected named range"""
        selected_rows = self.table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "No Selection", "Please select a range to delete.")
            return
            
        row = selected_rows[0].row()
        name = self.table.item(row, 0).text()
        
        # Confirm deletion
        if QMessageBox.question(self, "Confirm Delete", 
                              f"Delete named range '{name}'?",
                              QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            del self.named_ranges[name]
            self.load_named_ranges()
