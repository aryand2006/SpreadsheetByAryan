from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, 
    QLineEdit, QPushButton, QGroupBox, QRadioButton,
    QDialogButtonBox, QCheckBox, QSpinBox, QDoubleSpinBox
)
from PyQt5.QtCore import Qt

class DataValidationDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Data Validation")
        self.setMinimumWidth(450)
        
        self.main_layout = QVBoxLayout(self)
        
        # Validation type selection
        self.create_validation_type_group()
        
        # Validation criteria
        self.create_criteria_group()
        
        # Input message
        self.create_input_message_group()
        
        # Error alert
        self.create_error_alert_group()
        
        # Buttons
        self.button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | 
            QDialogButtonBox.Cancel
        )
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.main_layout.addWidget(self.button_box)
        
        # Connect validation type to criteria options
        self.validation_type_combo.currentIndexChanged.connect(self.update_criteria_options)
        
        # Initialize with default options
        self.update_criteria_options(0)
    
    def create_validation_type_group(self):
        group = QGroupBox("Validation Type")
        layout = QVBoxLayout()
        
        type_layout = QHBoxLayout()
        type_label = QLabel("Allow:")
        self.validation_type_combo = QComboBox()
        self.validation_type_combo.addItems([
            "Any value",
            "Whole number",
            "Decimal",
            "List",
            "Date",
            "Time",
            "Text length",
            "Custom formula"
        ])
        type_layout.addWidget(type_label)
        type_layout.addWidget(self.validation_type_combo)
        layout.addLayout(type_layout)
        
        self.ignore_blanks = QCheckBox("Ignore blank")
        self.ignore_blanks.setChecked(True)
        layout.addWidget(self.ignore_blanks)
        
        group.setLayout(layout)
        self.main_layout.addWidget(group)
    
    def create_criteria_group(self):
        self.criteria_group = QGroupBox("Criteria")
        self.criteria_layout = QVBoxLayout()
        
        # Criteria options will be dynamically added based on validation type
        self.criteria_combo = QComboBox()
        
        # Create input widgets for various criteria types
        self.create_criteria_widgets()
        
        self.criteria_group.setLayout(self.criteria_layout)
        self.main_layout.addWidget(self.criteria_group)
    
    def create_criteria_widgets(self):
        # Operator combo box
        operator_layout = QHBoxLayout()
        operator_label = QLabel("Data:")
        operator_layout.addWidget(operator_label)
        operator_layout.addWidget(self.criteria_combo)
        self.criteria_layout.addLayout(operator_layout)
        
        # Value inputs
        self.value_layout = QHBoxLayout()
        self.value_label1 = QLabel("Minimum:")
        self.value_edit1 = QLineEdit()
        self.value_label2 = QLabel("Maximum:")
        self.value_edit2 = QLineEdit()
        
        self.value_layout.addWidget(self.value_label1)
        self.value_layout.addWidget(self.value_edit1)
        self.value_layout.addWidget(self.value_label2)
        self.value_layout.addWidget(self.value_edit2)
        
        self.criteria_layout.addLayout(self.value_layout)
        
        # List source input
        self.list_layout = QHBoxLayout()
        self.list_label = QLabel("Source:")
        self.list_edit = QLineEdit()
        self.list_layout.addWidget(self.list_label)
        self.list_layout.addWidget(self.list_edit)
        
        # Hide by default, will show based on type
        self.criteria_layout.addLayout(self.list_layout)
        self.hide_list_input()
    
    def hide_value_inputs(self):
        self.value_label1.hide()
        self.value_edit1.hide()
        self.value_label2.hide()
        self.value_edit2.hide()
    
    def show_value_inputs(self, label1="Value:", label2=None):
        self.value_label1.setText(label1)
        self.value_label1.show()
        self.value_edit1.show()
        
        if label2:
            self.value_label2.setText(label2)
            self.value_label2.show()
            self.value_edit2.show()
        else:
            self.value_label2.hide()
            self.value_edit2.hide()
    
    def hide_list_input(self):
        self.list_label.hide()
        self.list_edit.hide()
    
    def show_list_input(self):
        self.list_label.show()
        self.list_edit.show()
    
    def update_criteria_options(self, index):
        # Clear and update criteria combo based on validation type
        self.criteria_combo.clear()
        validation_type = self.validation_type_combo.currentText()
        
        if validation_type == "Any value":
            self.criteria_combo.setEnabled(False)
            self.hide_value_inputs()
            self.hide_list_input()
            
        elif validation_type in ["Whole number", "Decimal"]:
            self.criteria_combo.setEnabled(True)
            self.criteria_combo.addItems([
                "between", "not between", "equal to", "not equal to",
                "greater than", "less than", "greater than or equal to",
                "less than or equal to"
            ])
            self.show_value_inputs("Minimum:", "Maximum:")
            self.hide_list_input()
            
        elif validation_type == "List":
            self.criteria_combo.setEnabled(False)
            self.hide_value_inputs()
            self.show_list_input()
            
        elif validation_type in ["Date", "Time"]:
            self.criteria_combo.setEnabled(True)
            self.criteria_combo.addItems([
                "between", "not between", "equal to", "not equal to",
                "greater than", "less than", "greater than or equal to",
                "less than or equal to"
            ])
            self.show_value_inputs("Start date:", "End date:" if validation_type == "Date" else "End time:")
            self.hide_list_input()
            
        elif validation_type == "Text length":
            self.criteria_combo.setEnabled(True)
            self.criteria_combo.addItems([
                "between", "not between", "equal to", "not equal to",
                "greater than", "less than", "greater than or equal to",
                "less than or equal to"
            ])
            self.show_value_inputs("Minimum length:", "Maximum length:")
            self.hide_list_input()
            
        elif validation_type == "Custom formula":
            self.criteria_combo.setEnabled(False)
            self.show_value_inputs("Formula:")
            self.hide_list_input()
    
    def create_input_message_group(self):
        group = QGroupBox("Input Message")
        group.setCheckable(True)
        group.setChecked(False)
        
        layout = QVBoxLayout()
        
        title_layout = QHBoxLayout()
        title_label = QLabel("Title:")
        self.input_title_edit = QLineEdit()
        title_layout.addWidget(title_label)
        title_layout.addWidget(self.input_title_edit)
        layout.addLayout(title_layout)
        
        message_layout = QHBoxLayout()
        message_label = QLabel("Message:")
        self.input_message_edit = QLineEdit()
        message_layout.addWidget(message_label)
        message_layout.addWidget(self.input_message_edit)
        layout.addLayout(message_layout)
        
        group.setLayout(layout)
        self.main_layout.addWidget(group)
    
    def create_error_alert_group(self):
        group = QGroupBox("Error Alert")
        group.setCheckable(True)
        group.setChecked(True)
        
        layout = QVBoxLayout()
        
        style_layout = QHBoxLayout()
        style_label = QLabel("Style:")
        self.error_style_combo = QComboBox()
        self.error_style_combo.addItems(["Stop", "Warning", "Information"])
        style_layout.addWidget(style_label)
        style_layout.addWidget(self.error_style_combo)
        layout.addLayout(style_layout)
        
        title_layout = QHBoxLayout()
        title_label = QLabel("Title:")
        self.error_title_edit = QLineEdit()
        self.error_title_edit.setText("Invalid Data")
        title_layout.addWidget(title_label)
        title_layout.addWidget(self.error_title_edit)
        layout.addLayout(title_layout)
        
        message_layout = QHBoxLayout()
        message_label = QLabel("Message:")
        self.error_message_edit = QLineEdit()
        self.error_message_edit.setText("The value you entered is not valid.")
        message_layout.addWidget(message_label)
        message_layout.addWidget(self.error_message_edit)
        layout.addLayout(message_layout)
        
        group.setLayout(layout)
        self.main_layout.addWidget(group)
    
    def get_validation_settings(self):
        """Get all validation settings from the dialog"""
        validation_type = self.validation_type_combo.currentText()
        criteria = self.criteria_combo.currentText() if self.criteria_combo.isEnabled() else None
        
        settings = {
            'type': validation_type,
            'ignore_blanks': self.ignore_blanks.isChecked(),
            'criteria': criteria,
            'input_message': {
                'enabled': self.findChild(QGroupBox, "Input Message").isChecked(),
                'title': self.input_title_edit.text(),
                'message': self.input_message_edit.text()
            },
            'error_alert': {
                'enabled': self.findChild(QGroupBox, "Error Alert").isChecked(),
                'style': self.error_style_combo.currentText(),
                'title': self.error_title_edit.text(),
                'message': self.error_message_edit.text()
            }
        }
        
        # Add type-specific values
        if validation_type in ["Whole number", "Decimal", "Date", "Time", "Text length"]:
            settings['value1'] = self.value_edit1.text()
            if criteria in ["between", "not between"]:
                settings['value2'] = self.value_edit2.text()
        elif validation_type == "List":
            settings['source'] = self.list_edit.text()
        elif validation_type == "Custom formula":
            settings['formula'] = self.value_edit1.text()
        
        return settings
