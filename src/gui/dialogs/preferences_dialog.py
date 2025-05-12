from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTabWidget, QWidget, QLabel, QComboBox,
    QCheckBox, QSpinBox, QGroupBox, QPushButton, QDialogButtonBox, QColorDialog,
    QLineEdit, QFileDialog, QFormLayout
)
from PyQt5.QtGui import QColor, QFont
from PyQt5.QtCore import QSettings, Qt
from ...utils.config import USER_PREFERENCES, DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE

class PreferencesDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Preferences")
        self.setMinimumWidth(500)
        self.setMinimumHeight(400)
        
        self.settings = QSettings("AryanTech", "PySpreadsheet")
        self.preferences = USER_PREFERENCES.copy()
        
        # Load current preferences from settings
        self.load_settings()
        
        # Create the main layout
        layout = QVBoxLayout(self)
        
        # Create tabs
        self.tab_widget = QTabWidget()
        layout.addWidget(self.tab_widget)
        
        # General tab
        self.create_general_tab()
        
        # Appearance tab
        self.create_appearance_tab()
        
        # Editor tab
        self.create_editor_tab()
        
        # Create buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def load_settings(self):
        """Load preferences from QSettings"""
        # Load each preference from settings, falling back to the defaults
        for key, default_value in self.preferences.items():
            value = self.settings.value(f"preferences/{key}", default_value)
            
            # Convert string values to appropriate types if needed
            if isinstance(default_value, bool):
                if isinstance(value, str):
                    value = value.lower() == 'true'
                else:
                    value = bool(value)
            
            self.preferences[key] = value

    def save_settings(self):
        """Save preferences to QSettings"""
        for key, value in self.preferences.items():
            self.settings.setValue(f"preferences/{key}", value)

    def create_general_tab(self):
        """Create the general preferences tab"""
        general_tab = QWidget()
        layout = QVBoxLayout(general_tab)
        
        # Auto-save options
        auto_save_group = QGroupBox("Auto Save")
        auto_save_layout = QVBoxLayout()
        
        self.auto_save_enabled = QCheckBox("Enable auto-save")
        self.auto_save_enabled.setChecked(self.preferences.get("auto_save_enabled", True))
        auto_save_layout.addWidget(self.auto_save_enabled)
        
        auto_save_interval_layout = QHBoxLayout()
        auto_save_interval_layout.addWidget(QLabel("Auto-save interval (minutes):"))
        self.auto_save_interval = QSpinBox()
        self.auto_save_interval.setRange(1, 60)
        self.auto_save_interval.setValue(self.preferences.get("auto_save_interval", 5))
        auto_save_interval_layout.addWidget(self.auto_save_interval)
        auto_save_layout.addLayout(auto_save_interval_layout)
        
        auto_save_group.setLayout(auto_save_layout)
        layout.addWidget(auto_save_group)
        
        # Default file location
        file_location_group = QGroupBox("File Locations")
        file_location_layout = QFormLayout()
        
        self.default_save_path = QLineEdit()
        self.default_save_path.setText(self.preferences.get("default_save_path", ""))
        browse_button = QPushButton("Browse...")
        browse_button.clicked.connect(self.browse_save_path)
        
        path_layout = QHBoxLayout()
        path_layout.addWidget(self.default_save_path)
        path_layout.addWidget(browse_button)
        
        file_location_layout.addRow("Default save location:", path_layout)
        file_location_group.setLayout(file_location_layout)
        layout.addWidget(file_location_group)
        
        # Calculation options
        calc_group = QGroupBox("Calculation")
        calc_layout = QVBoxLayout()
        
        self.auto_calculate = QCheckBox("Automatically calculate formulas")
        self.auto_calculate.setChecked(self.preferences.get("auto_calculate", True))
        calc_layout.addWidget(self.auto_calculate)
        
        calc_group.setLayout(calc_layout)
        layout.addWidget(calc_group)
        
        layout.addStretch(1)  # Add stretch to push widgets to the top
        
        self.tab_widget.addTab(general_tab, "General")

    def create_appearance_tab(self):
        """Create the appearance preferences tab"""
        appearance_tab = QWidget()
        layout = QVBoxLayout(appearance_tab)
        
        # Theme selection
        theme_group = QGroupBox("Theme")
        theme_layout = QHBoxLayout()
        
        theme_layout.addWidget(QLabel("Application theme:"))
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Light", "Dark"])
        current_theme = self.preferences.get("theme", "light")
        self.theme_combo.setCurrentText(current_theme.capitalize())
        theme_layout.addWidget(self.theme_combo)
        
        theme_group.setLayout(theme_layout)
        layout.addWidget(theme_group)
        
        # Grid options
        grid_group = QGroupBox("Grid")
        grid_layout = QVBoxLayout()
        
        self.show_grid = QCheckBox("Show gridlines")
        self.show_grid.setChecked(self.preferences.get("show_gridlines", True))
        grid_layout.addWidget(self.show_grid)
        
        grid_color_layout = QHBoxLayout()
        grid_color_layout.addWidget(QLabel("Grid color:"))
        self.grid_color_button = QPushButton()
        self.grid_color = QColor(self.preferences.get("grid_color", "#d0d0d0"))
        self.update_color_button(self.grid_color_button, self.grid_color)
        self.grid_color_button.clicked.connect(self.select_grid_color)
        grid_color_layout.addWidget(self.grid_color_button)
        
        grid_layout.addLayout(grid_color_layout)
        grid_group.setLayout(grid_layout)
        layout.addWidget(grid_group)
        
        # Font options
        font_group = QGroupBox("Default Font")
        font_layout = QVBoxLayout()
        
        font_family_layout = QHBoxLayout()
        font_family_layout.addWidget(QLabel("Font family:"))
        self.font_family_combo = QComboBox()
        common_fonts = ["Arial", "Times New Roman", "Courier New", "Verdana", "Helvetica", "Calibri"]
        self.font_family_combo.addItems(common_fonts)
        current_font = self.preferences.get("default_font_family", DEFAULT_FONT_FAMILY)
        if current_font in common_fonts:
            self.font_family_combo.setCurrentText(current_font)
        font_family_layout.addWidget(self.font_family_combo)
        
        font_layout.addLayout(font_family_layout)
        
        font_size_layout = QHBoxLayout()
        font_size_layout.addWidget(QLabel("Font size:"))
        self.font_size_spin = QSpinBox()
        self.font_size_spin.setRange(6, 72)
        self.font_size_spin.setValue(self.preferences.get("default_font_size", DEFAULT_FONT_SIZE))
        font_size_layout.addWidget(self.font_size_spin)
        
        font_layout.addLayout(font_size_layout)
        font_group.setLayout(font_layout)
        layout.addWidget(font_group)
        
        layout.addStretch(1)
        
        self.tab_widget.addTab(appearance_tab, "Appearance")

    def create_editor_tab(self):
        """Create the editor preferences tab"""
        editor_tab = QWidget()
        layout = QVBoxLayout(editor_tab)
        
        # Editing options
        editing_group = QGroupBox("Editing Options")
        editing_layout = QVBoxLayout()
        
        self.auto_complete = QCheckBox("Enable auto-complete")
        self.auto_complete.setChecked(self.preferences.get("auto_complete", True))
        editing_layout.addWidget(self.auto_complete)
        
        self.auto_format = QCheckBox("Auto-format formulas")
        self.auto_format.setChecked(self.preferences.get("auto_format", True))
        editing_layout.addWidget(self.auto_format)
        
        editing_group.setLayout(editing_layout)
        layout.addWidget(editing_group)
        
        # Default cell options
        cell_group = QGroupBox("Default Cell Options")
        cell_layout = QVBoxLayout()
        
        # Default cell color
        cell_color_layout = QHBoxLayout()
        cell_color_layout.addWidget(QLabel("Default cell color:"))
        self.cell_color_button = QPushButton()
        self.cell_color = QColor(self.preferences.get("default_cell_color", "#FFFFFF"))
        self.update_color_button(self.cell_color_button, self.cell_color)
        self.cell_color_button.clicked.connect(self.select_cell_color)
        cell_color_layout.addWidget(self.cell_color_button)
        
        cell_layout.addLayout(cell_color_layout)
        
        # Default text color
        text_color_layout = QHBoxLayout()
        text_color_layout.addWidget(QLabel("Default text color:"))
        self.text_color_button = QPushButton()
        self.text_color = QColor(self.preferences.get("default_text_color", "#000000"))
        self.update_color_button(self.text_color_button, self.text_color)
        self.text_color_button.clicked.connect(self.select_text_color)
        text_color_layout.addWidget(self.text_color_button)
        
        cell_layout.addLayout(text_color_layout)
        
        cell_group.setLayout(cell_layout)
        layout.addWidget(cell_group)
        
        layout.addStretch(1)
        
        self.tab_widget.addTab(editor_tab, "Editor")

    def update_color_button(self, button, color):
        """Update a color button's appearance to show the selected color"""
        button.setStyleSheet(f"background-color: {color.name()}; min-width: 60px;")

    def select_grid_color(self):
        """Open color dialog to select grid color"""
        color = QColorDialog.getColor(self.grid_color, self)
        if color.isValid():
            self.grid_color = color
            self.update_color_button(self.grid_color_button, color)

    def select_cell_color(self):
        """Open color dialog to select default cell color"""
        color = QColorDialog.getColor(self.cell_color, self)
        if color.isValid():
            self.cell_color = color
            self.update_color_button(self.cell_color_button, color)

    def select_text_color(self):
        """Open color dialog to select default text color"""
        color = QColorDialog.getColor(self.text_color, self)
        if color.isValid():
            self.text_color = color
            self.update_color_button(self.text_color_button, color)

    def browse_save_path(self):
        """Open directory browser to select default save path"""
        directory = QFileDialog.getExistingDirectory(
            self, "Select Default Save Directory", 
            self.default_save_path.text()
        )
        if directory:
            self.default_save_path.setText(directory)

    def accept(self):
        """Save preferences when dialog is accepted"""
        # Update the preferences dict with the selected values
        self.preferences["theme"] = self.theme_combo.currentText().lower()
        self.preferences["auto_save_enabled"] = self.auto_save_enabled.isChecked()
        self.preferences["auto_save_interval"] = self.auto_save_interval.value()
        self.preferences["default_save_path"] = self.default_save_path.text()
        self.preferences["auto_calculate"] = self.auto_calculate.isChecked()
        self.preferences["show_gridlines"] = self.show_grid.isChecked()
        self.preferences["grid_color"] = self.grid_color.name()
        self.preferences["default_font_family"] = self.font_family_combo.currentText()
        self.preferences["default_font_size"] = self.font_size_spin.value()
        self.preferences["auto_complete"] = self.auto_complete.isChecked()
        self.preferences["auto_format"] = self.auto_format.isChecked()
        self.preferences["default_cell_color"] = self.cell_color.name()
        self.preferences["default_text_color"] = self.text_color.name()
        
        # Save to settings
        self.save_settings()
        
        super().accept()
