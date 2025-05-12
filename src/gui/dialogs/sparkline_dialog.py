from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, 
    QPushButton, QComboBox, QDialogButtonBox, QGroupBox,
    QRadioButton, QColorDialog, QSpinBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor

class SparklineDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert Sparkline")
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout(self)
        
        # Data range
        data_group = QGroupBox("Data Range")
        data_layout = QVBoxLayout()
        
        range_layout = QHBoxLayout()
        range_label = QLabel("Data Range:")
        self.range_edit = QLineEdit()
        range_layout.addWidget(range_label)
        range_layout.addWidget(self.range_edit)
        data_layout.addLayout(range_layout)
        
        location_layout = QHBoxLayout()
        location_label = QLabel("Location:")
        self.location_edit = QLineEdit()
        location_layout.addWidget(location_label)
        location_layout.addWidget(self.location_edit)
        data_layout.addLayout(location_layout)
        
        data_group.setLayout(data_layout)
        layout.addWidget(data_group)
        
        # Sparkline type
        type_group = QGroupBox("Type")
        type_layout = QHBoxLayout()
        
        self.line_radio = QRadioButton("Line")
        self.line_radio.setChecked(True)
        type_layout.addWidget(self.line_radio)
        
        self.column_radio = QRadioButton("Column")
        type_layout.addWidget(self.column_radio)
        
        self.win_loss_radio = QRadioButton("Win/Loss")
        type_layout.addWidget(self.win_loss_radio)
        
        type_group.setLayout(type_layout)
        layout.addWidget(type_group)
        
        # Style options
        style_group = QGroupBox("Style")
        style_layout = QVBoxLayout()
        
        # Line color
        line_color_layout = QHBoxLayout()
        line_color_label = QLabel("Line Color:")
        self.line_color_button = QPushButton()
        self.line_color_button.setFixedSize(30, 20)
        self.line_color = QColor(0, 0, 255)  # Default blue
        self.update_line_color_button()
        self.line_color_button.clicked.connect(self.select_line_color)
        line_color_layout.addWidget(line_color_label)
        line_color_layout.addWidget(self.line_color_button)
        style_layout.addLayout(line_color_layout)
        
        # Markers
        markers_layout = QHBoxLayout()
        markers_label = QLabel("Show Markers:")
        self.markers_combo = QComboBox()
        self.markers_combo.addItems(["None", "All Points", "High Point", "Low Point", "First Point", "Last Point"])
        markers_layout.addWidget(markers_label)
        markers_layout.addWidget(self.markers_combo)
        style_layout.addLayout(markers_layout)
        
        # Marker color
        marker_color_layout = QHBoxLayout()
        marker_color_label = QLabel("Marker Color:")
        self.marker_color_button = QPushButton()
        self.marker_color_button.setFixedSize(30, 20)
        self.marker_color = QColor(255, 0, 0)  # Default red
        self.update_marker_color_button()
        self.marker_color_button.clicked.connect(self.select_marker_color)
        marker_color_layout.addWidget(marker_color_label)
        marker_color_layout.addWidget(self.marker_color_button)
        style_layout.addLayout(marker_color_layout)
        
        # Line weight
        weight_layout = QHBoxLayout()
        weight_label = QLabel("Line Weight:")
        self.weight_spin = QSpinBox()
        self.weight_spin.setRange(1, 10)
        self.weight_spin.setValue(2)
        weight_layout.addWidget(weight_label)
        weight_layout.addWidget(self.weight_spin)
        style_layout.addLayout(weight_layout)
        
        style_group.setLayout(style_layout)
        layout.addWidget(style_group)
        
        # Axis options
        axis_group = QGroupBox("Axis")
        axis_layout = QVBoxLayout()
        
        show_axis_layout = QHBoxLayout()
        show_axis_label = QLabel("Show Axis:")
        self.show_axis_combo = QComboBox()
        self.show_axis_combo.addItems(["None", "Horizontal", "Vertical", "Both"])
        show_axis_layout.addWidget(show_axis_label)
        show_axis_layout.addWidget(self.show_axis_combo)
        axis_layout.addLayout(show_axis_layout)
        
        axis_group.setLayout(axis_layout)
        layout.addWidget(axis_group)
        
        # Dialog buttons
        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
    
    def update_line_color_button(self):
        self.line_color_button.setStyleSheet(f"background-color: {self.line_color.name()};")
    
    def update_marker_color_button(self):
        self.marker_color_button.setStyleSheet(f"background-color: {self.marker_color.name()};")
    
    def select_line_color(self):
        color = QColorDialog.getColor(self.line_color, self)
        if color.isValid():
            self.line_color = color
            self.update_line_color_button()
    
    def select_marker_color(self):
        color = QColorDialog.getColor(self.marker_color, self)
        if color.isValid():
            self.marker_color = color
            self.update_marker_color_button()
    
    def get_settings(self):
        """Get all sparkline settings from the dialog"""
        sparkline_type = "line"
        if self.column_radio.isChecked():
            sparkline_type = "column"
        elif self.win_loss_radio.isChecked():
            sparkline_type = "win_loss"
        
        return {
            'data_range': self.range_edit.text(),
            'location': self.location_edit.text(),
            'type': sparkline_type,
            'line_color': self.line_color.name(),
            'show_markers': self.markers_combo.currentText(),
            'marker_color': self.marker_color.name(),
            'line_weight': self.weight_spin.value(),
            'show_axis': self.show_axis_combo.currentText()
        }
