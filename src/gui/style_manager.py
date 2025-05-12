from PyQt5.QtWidgets import QApplication, QStyleFactory
from PyQt5.QtGui import QPalette, QColor
from PyQt5.QtCore import Qt

def apply_stylesheet(window):
    """Apply a modern stylesheet to the application"""
    # First set Fusion style for a clean modern base
    QApplication.setStyle(QStyleFactory.create("Fusion"))
    
    # Create a palette for the application
    palette = QPalette()
    
    # Set colors for the light theme
    palette.setColor(QPalette.Window, QColor(240, 240, 240))
    palette.setColor(QPalette.WindowText, QColor(0, 0, 0))
    palette.setColor(QPalette.Base, QColor(255, 255, 255))
    palette.setColor(QPalette.AlternateBase, QColor(245, 245, 245))
    palette.setColor(QPalette.ToolTipBase, QColor(255, 255, 255))
    palette.setColor(QPalette.ToolTipText, QColor(0, 0, 0))
    palette.setColor(QPalette.Text, QColor(0, 0, 0))
    palette.setColor(QPalette.Button, QColor(240, 240, 240))
    palette.setColor(QPalette.ButtonText, QColor(0, 0, 0))
    palette.setColor(QPalette.BrightText, Qt.red)
    palette.setColor(QPalette.Link, QColor(42, 130, 218))
    palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
    palette.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
    
    # Set the modified palette
    QApplication.setPalette(palette)
    
    # Apply additional stylesheet for fine-tuning
    stylesheet = """
    QMainWindow {
        border: none;
    }
    
    QMenuBar {
        background-color: #f0f0f0;
        border-bottom: 1px solid #d0d0d0;
    }
    
    QMenuBar::item {
        padding: 5px 10px;
        background: transparent;
    }
    
    QMenuBar::item:selected {
        background: #e0e0e0;
    }
    
    QMenu {
        background-color: #ffffff;
        border: 1px solid #d0d0d0;
    }
    
    QMenu::item {
        padding: 5px 30px 5px 30px;
    }
    
    QMenu::item:selected {
        background-color: #e6f2ff;
    }
    
    QToolBar {
        background-color: #f8f8f8;
        border-bottom: 1px solid #d0d0d0;
        spacing: 3px;
    }
    
    QToolButton {
        background-color: transparent;
        border: 1px solid transparent;
        border-radius: 2px;
        padding: 3px;
    }
    
    QToolButton:hover {
        background-color: #e6f2ff;
        border: 1px solid #99ccff;
    }
    
    QToolButton:pressed {
        background-color: #cce6ff;
    }
    
    QTabWidget::pane {
        border: 1px solid #d0d0d0;
        border-radius: 3px;
    }
    
    QTabBar::tab {
        background-color: #e8e8e8;
        border: 1px solid #d0d0d0;
        border-bottom: none;
        border-top-left-radius: 3px;
        border-top-right-radius: 3px;
        padding: 5px 10px;
        margin-right: 2px;
    }
    
    QTabBar::tab:selected {
        background-color: #ffffff;
    }
    
    QTabBar::tab:hover:!selected {
        background-color: #f0f0f0;
    }
    
    QTableView {
        gridline-color: #d0d0d0;
        selection-background-color: #e6f2ff;
        selection-color: #000000;
    }
    
    QHeaderView::section {
        background-color: #f0f0f0;
        border: 1px solid #d0d0d0;
        padding: 4px;
    }
    
    QLineEdit {
        border: 1px solid #d0d0d0;
        border-radius: 2px;
        padding: 3px;
    }
    
    QPushButton {
        background-color: #f0f0f0;
        border: 1px solid #d0d0d0;
        border-radius: 2px;
        padding: 5px 15px;
    }
    
    QPushButton:hover {
        background-color: #e6f2ff;
        border: 1px solid #99ccff;
    }
    
    QPushButton:pressed {
        background-color: #cce6ff;
    }
    
    QStatusBar {
        background-color: #f8f8f8;
        border-top: 1px solid #d0d0d0;
    }
    
    QComboBox {
        border: 1px solid #d0d0d0;
        border-radius: 2px;
        padding: 3px;
        background-color: white;
    }
    
    QComboBox::drop-down {
        subcontrol-origin: padding;
        subcontrol-position: top right;
        width: 15px;
        border-left: 1px solid #d0d0d0;
    }
    
    QComboBox::down-arrow {
        image: url(:/arrow_down.png);
    }
    
    QSpinBox {
        border: 1px solid #d0d0d0;
        border-radius: 2px;
        padding: 3px;
    }
    
    QCheckBox {
        spacing: 5px;
    }
    
    QGroupBox {
        font-weight: bold;
        border: 1px solid #d0d0d0;
        border-radius: 3px;
        margin-top: 10px;
        padding-top: 10px;
    }
    
    QGroupBox::title {
        subcontrol-origin: margin;
        subcontrol-position: top left;
        left: 10px;
        padding: 0 3px;
    }
    """
    
    window.setStyleSheet(stylesheet)

def create_dark_theme(window):
    """Create a dark theme for the application"""
    # Set Fusion style as the base
    QApplication.setStyle(QStyleFactory.create("Fusion"))
    
    # Create dark palette
    palette = QPalette()
    dark_color = QColor(53, 53, 53)
    palette.setColor(QPalette.Window, dark_color)
    palette.setColor(QPalette.WindowText, Qt.white)
    palette.setColor(QPalette.Base, QColor(25, 25, 25))
    palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
    palette.setColor(QPalette.ToolTipBase, Qt.white)
    palette.setColor(QPalette.ToolTipText, Qt.white)
    palette.setColor(QPalette.Text, Qt.white)
    palette.setColor(QPalette.Button, dark_color)
    palette.setColor(QPalette.ButtonText, Qt.white)
    palette.setColor(QPalette.BrightText, Qt.red)
    palette.setColor(QPalette.Link, QColor(42, 130, 218))
    palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
    palette.setColor(QPalette.HighlightedText, Qt.black)
    
    # Apply the dark palette
    QApplication.setPalette(palette)
    
    # Apply additional stylesheet for dark theme
    dark_stylesheet = """
    QMainWindow {
        border: none;
    }
    
    QMenuBar {
        background-color: #353535;
        border-bottom: 1px solid #1e1e1e;
    }
    
    QMenuBar::item {
        padding: 5px 10px;
        background: transparent;
    }
    
    QMenuBar::item:selected {
        background: #454545;
    }
    
    QMenu {
        background-color: #353535;
        border: 1px solid #1e1e1e;
    }
    
    QMenu::item {
        padding: 5px 30px 5px 30px;
        color: white;
    }
    
    QMenu::item:selected {
        background-color: #454545;
    }
    
    QToolBar {
        background-color: #353535;
        border-bottom: 1px solid #1e1e1e;
        spacing: 3px;
    }
    
    QToolButton {
        background-color: transparent;
        border: 1px solid transparent;
        border-radius: 2px;
        padding: 3px;
        color: white;
    }
    
    QToolButton:hover {
        background-color: #454545;
        border: 1px solid #555555;
    }
    
    QToolButton:pressed {
        background-color: #252525;
    }
    
    QTabWidget::pane {
        border: 1px solid #1e1e1e;
        border-radius: 3px;
    }
    
    QTabBar::tab {
        background-color: #353535;
        border: 1px solid #1e1e1e;
        border-bottom: none;
        border-top-left-radius: 3px;
        border-top-right-radius: 3px;
        padding: 5px 10px;
        margin-right: 2px;
        color: white;
    }
    
    QTabBar::tab:selected {
        background-color: #454545;
    }
    
    QTabBar::tab:hover:!selected {
        background-color: #404040;
    }
    
    QTableView {
        gridline-color: #1e1e1e;
        selection-background-color: #454545;
        selection-color: white;
        background-color: #252525;
        color: white;
    }
    
    QHeaderView::section {
        background-color: #353535;
        border: 1px solid #1e1e1e;
        padding: 4px;
        color: white;
    }
    
    QLineEdit {
        border: 1px solid #1e1e1e;
        border-radius: 2px;
        padding: 3px;
        background-color: #252525;
        color: white;
    }
    
    QPushButton {
        background-color: #353535;
        border: 1px solid #1e1e1e;
        border-radius: 2px;
        padding: 5px 15px;
        color: white;
    }
    
    QPushButton:hover {
        background-color: #454545;
        border: 1px solid #555555;
    }
    
    QPushButton:pressed {
        background-color: #252525;
    }
    
    QStatusBar {
        background-color: #353535;
        border-top: 1px solid #1e1e1e;
        color: white;
    }
    
    QComboBox {
        border: 1px solid #1e1e1e;
        border-radius: 2px;
        padding: 3px;
        background-color: #252525;
        color: white;
    }
    
    QComboBox::drop-down {
        subcontrol-origin: padding;
        subcontrol-position: top right;
        width: 15px;
        border-left: 1px solid #1e1e1e;
    }
    
    QSpinBox {
        border: 1px solid #1e1e1e;
        border-radius: 2px;
        padding: 3px;
        background-color: #252525;
        color: white;
    }
    
    QCheckBox {
        spacing: 5px;
        color: white;
    }
    
    QGroupBox {
        font-weight: bold;
        border: 1px solid #1e1e1e;
        border-radius: 3px;
        margin-top: 10px;
        padding-top: 10px;
        color: white;
    }
    
    QGroupBox::title {
        subcontrol-origin: margin;
        subcontrol-position: top left;
        left: 10px;
        padding: 0 3px;
    }
    """
    
    window.setStyleSheet(dark_stylesheet)

def toggle_theme(window, theme_name="light"):
    """Toggle between light and dark themes"""
    if theme_name.lower() == "dark":
        create_dark_theme(window)
    else:
        apply_stylesheet(window)  # Default to light theme
