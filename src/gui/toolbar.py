from PyQt5.QtWidgets import QToolBar, QAction, QComboBox, QSpinBox, QLabel, QWidget, QToolButton, QMenu
from PyQt5.QtGui import QIcon, QFont, QPixmap
from PyQt5.QtCore import QSize, Qt
import os

class Toolbar(QToolBar):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Toolbar")
        self.setObjectName("MainToolBar")
        self.setIconSize(QSize(22, 22))
        self.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        self.init_toolbar()

    def addSection(self, title):
        """Add a section dropdown button to the toolbar"""
        # Add a small spacer before the section if not the first section
        if self.actions():
            self.addSeparator()
            
        # Create a dropdown button for the section
        section_button = QToolButton(self)
        section_button.setText(title)
        section_button.setPopupMode(QToolButton.InstantPopup)
        section_button.setToolButtonStyle(Qt.ToolButtonTextOnly)
        section_button.setStyleSheet("QToolButton { font-weight: bold; color: #666; margin: 2px; padding: 3px; }")
        
        # Create a menu for the dropdown
        section_menu = QMenu(section_button)
        section_button.setMenu(section_menu)
        
        # Add the button to toolbar
        self.addWidget(section_button)
        return section_button, section_menu

    def init_toolbar(self):
        # Ensure resource path exists
        icon_path = os.path.join(os.path.dirname(__file__), '..', '..', 'resources', 'icons')
        
        # File operations
        file_button, file_menu = self.addSection("File")
        
        # Create file actions and add to menu
        self.new_action = self.create_action("New", "new.png", "Create a new spreadsheet (Ctrl+N)", icon_path)
        file_menu.addAction(self.new_action)
        
        self.open_action = self.create_action("Open", "open.png", "Open an existing spreadsheet (Ctrl+O)", icon_path)
        file_menu.addAction(self.open_action)
        
        self.save_action = self.create_action("Save", "save.png", "Save the current spreadsheet (Ctrl+S)", icon_path)
        file_menu.addAction(self.save_action)
        
        # Also add the main actions directly to the toolbar for quick access
        self.addAction(self.new_action)
        self.addAction(self.open_action)
        self.addAction(self.save_action)
        
        self.addSeparator()
        
        # Edit operations
        edit_button, edit_menu = self.addSection("Edit")
        
        self.undo_action = self.create_action("Undo", "undo.png", "Undo the last action (Ctrl+Z)", icon_path)
        edit_menu.addAction(self.undo_action)
        
        self.redo_action = self.create_action("Redo", "redo.png", "Redo the last undone action (Ctrl+Y)", icon_path)
        edit_menu.addAction(self.redo_action)
        
        # Add to toolbar for quick access
        self.addAction(self.undo_action)
        self.addAction(self.redo_action)
        
        self.addSeparator()
        
        # Cut/Copy/Paste
        clipboard_button, clipboard_menu = self.addSection("Clipboard")
        
        self.cut_action = self.create_action("Cut", "cut.png", "Cut the selected cells to clipboard (Ctrl+X)", icon_path)
        clipboard_menu.addAction(self.cut_action)
        
        self.copy_action = self.create_action("Copy", "copy.png", "Copy the selected cells to clipboard (Ctrl+C)", icon_path)
        clipboard_menu.addAction(self.copy_action)
        
        self.paste_action = self.create_action("Paste", "paste.png", "Paste from clipboard (Ctrl+V)", icon_path)
        clipboard_menu.addAction(self.paste_action)
        
        # Add to toolbar for quick access
        self.addAction(self.cut_action)
        self.addAction(self.copy_action)
        self.addAction(self.paste_action)
        
        self.addSeparator()
        
        # Format operations
        format_button, format_menu = self.addSection("Format")
        
        # Font family
        font_label = QLabel("Font:")
        font_label.setStyleSheet("margin-left: 5px; margin-right: 2px;")
        self.addWidget(font_label)
        
        self.font_family = QComboBox(self)
        self.font_family.addItems(["Arial", "Times New Roman", "Courier New", "Verdana", "Helvetica", "Calibri", "Cambria"])
        self.font_family.setCurrentText("Arial")
        self.font_family.setMinimumWidth(120)
        self.font_family.setToolTip("Change font family")
        self.addWidget(self.font_family)
        
        # Font size
        size_label = QLabel("Size:")
        size_label.setStyleSheet("margin-left: 5px; margin-right: 2px;")
        self.addWidget(size_label)
        
        self.font_size = QSpinBox(self)
        self.font_size.setRange(8, 72)
        self.font_size.setValue(11)
        self.font_size.setMinimumWidth(50)
        self.font_size.setToolTip("Change font size")
        self.addWidget(self.font_size)
        
        self.addSeparator()
        
        # Text formatting buttons
        self.bold_action = self.create_action("Bold", "bold.png", "Bold text formatting (Ctrl+B)", icon_path)
        self.bold_action.setCheckable(True)
        format_menu.addAction(self.bold_action)
        
        self.italic_action = self.create_action("Italic", "italic.png", "Italic text formatting (Ctrl+I)", icon_path)
        self.italic_action.setCheckable(True)
        format_menu.addAction(self.italic_action)
        
        self.underline_action = self.create_action("Underline", "underline.png", "Underline text formatting (Ctrl+U)", icon_path)
        self.underline_action.setCheckable(True)
        format_menu.addAction(self.underline_action)
        
        # Add to toolbar for quick access
        self.addAction(self.bold_action)
        self.addAction(self.italic_action)
        self.addAction(self.underline_action)
        
        self.addSeparator()
        
        # Color formatting buttons
        self.text_color_action = self.create_action("Text Color", "text_color.png", "Change text color", icon_path)
        format_menu.addAction(self.text_color_action)
        
        self.fill_color_action = self.create_action("Fill Color", "fill_color.png", "Change cell background color", icon_path)
        format_menu.addAction(self.fill_color_action)
        
        # Add to toolbar for quick access
        self.addAction(self.text_color_action)
        self.addAction(self.fill_color_action)
        
        self.addSeparator()
        
        # Alignment
        alignment_button, alignment_menu = self.addSection("Alignment")
        
        self.align_left_action = self.create_action("Left", "align_left.png", "Align text to left", icon_path)
        alignment_menu.addAction(self.align_left_action)
        
        self.align_center_action = self.create_action("Center", "align_center.png", "Align text to center", icon_path)
        alignment_menu.addAction(self.align_center_action)
        
        self.align_right_action = self.create_action("Right", "align_right.png", "Align text to right", icon_path)
        alignment_menu.addAction(self.align_right_action)
        
        # Add to toolbar for quick access
        self.addAction(self.align_left_action)
        self.addAction(self.align_center_action)
        self.addAction(self.align_right_action)
        
        self.addSeparator()
        
        # Number formatting
        number_format_button, number_format_menu = self.addSection("Number Format")
        
        # Create a dropdown button for number formats
        num_format_button = QToolButton(self)
        num_format_button.setText("Format")
        num_format_button.setToolTip("Apply number formatting")
        num_format_button.setPopupMode(QToolButton.InstantPopup)
        
        num_format_menu = QMenu(num_format_button)
        
        currency_action = self.create_action("Currency", "currency.png", "Format as currency", icon_path)
        percentage_action = self.create_action("Percentage", "percentage.png", "Format as percentage", icon_path)
        comma_action = self.create_action("Comma", "comma.png", "Format with comma separators", icon_path)
        decimal_action = self.create_action("Decimals", "decimal.png", "Format with decimal places", icon_path)
        
        num_format_menu.addAction(currency_action)
        num_format_menu.addAction(percentage_action)
        num_format_menu.addAction(comma_action)
        num_format_menu.addAction(decimal_action)
        
        num_format_button.setMenu(num_format_menu)
        self.addWidget(num_format_button)
        
        # Connect number format actions
        currency_action.triggered.connect(lambda: self.apply_number_format("currency"))
        percentage_action.triggered.connect(lambda: self.apply_number_format("percentage"))
        comma_action.triggered.connect(lambda: self.apply_number_format("comma"))
        decimal_action.triggered.connect(lambda: self.apply_number_format("decimal_2"))
        
        self.addSeparator()
        
        # Insert menu button
        insert_button = QToolButton(self)
        insert_button.setText("Insert")
        insert_button.setToolTip("Insert elements")
        insert_button.setPopupMode(QToolButton.InstantPopup)
        
        insert_menu = QMenu(insert_button)
        
        chart_action = self.create_action("Chart", "chart.png", "Insert a chart", icon_path)
        formula_action = self.create_action("Formula", "formula.png", "Insert a formula", icon_path)
        comment_action = self.create_action("Comment", "comment.png", "Insert a comment", icon_path)
        
        insert_menu.addAction(chart_action)
        insert_menu.addAction(formula_action)
        insert_menu.addAction(comment_action)
        
        insert_button.setMenu(insert_menu)
        self.addWidget(insert_button)
        
        # Help button
        self.help_action = self.create_action("Help", "help.png", "Get help on using PySpreadsheet", icon_path)
        self.addAction(self.help_action)
        
        # Connect actions to format application
        self.connect_formatting_actions()

    def create_action(self, text, icon_name, tooltip, icon_path):
        """Helper to create action with icon"""
        action = QAction(text, self)
        
        # Try to load the icon
        try:
            icon_file = os.path.join(icon_path, icon_name)
            if os.path.exists(icon_file):
                action.setIcon(QIcon(icon_file))
            else:
                # If icon doesn't exist, create a simple colored icon
                pixmap = QPixmap(22, 22)
                pixmap.fill(Qt.transparent)
                action.setIcon(QIcon(pixmap))
        except:
            # Handle any issues with icon loading
            pass
            
        action.setToolTip(tooltip)
        return action

    def connect_formatting_actions(self):
        # Font family and size changes
        self.font_family.currentTextChanged.connect(self.apply_font_family)
        self.font_size.valueChanged.connect(self.apply_font_size)
        
        # Bold, italic, underline
        self.bold_action.triggered.connect(self.apply_bold)
        self.italic_action.triggered.connect(self.apply_italic)
        self.underline_action.triggered.connect(self.apply_underline)
        
        # Alignment
        self.align_left_action.triggered.connect(lambda: self.apply_alignment(Qt.AlignLeft))
        self.align_center_action.triggered.connect(lambda: self.apply_alignment(Qt.AlignCenter))
        self.align_right_action.triggered.connect(lambda: self.apply_alignment(Qt.AlignRight))
        
        # Colors
        self.text_color_action.triggered.connect(self.apply_text_color)
        self.fill_color_action.triggered.connect(self.apply_fill_color)

    def apply_font_family(self, family):
        """Apply the selected font family to the selected cells"""
        main_window = self.parent()
        sheet_view = main_window.current_sheet_view()
        if sheet_view:
            selected_items = sheet_view.table.selectedItems()
            for item in selected_items:
                font = item.font()
                font.setFamily(family)
                item.setFont(font)

    def apply_font_size(self, size):
        """Apply the selected font size to the selected cells"""
        main_window = self.parent()
        sheet_view = main_window.current_sheet_view()
        if sheet_view:
            selected_items = sheet_view.table.selectedItems()
            for item in selected_items:
                font = item.font()
                font.setPointSize(size)
                item.setFont(font)

    def apply_bold(self, checked):
        """Apply bold formatting to the selected cells"""
        main_window = self.parent()
        sheet_view = main_window.current_sheet_view()
        if sheet_view:
            selected_items = sheet_view.table.selectedItems()
            for item in selected_items:
                font = item.font()
                font.setBold(checked)
                item.setFont(font)

    def apply_italic(self, checked):
        """Apply italic formatting to the selected cells"""
        main_window = self.parent()
        sheet_view = main_window.current_sheet_view()
        if sheet_view:
            selected_items = sheet_view.table.selectedItems()
            for item in selected_items:
                font = item.font()
                font.setItalic(checked)
                item.setFont(font)

    def apply_underline(self, checked):
        """Apply underline formatting to the selected cells"""
        main_window = self.parent()
        sheet_view = main_window.current_sheet_view()
        if sheet_view:
            selected_items = sheet_view.table.selectedItems()
            for item in selected_items:
                font = item.font()
                font.setUnderline(checked)
                item.setFont(font)

    def apply_alignment(self, alignment):
        """Apply the specified alignment to the selected cells"""
        main_window = self.parent()
        sheet_view = main_window.current_sheet_view()
        if sheet_view:
            selected_items = sheet_view.table.selectedItems()
            for item in selected_items:
                item.setTextAlignment(alignment)

    def apply_text_color(self):
        """Open color dialog and apply text color"""
        main_window = self.parent()
        main_window.change_text_color()

    def apply_fill_color(self):
        """Open color dialog and apply cell color"""
        main_window = self.parent()
        main_window.change_cell_color()

    def apply_number_format(self, format_type):
        """Apply number formatting to selected cells"""
        main_window = self.parent()
        main_window.apply_number_format(format_type)