import sys
import datetime  # Added missing import for datetime
from PyQt5.QtWidgets import (
    QMainWindow, QAction, QFileDialog, QApplication, QVBoxLayout, QHBoxLayout, 
    QWidget, QLabel, QStatusBar, QTabWidget, QColorDialog, QFontDialog, QMessageBox,
    QDialog, QInputDialog, QMenu, QSplitter, QGridLayout, QLineEdit, QPushButton,
    QComboBox, QCheckBox, QDialogButtonBox, QListWidget, QGroupBox, QRadioButton, QTableWidgetItem
)
from PyQt5.QtCore import Qt, QSize, QSettings
from PyQt5.QtGui import QFont, QIcon, QPalette, QColor
from src.gui.sheet_view import SheetView, FindDialog, ReplaceDialog
from src.gui.toolbar import Toolbar
from src.gui.widgets import CustomLineEdit, FormulaLineEdit
from src.gui.style_manager import apply_stylesheet
from src.data_io.file_manager import FileManager
from src.data_io.excel_handler import read_excel, write_excel
from src.data_io.csv_handler import read_csv, write_csv
from src.core.workbook import Workbook
from src.engine.calculator import Calculator
from src.engine.chart import ChartDialog
from src.utils.config import USER_PREFERENCES
from src.gui.dialogs.preferences_dialog import PreferencesDialog

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PySpreadsheet")
        self.setGeometry(100, 100, 1200, 800)
        
        # Initialize components
        self.workbook = Workbook()
        self.workbook.add_sheet("Sheet1")
        self.file_manager = FileManager()
        self.calculator = Calculator()
        self.current_sheet_name = "Sheet1"
        self.settings = QSettings("AryanTech", "PySpreadsheet")
        self.named_ranges = {}  # Initialize named ranges dictionary
        
        # Set application style
        apply_stylesheet(self)
        
        # Create central widget and layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(4, 4, 4, 4)
        self.main_layout.setSpacing(2)
        
        # Create menu bar
        self.create_menu()
        
        # Create toolbar
        self.toolbar = Toolbar(self)
        self.toolbar.setObjectName("MainToolBar")
        self.addToolBar(self.toolbar)
        self.connect_toolbar_actions()
        
        # Create formula bar first
        self.create_formula_bar()
        
        # Then create sheet tabs
        self.create_sheet_tabs()
        
        # Connect formula input to apply formula
        self.formula_input.returnPressed.connect(self.apply_formula)
        
        # Set calculator for formula autocomplete
        self.formula_input.set_calculator(self.calculator)
        
        # Create status bar
        self.statusBar().showMessage("Ready")
        
        # Load settings
        self.load_settings()

    def create_menu(self):
        menu_bar = self.menuBar()
        
        # File menu
        file_menu = menu_bar.addMenu("&File")
        
        new_action = QAction("&New", self)
        new_action.setShortcut("Ctrl+N")
        new_action.triggered.connect(self.new_file)
        file_menu.addAction(new_action)
        
        open_action = QAction("&Open...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)
        
        save_action = QAction("&Save", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_file)
        file_menu.addAction(save_action)
        
        save_as_action = QAction("Save &As...", self)
        save_as_action.setShortcut("Ctrl+Shift+S")
        save_as_action.triggered.connect(self.save_file_as)
        file_menu.addAction(save_as_action)
        
        file_menu.addSeparator()
        
        # Version history submenu
        version_menu = QMenu("Version &History", self)

        save_version_action = QAction("Save Current &Version", self)
        save_version_action.triggered.connect(self.save_document_version)
        version_menu.addAction(save_version_action)

        view_versions_action = QAction("&View Version History", self)
        view_versions_action.triggered.connect(self.view_version_history)
        version_menu.addAction(view_versions_action)

        file_menu.addMenu(version_menu)
        
        export_pdf_action = QAction("Export to &PDF...", self)
        export_pdf_action.triggered.connect(self.export_to_pdf)
        file_menu.addAction(export_pdf_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("E&xit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Edit menu
        edit_menu = menu_bar.addMenu("&Edit")
        
        undo_action = QAction("&Undo", self)
        undo_action.setShortcut("Ctrl+Z")
        undo_action.triggered.connect(self.undo)
        edit_menu.addAction(undo_action)
        
        redo_action = QAction("&Redo", self)
        redo_action.setShortcut("Ctrl+Y")
        redo_action.triggered.connect(self.redo)
        edit_menu.addAction(redo_action)
        
        edit_menu.addSeparator()
        
        cut_action = QAction("Cu&t", self)
        cut_action.setShortcut("Ctrl+X")
        cut_action.triggered.connect(self.cut)
        edit_menu.addAction(cut_action)
        
        copy_action = QAction("&Copy", self)
        copy_action.setShortcut("Ctrl+C")
        copy_action.triggered.connect(self.copy)
        edit_menu.addAction(copy_action)
        
        paste_action = QAction("&Paste", self)
        paste_action.setShortcut("Ctrl+V")
        paste_action.triggered.connect(self.paste)
        edit_menu.addAction(paste_action)
        
        edit_menu.addSeparator()
        
        find_action = QAction("&Find...", self)
        find_action.setShortcut("Ctrl+F")
        find_action.triggered.connect(self.find_text)
        edit_menu.addAction(find_action)
        
        replace_action = QAction("&Replace...", self)
        replace_action.setShortcut("Ctrl+H")
        replace_action.triggered.connect(self.replace_text)
        edit_menu.addAction(replace_action)
        
        # View menu
        view_menu = menu_bar.addMenu("&View")
        
        toggle_toolbar_action = QAction("&Toolbar", self)
        toggle_toolbar_action.setCheckable(True)
        toggle_toolbar_action.setChecked(True)
        toggle_toolbar_action.triggered.connect(self.toggle_toolbar)
        view_menu.addAction(toggle_toolbar_action)
        
        toggle_formula_bar_action = QAction("&Formula Bar", self)
        toggle_formula_bar_action.setCheckable(True)
        toggle_formula_bar_action.setChecked(True)
        toggle_formula_bar_action.triggered.connect(self.toggle_formula_bar)
        view_menu.addAction(toggle_formula_bar_action)
        
        toggle_status_bar_action = QAction("&Status Bar", self)
        toggle_status_bar_action.setCheckable(True)
        toggle_status_bar_action.setChecked(True)
        toggle_status_bar_action.triggered.connect(self.toggle_status_bar)
        view_menu.addAction(toggle_status_bar_action)
        
        view_menu.addSeparator()
        
        zoom_in_action = QAction("Zoom &In", self)
        zoom_in_action.setShortcut("Ctrl++")
        zoom_in_action.triggered.connect(self.zoom_in)
        view_menu.addAction(zoom_in_action)
        
        zoom_out_action = QAction("Zoom &Out", self)
        zoom_out_action.setShortcut("Ctrl+-")
        zoom_out_action.triggered.connect(self.zoom_out)
        view_menu.addAction(zoom_out_action)
        
        # Format menu
        format_menu = menu_bar.addMenu("F&ormat")
        
        font_action = QAction("&Font...", self)
        font_action.triggered.connect(self.change_font)
        format_menu.addAction(font_action)
        
        cell_color_action = QAction("&Cell Color...", self)
        cell_color_action.triggered.connect(self.change_cell_color)
        format_menu.addAction(cell_color_action)
        
        text_color_action = QAction("&Text Color...", self)
        text_color_action.triggered.connect(self.change_text_color)
        format_menu.addAction(text_color_action)
        
        format_menu.addSeparator()
        
        conditional_format_action = QAction("&Conditional Formatting...", self)
        conditional_format_action.triggered.connect(self.apply_conditional_formatting)
        format_menu.addAction(conditional_format_action)
        
        format_menu.addSeparator()
        
        # Number formats submenu
        number_format_action = QAction("&Number Format", self)
        number_format_menu = QMenu(self)
        
        currency_format_action = QAction("&Currency", self)
        currency_format_action.triggered.connect(lambda: self.apply_number_format("currency"))
        number_format_menu.addAction(currency_format_action)
        
        percentage_format_action = QAction("&Percentage", self)
        percentage_format_action.triggered.connect(lambda: self.apply_number_format("percentage"))
        number_format_menu.addAction(percentage_format_action)
        
        comma_format_action = QAction("C&omma", self)
        comma_format_action.triggered.connect(lambda: self.apply_number_format("comma"))
        number_format_menu.addAction(comma_format_action)
        
        decimal_format_action = QAction("&Decimal Places", self)
        decimal_format_menu = QMenu(self)
        decimal_2_action = QAction("&2 Decimal Places", self)
        decimal_2_action.triggered.connect(lambda: self.apply_number_format("decimal_2"))
        decimal_4_action = QAction("&4 Decimal Places", self)
        decimal_4_action.triggered.connect(lambda: self.apply_number_format("decimal_4"))
        decimal_format_menu.addAction(decimal_2_action)
        decimal_format_menu.addAction(decimal_4_action)
        decimal_format_action.setMenu(decimal_format_menu)
        number_format_menu.addAction(decimal_format_action)
        
        scientific_format_action = QAction("&Scientific", self)
        scientific_format_action.triggered.connect(lambda: self.apply_number_format("scientific"))
        number_format_menu.addAction(scientific_format_action)
        
        number_format_action.setMenu(number_format_menu)
        format_menu.addAction(number_format_action)
        
        format_menu.addSeparator()
        
        # Cell operations
        cell_merge_action = QAction("&Merge Cells", self)
        cell_merge_action.triggered.connect(self.merge_cells)
        format_menu.addAction(cell_merge_action)
        
        cell_unmerge_action = QAction("&Unmerge Cells", self)
        cell_unmerge_action.triggered.connect(self.unmerge_cells)
        format_menu.addAction(cell_unmerge_action)
        
        format_menu.addSeparator()
        
        # Auto-fit
        autofit_menu = QMenu("Auto&fit", self)
        autofit_columns_action = QAction("Autofit &Columns", self)
        autofit_columns_action.triggered.connect(self.autofit_columns)
        autofit_menu.addAction(autofit_columns_action)
        
        autofit_rows_action = QAction("Autofit &Rows", self)
        autofit_rows_action.triggered.connect(self.autofit_rows)
        autofit_menu.addAction(autofit_rows_action)
        
        format_menu.addMenu(autofit_menu)
        
        # Data menu
        data_menu = menu_bar.addMenu("&Data")
        
        sort_asc_action = QAction("Sort &Ascending", self)
        sort_asc_action.triggered.connect(lambda: self.sort_data(True))
        data_menu.addAction(sort_asc_action)
        
        sort_desc_action = QAction("Sort &Descending", self)
        sort_desc_action.triggered.connect(lambda: self.sort_data(False))
        data_menu.addAction(sort_desc_action)
        
        data_menu.addSeparator()
        
        filter_action = QAction("&Filter...", self)
        filter_action.triggered.connect(self.filter_data)
        data_menu.addAction(filter_action)
        
        remove_filter_action = QAction("&Remove Filter", self)
        remove_filter_action.triggered.connect(self.remove_filter)
        data_menu.addAction(remove_filter_action)
        
        # Data Analysis menu (new)
        data_analysis_menu = menu_bar.addMenu("&Data Analysis")
        
        pivot_action = QAction("P&ivot Table...", self)
        pivot_action.triggered.connect(self.create_pivot_table)
        data_analysis_menu.addAction(pivot_action)
        
        data_analysis_menu.addSeparator()
        
        validate_action = QAction("Data &Validation...", self)
        validate_action.triggered.connect(self.show_data_validation)
        data_analysis_menu.addAction(validate_action)
        
        data_analysis_menu.addSeparator()
        
        remove_duplicates_action = QAction("Remove &Duplicates...", self)
        remove_duplicates_action.triggered.connect(self.remove_duplicates)
        data_analysis_menu.addAction(remove_duplicates_action)
        
        data_analysis_menu.addSeparator()
        
        what_if_action = QAction("&What-If Analysis", self)
        what_if_menu = QMenu(self)
        goal_seek_action = QAction("&Goal Seek...", self)
        goal_seek_action.triggered.connect(self.show_goal_seek)
        what_if_menu.addAction(goal_seek_action)
        what_if_action.setMenu(what_if_menu)
        data_analysis_menu.addAction(what_if_action)
        
        # Data Analysis menu additions
        data_analysis_menu.addSeparator()

        regression_action = QAction("&Regression Analysis...", self)
        regression_action.triggered.connect(self.show_regression_analysis)
        data_analysis_menu.addAction(regression_action)

        descriptive_stats_action = QAction("&Descriptive Statistics...", self)
        descriptive_stats_action.triggered.connect(self.show_descriptive_statistics)
        data_analysis_menu.addAction(descriptive_stats_action)

        forecast_action = QAction("&Forecast Sheet...", self)
        forecast_action.triggered.connect(self.create_forecast_sheet)
        data_analysis_menu.addAction(forecast_action)
        
        correlation_action = QAction("&Correlation Matrix...", self)
        correlation_action.triggered.connect(self.create_correlation_matrix)
        data_analysis_menu.addAction(correlation_action)
        
        # Sheet menu
        sheet_menu = menu_bar.addMenu("&Sheet")
        
        insert_sheet_action = QAction("&Insert Sheet", self)
        insert_sheet_action.triggered.connect(self.insert_sheet)
        sheet_menu.addAction(insert_sheet_action)
        
        rename_sheet_action = QAction("&Rename Sheet", self)
        rename_sheet_action.triggered.connect(self.rename_sheet)
        sheet_menu.addAction(rename_sheet_action)
        
        delete_sheet_action = QAction("&Delete Sheet", self)
        delete_sheet_action.triggered.connect(self.delete_sheet)
        sheet_menu.addAction(delete_sheet_action)
        
        # Insert menu
        insert_menu = menu_bar.addMenu("&Insert")
        
        chart_action = QAction("&Chart...", self)
        chart_action.triggered.connect(self.insert_chart)
        insert_menu.addAction(chart_action)
        
        sparkline_action = QAction("&Sparkline...", self)
        sparkline_action.triggered.connect(self.insert_sparkline)
        insert_menu.addAction(sparkline_action)
        
        insert_menu.addSeparator()
        
        comment_action = QAction("Co&mment", self)
        comment_action.triggered.connect(self.add_comment)
        insert_menu.addAction(comment_action)
        
        insert_menu.addSeparator()
        
        named_range_action = QAction("&Named Range...", self)
        named_range_action.triggered.connect(self.create_named_range)
        insert_menu.addAction(named_range_action)
        
        named_ranges_action = QAction("&Manage Named Ranges...", self)
        named_ranges_action.triggered.connect(self.manage_named_ranges)
        insert_menu.addAction(named_ranges_action)
        
        chart_submenu = QMenu("Advanced Charts", self)

        scatter_chart_action = QAction("Scatter Chart", self)
        scatter_chart_action.triggered.connect(lambda: self.insert_advanced_chart("scatter"))
        chart_submenu.addAction(scatter_chart_action)

        bubble_chart_action = QAction("Bubble Chart", self)
        bubble_chart_action.triggered.connect(lambda: self.insert_advanced_chart("bubble"))
        chart_submenu.addAction(bubble_chart_action)

        area_chart_action = QAction("Area Chart", self)
        area_chart_action.triggered.connect(lambda: self.insert_advanced_chart("area"))
        chart_submenu.addAction(area_chart_action)

        stock_chart_action = QAction("Stock Chart", self)
        stock_chart_action.triggered.connect(lambda: self.insert_advanced_chart("stock"))
        chart_submenu.addAction(stock_chart_action)

        insert_menu.addMenu(chart_submenu)
        
        # View menu enhancements
        freeze_panes_action = QAction("&Freeze Panes", self)
        freeze_panes_action.triggered.connect(self.freeze_panes)
        view_menu.addAction(freeze_panes_action)
        
        # Help menu - Update with more comprehensive options
        help_menu = menu_bar.addMenu("&Help")
        
        quick_help_action = QAction("&Quick Start Guide", self)
        quick_help_action.triggered.connect(self.show_quick_help)
        help_menu.addAction(quick_help_action)
        
        formulas_help_action = QAction("&Formulas and Functions", self)
        formulas_help_action.triggered.connect(self.show_formulas_help)
        help_menu.addAction(formulas_help_action)
        
        shortcuts_action = QAction("Keyboard &Shortcuts", self)
        shortcuts_action.triggered.connect(self.show_shortcuts)
        help_menu.addAction(shortcuts_action)
        
        help_menu.addSeparator()
        
        # Add "Check for Updates" option
        updates_action = QAction("Check for &Updates", self)
        updates_action.triggered.connect(self.check_for_updates)
        help_menu.addAction(updates_action)
        
        help_menu.addSeparator()
        
        about_action = QAction("&About PySpreadsheet", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
        
        # Tools menu
        tools_menu = menu_bar.addMenu("&Tools")
        
        preferences_action = QAction("&Preferences...", self)
        preferences_action.triggered.connect(self.show_preferences)
        tools_menu.addAction(preferences_action)

    def create_formula_bar(self):
        self.formula_container = QWidget()
        formula_layout = QHBoxLayout(self.formula_container)
        formula_layout.setContentsMargins(5, 2, 5, 2)
        
        cell_label = QLabel("Cell:")
        cell_label.setFixedWidth(30)
        formula_layout.addWidget(cell_label)
        
        self.current_cell_label = QLabel("A1")
        self.current_cell_label.setFixedWidth(50)
        self.current_cell_label.setFrameStyle(QLabel.StyledPanel | QLabel.Sunken)
        formula_layout.addWidget(self.current_cell_label)
        
        formula_label = QLabel("Formula:")
        formula_label.setFixedWidth(60)
        formula_layout.addWidget(formula_label)
        
        # Use the custom FormulaLineEdit
        self.formula_input = FormulaLineEdit()
        formula_layout.addWidget(self.formula_input)
        
        self.main_layout.addWidget(self.formula_container)

    def create_sheet_tabs(self):
        # Create sheet view
        self.sheet_view = SheetView(self)
        self.main_layout.addWidget(self.sheet_view)
        
        # Connect cell selection to formula bar
        self.sheet_view.currentCellChanged.connect(self.update_formula_bar)

    def connect_toolbar_actions(self):
        # Connect toolbar actions to their respective methods
        self.toolbar.new_action.triggered.connect(self.new_file)
        self.toolbar.open_action.triggered.connect(self.open_file)
        self.toolbar.save_action.triggered.connect(self.save_file)
        self.toolbar.undo_action.triggered.connect(self.undo)
        self.toolbar.redo_action.triggered.connect(self.redo)
        self.toolbar.cut_action.triggered.connect(self.cut)
        self.toolbar.copy_action.triggered.connect(self.copy)
        self.toolbar.paste_action.triggered.connect(self.paste)
        
        # Connect formatting actions
        if hasattr(self.toolbar, 'font_family'):
            self.toolbar.font_family.currentTextChanged.connect(self.toolbar.apply_font_family)
        
        if hasattr(self.toolbar, 'font_size'):
            self.toolbar.font_size.valueChanged.connect(self.toolbar.apply_font_size)
        
        if hasattr(self.toolbar, 'bold_action'):
            self.toolbar.bold_action.triggered.connect(self.toolbar.apply_bold)
        
        if hasattr(self.toolbar, 'italic_action'):
            self.toolbar.italic_action.triggered.connect(self.toolbar.apply_italic)
        
        if hasattr(self.toolbar, 'underline_action'):
            self.toolbar.underline_action.triggered.connect(self.toolbar.apply_underline)
        
        if hasattr(self.toolbar, 'align_left_action'):
            self.toolbar.align_left_action.triggered.connect(lambda: self.toolbar.apply_alignment(Qt.AlignLeft))
        
        if hasattr(self.toolbar, 'align_center_action'):
            self.toolbar.align_center_action.triggered.connect(lambda: self.toolbar.apply_alignment(Qt.AlignCenter))
        
        if hasattr(self.toolbar, 'align_right_action'):
            self.toolbar.align_right_action.triggered.connect(lambda: self.toolbar.apply_alignment(Qt.AlignRight))

    def load_settings(self):
        """Load application settings"""
        geometry = self.settings.value("geometry")
        if geometry:
            self.restoreGeometry(geometry)
            
        window_state = self.settings.value("windowState")
        if window_state:
            self.restoreState(window_state)
            
        # Other settings can be loaded here

    def closeEvent(self, event):
        """Save settings when the application is closed"""
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("windowState", self.saveState())
        
        # Check if there are unsaved changes
        if self.file_manager.get_current_file_path() and not self.prompt_save_changes():
            event.ignore()
            return
            
        event.accept()

    def current_sheet_view(self):
        """Get the currently active SheetView widget"""
        return self.sheet_view

    def new_file(self):
        """Create a new spreadsheet"""
        if self.file_manager.get_current_file_path() and not self.prompt_save_changes():
            return
            
        # Clear all sheets and create a fresh one
        self.sheet_view.clear()
        
        self.workbook = Workbook()
        self.workbook.add_sheet("Sheet1")
        self.sheet_view = SheetView(self)
        self.current_sheet_name = "Sheet1"
        
        self.sheet_view.currentCellChanged.connect(self.update_formula_bar)
        self.calculator.set_sheet_view(self.sheet_view)
        
        self.file_manager.close_file()
        self.statusBar().showMessage("New spreadsheet created")

    def prompt_save_changes(self):
        """Prompt the user to save changes if there are any"""
        response = QMessageBox.question(
            self, "Save Changes?",
            "Do you want to save your changes?",
            QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel
        )
        
        if response == QMessageBox.Save:
            return self.save_file()
        elif response == QMessageBox.Cancel:
            return False
        
        return True  # User selected Discard

    def open_file(self):
        """Open a spreadsheet file"""
        if self.file_manager.get_current_file_path() and not self.prompt_save_changes():
            return
            
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Open Spreadsheet File", "", 
            "CSV Files (*.csv);;Excel Files (*.xlsx *.xls);;All Files (*)",
            options=options
        )
        
        if not file_name:
            return
            
        try:
            # Clear all sheets first
            self.sheet_view.clear()
                
            # Determine file type and use appropriate handler
            if file_name.endswith('.csv'):
                data = read_csv(file_name)
                
                self.sheet_view = SheetView(self)
                self.sheet_view.load_data(data)
                self.calculator.set_sheet_view(self.sheet_view)
                
            elif file_name.endswith(('.xlsx', '.xls')):
                # Excel files can have multiple sheets
                workbook_data = read_excel(file_name)
                
                for sheet_name, sheet_data in workbook_data.items():
                    self.sheet_view = SheetView(self)
                    self.sheet_view.load_data(sheet_data)
                    
                # Set calculator to use first sheet
                self.calculator.set_sheet_view(self.sheet_view)
                
            else:
                self.statusBar().showMessage("Unsupported file format")
                return
                
            self.file_manager.open_file(file_name)
            self.statusBar().showMessage(f"Opened file: {file_name}")
            
        except Exception as e:
            QMessageBox.critical(self, "Error Opening File", f"An error occurred: {str(e)}")
            self.statusBar().showMessage(f"Error opening file: {str(e)}")

    def save_file(self):
        """Save the current spreadsheet"""
        current_path = self.file_manager.get_current_file_path()
        if not current_path:
            return self.save_file_as()
            
        return self._save_to_path(current_path)

    def save_file_as(self):
        """Save the spreadsheet with a new filename"""
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(
            self, "Save Spreadsheet File", "",
            "CSV Files (*.csv);;Excel Files (*.xlsx);;All Files (*)",
            options=options
        )
        
        if not file_name:
            return False
            
        return self._save_to_path(file_name)

    def _save_to_path(self, file_path):
        """Save the spreadsheet to the specified path"""
        try:
            if file_path.endswith('.csv'):
                # CSV only supports a single sheet, so use the current one
                data = self.sheet_view.get_all_data()
                write_csv(file_path, data)
                
            elif file_path.endswith(('.xlsx', '.xls')):
                # Create a dictionary with all sheets
                workbook_data = {}
                workbook_data[self.current_sheet_name] = self.sheet_view.get_all_data()
                    
                write_excel(file_path, workbook_data)
                
            else:
                # Default to CSV if extension is not recognized
                data = self.sheet_view.get_all_data()
                write_csv(file_path, data)
                
            self.file_manager.save_file(file_path, "")  # Just to update the current path
            self.statusBar().showMessage(f"Saved to: {file_path}")
            return True
            
        except Exception as e:
            QMessageBox.critical(self, "Error Saving File", f"An error occurred: {str(e)}")
            self.statusBar().showMessage(f"Error saving file: {str(e)}")
            return False

    def export_to_pdf(self):
        """Export the current sheet to PDF"""
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(
            self, "Export to PDF", "",
            "PDF Files (*.pdf);;All Files (*)",
            options=options
        )
        
        if not file_name:
            return
            
        try:
            self.sheet_view.export_to_pdf(file_name)
            self.statusBar().showMessage(f"Exported to PDF: {file_name}")
        except Exception as e:
            QMessageBox.critical(self, "Error Exporting to PDF", f"An error occurred: {str(e)}")
            self.statusBar().showMessage(f"Error exporting to PDF: {str(e)}")

    def update_formula_bar(self, row, column, prev_row, prev_column):
        """Update the formula bar with the current cell's address and content"""
        # Get the cell's address (e.g., "A1", "B2")
        col_letter = chr(65 + column) if column < 26 else chr(64 + int(column/26)) + chr(65 + (column % 26))
        cell_address = f"{col_letter}{row + 1}"
        self.current_cell_label.setText(cell_address)
        
        # Get the cell's content - show formula if it exists
        cell_item = self.sheet_view.item(row, column)
        cell_value = cell_item.text() if cell_item else ""
        
        # Temporarily disconnect signal to avoid triggering autocomplete
        self.formula_input.blockSignals(True)
        
        if cell_value:
            self.formula_input.setText(cell_value)
        else:
            self.formula_input.clear()
        
        # Re-enable signals
        self.formula_input.blockSignals(False)

    def apply_formula(self):
        """Apply the formula or value from the formula bar to the current cell"""
        current_row = self.sheet_view.currentRow()
        current_column = self.sheet_view.currentColumn()
        formula = self.formula_input.text()
        
        # Store the original formula but display the result
        if formula and formula.startswith('='):
            # This is a formula, should be evaluated
            result = self.calculator.evaluate(formula, current_row, current_column)
            
            # Create or update the cell item
            cell_item = self.sheet_view.item(current_row, current_column)
            if not cell_item:
                cell_item = QTableWidgetItem()
                self.sheet_view.setItem(current_row, current_column, cell_item)
            
            # Store the formula as the user data and display the result
            cell_item.setData(Qt.UserRole, formula)
            cell_item.setText(result)
        else:
            # Just a regular value
            cell_item = self.sheet_view.item(current_row, current_column)
            if not cell_item:
                cell_item = QTableWidgetItem()
                self.sheet_view.setItem(current_row, current_column, cell_item)
            cell_item.setText(formula)
        
        # Move to the next cell below
        self.sheet_view.setCurrentCell(current_row + 1, current_column)

    def on_cell_changed(self, item):
        """Handle cell changes and update dependent cells"""
        if not hasattr(self, 'calculator'):
            return
            
        row = item.row()
        column = item.column()
        value = item.text()
        
        # Check if it's a formula
        if value and value.startswith('='):
            # Evaluate the formula
            result = self.calculator.evaluate(value, row, column)
            # Store both the formula and the result
            item.setData(Qt.UserRole, value)
            item.setText(result)
            # Update any cells that might depend on this one
            self.update_dependent_cells(row, column)

    def update_dependent_cells(self, changed_row, changed_col):
        """Update cells that depend on the changed cell"""
        if not hasattr(self, 'calculator'):
            return
            
        # Use the calculator's dependency tracking
        if hasattr(self.calculator, 'recalculate_all'):
            self.calculator.recalculate_all()

    def undo(self):
        """Undo the last action"""
        self.sheet_view.undo()
        self.statusBar().showMessage("Undo operation")

    def redo(self):
        """Redo the last undone action"""
        self.sheet_view.redo()
        self.statusBar().showMessage("Redo operation")

    def cut(self):
        """Cut selected cells"""
        self.sheet_view.cut_cells()
        self.statusBar().showMessage("Cut cells")

    def copy(self):
        """Copy selected cells"""
        self.sheet_view.copy_cells()
        self.statusBar().showMessage("Copied cells")

    def paste(self):
        """Paste cells"""
        self.sheet_view.paste_cells()
        self.statusBar().showMessage("Pasted cells")

    # Format operations
    def change_font(self):
        """Change the font of selected cells"""
        font, ok = QFontDialog.getFont()
        if ok:
            self.sheet_view.apply_font_to_selected_cells(font)
            self.statusBar().showMessage(f"Applied font: {font.family()}, {font.pointSize()}pt")

    def change_cell_color(self):
        """Change the background color of selected cells"""
        color = QColorDialog.getColor()
        if color.isValid():
            self.sheet_view.apply_background_color_to_selected_cells(color)
            self.statusBar().showMessage("Applied cell color")

    def change_text_color(self):
        """Change the text color of selected cells"""
        color = QColorDialog.getColor()
        if color.isValid():
            self.sheet_view.apply_text_color_to_selected_cells(color)
            self.statusBar().showMessage("Applied text color")

    def apply_conditional_formatting(self):
        """Apply conditional formatting to selected cells"""
        self.sheet_view.apply_conditional_formatting()
        self.statusBar().showMessage("Applied conditional formatting")

    def apply_number_format(self, format_type):
        """Apply number formatting to selected cells"""
        self.sheet_view.apply_number_format(format_type)
        self.statusBar().showMessage(f"Applied {format_type} format")

    def merge_cells(self):
        """Merge selected cells"""
        self.sheet_view.merge_cells()
        self.statusBar().showMessage("Merged cells")

    def unmerge_cells(self):
        """Unmerge selected cells"""
        self.sheet_view.split_cells()
        self.statusBar().showMessage("Unmerged cells")

    def add_comment(self):
        """Add comment to selected cell"""
        self.sheet_view.add_comment_to_cell()
        self.statusBar().showMessage("Added cell comment")

    def autofit_columns(self):
        """Auto-fit column widths based on content"""
        self.sheet_view.autofit_columns()
        self.statusBar().showMessage("Auto-fit columns applied")

    def autofit_rows(self):
        """Auto-fit row heights based on content"""
        self.sheet_view.autofit_rows()
        self.statusBar().showMessage("Auto-fit rows applied")

    def freeze_panes(self):
        """Freeze panes at current cell position"""
        current_row = self.sheet_view.currentRow()
        current_col = self.sheet_view.currentColumn()
        self.sheet_view.freeze_panes(current_row, current_col)
        self.statusBar().showMessage(f"Froze panes at cell {chr(65 + current_col)}{current_row + 1}")

    # View operations
    def toggle_toolbar(self, checked):
        """Toggle the visibility of the toolbar"""
        self.toolbar.setVisible(checked)

    def toggle_formula_bar(self, checked):
        """Toggle the visibility of the formula bar"""
        self.formula_container.setVisible(checked)

    def toggle_status_bar(self, checked):
        """Toggle the visibility of the status bar"""
        self.statusBar().setVisible(checked)

    def zoom_in(self):
        """Zoom in on the current sheet"""
        self.sheet_view.zoom_in()
        self.statusBar().showMessage("Zoomed in")

    def zoom_out(self):
        """Zoom out on the current sheet"""
        self.sheet_view.zoom_out()
        self.statusBar().showMessage("Zoomed out")

    # Sheet operations
    def insert_sheet(self):
        """Insert a new sheet"""
        sheet_name, ok = QInputDialog.getText(self, "New Sheet", "Enter sheet name:")
        if not ok or not sheet_name:
            return
            
        # Check if the name is already used
        if self.workbook.sheet_exists(sheet_name):
            QMessageBox.warning(self, "Duplicate Name", f"A sheet named '{sheet_name}' already exists.")
            return
                
        self.sheet_view = SheetView(self)
        self.sheet_view.currentCellChanged.connect(self.update_formula_bar)
        self.statusBar().showMessage(f"Added sheet: {sheet_name}")
        
        # Update workbook model
        self.workbook.add_sheet(sheet_name)

    def rename_sheet(self):
        """Rename the current sheet"""
        current_name = self.current_sheet_name
        new_name, ok = QInputDialog.getText(self, "Rename Sheet", 
                                           "Enter new name:", text=current_name)
        if not ok or not new_name or new_name == current_name:
            return
            
        # Check if the name is already used
        if self.workbook.sheet_exists(new_name):
            QMessageBox.warning(self, "Duplicate Name", f"A sheet named '{new_name}' already exists.")
            return
                
        self.current_sheet_name = new_name
        self.statusBar().showMessage(f"Renamed sheet to: {new_name}")
        
        # Update workbook model
        self.workbook.rename_sheet(current_name, new_name)

    def delete_sheet(self):
        """Delete the current sheet"""
        if self.workbook.sheet_count() <= 1:
            QMessageBox.warning(self, "Cannot Delete", "Cannot delete the only sheet.")
            return
            
        sheet_name = self.current_sheet_name
        if QMessageBox.question(self, "Confirm Delete", 
                               f"Are you sure you want to delete the sheet '{sheet_name}'?",
                               QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
                                   
            self.sheet_view.clear()
            self.statusBar().showMessage(f"Deleted sheet: {sheet_name}")
            
            # Update workbook model
            self.workbook.remove_sheet(sheet_name)
            
    def close_sheet(self, index):
        """Close the specified sheet tab"""
        if self.workbook.sheet_count() <= 1:
            QMessageBox.warning(self, "Cannot Close", "Cannot close the only sheet.")
            return
            
        sheet_name = self.current_sheet_name
        if QMessageBox.question(self, "Confirm Close", 
                               f"Are you sure you want to close the sheet '{sheet_name}'?", 
                               QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            
            # Remove tab first
            self.sheet_view.clear()
            self.statusBar().showMessage(f"Closed sheet: {sheet_name}")
            
            # Update workbook model - wrap in try/except to prevent app crashes
            try:
                self.workbook.remove_sheet(sheet_name)
            except ValueError as e:
                # If the sheet doesn't exist in the workbook model, just log it
                print(f"Warning: {e}")
            except Exception as e:
                # Catch any other exceptions that might occur
                print(f"Error while removing sheet from workbook: {e}")

    def sheet_changed(self, index):
        """Handle sheet tab change"""
        sheet_name = self.current_sheet_name
        self.calculator.set_sheet_view(self.sheet_view)
        self.statusBar().showMessage(f"Current sheet: {sheet_name}")

    # Data operations
    def sort_data(self, ascending=True):
        """Sort the selected data"""
        self.sheet_view.sort_selected_data(ascending)
        direction = "ascending" if ascending else "descending"
        self.statusBar().showMessage(f"Sorted data {direction}")

    def filter_data(self):
        """Filter the data"""
        self.sheet_view.show_filter_dialog()
        self.statusBar().showMessage("Applied filter")

    def remove_filter(self):
        """Remove the filter"""
        self.sheet_view.remove_filter()
        self.statusBar().showMessage("Removed filter")

    # Search operations
    def find_text(self):
        """Find text in the spreadsheet"""
        if not hasattr(self, "find_dialog") or not self.find_dialog:
            self.find_dialog = FindDialog(self.sheet_view)
        
        self.find_dialog.show()
        self.find_dialog.activateWindow()
        self.find_dialog.raise_()

    def replace_text(self):
        """Replace text in the spreadsheet"""
        if not hasattr(self, "replace_dialog") or not self.replace_dialog:
            self.replace_dialog = ReplaceDialog(self.sheet_view)
        
        self.replace_dialog.show()
        self.replace_dialog.activateWindow()
        self.replace_dialog.raise_()

    # Insert operations
    def insert_chart(self):
        """Open the chart creation dialog"""
        selected_items = self.sheet_view.selectedItems()
        
        if not selected_items:
            self.statusBar().showMessage("No data selected for chart")
            return
        
        # Get the selected data
        data = []
        rows = set()
        cols = set()
        
        for item in selected_items:
            rows.add(item.row())
            cols.add(item.column())
        
        rows = sorted(list(rows))
        cols = sorted(list(cols))
        
        for row in rows:
            row_data = []
            for col in cols:
                item = self.sheet_view.item(row, col)
                value = item.text() if item else ""
                row_data.append(value)
            data.append(row_data)
            
        # Create and show the chart dialog
        dialog = ChartDialog(data, self)
        dialog.exec_()

    def create_named_range(self):
        """Create a named range for the selected cells"""
        selected_ranges = self.sheet_view.selectedRanges()
        if not selected_ranges:
            QMessageBox.warning(self, "No Selection", "Please select a range to create a named range.")
            return
            
        # Get the range as text (e.g., "A1:B5")
        range_ = selected_ranges[0]
        top_left_col = chr(65 + range_.leftColumn())
        top_left_row = range_.topRow() + 1
        bottom_right_col = chr(65 + range_.rightColumn())
        bottom_right_row = range_.bottomRow() + 1
        
        range_str = f"{top_left_col}{top_left_row}:{bottom_right_col}{bottom_right_row}"
        
        # Ask for a name
        name, ok = QInputDialog.getText(self, "Named Range", 
                                       f"Enter name for range {range_str}:")
        if ok and name:
            self.sheet_view.create_named_range(name, range_str)
            self.statusBar().showMessage(f"Created named range: {name} = {range_str}")

    def insert_sparkline(self):
        """Insert a sparkline chart"""
        # Create sparkline dialog
        dialog = QDialog(self)
        dialog.setWindowTitle("Insert Sparkline")
        dialog.setMinimumWidth(400)
        
        layout = QVBoxLayout(dialog)
        
        # Data range
        data_group = QGroupBox("Data Range")
        data_layout = QVBoxLayout()
        
        range_layout = QHBoxLayout()
        range_label = QLabel("Data Range:")
        range_edit = QLineEdit()
        range_layout.addWidget(range_label)
        range_layout.addWidget(range_edit)
        data_layout.addLayout(range_layout)
        
        location_layout = QHBoxLayout()
        location_label = QLabel("Location:")
        location_edit = QLineEdit()
        location_layout.addWidget(location_label)
        location_layout.addWidget(location_edit)
        data_layout.addLayout(location_layout)
        
        data_group.setLayout(data_layout)
        layout.addWidget(data_group)
        
        # Sparkline type
        type_group = QGroupBox("Type")
        type_layout = QHBoxLayout()
        
        line_radio = QRadioButton("Line")
        line_radio.setChecked(True)
        type_layout.addWidget(line_radio)
        
        column_radio = QRadioButton("Column")
        type_layout.addWidget(column_radio)
        
        win_loss_radio = QRadioButton("Win/Loss")
        type_layout.addWidget(win_loss_radio)
        
        type_group.setLayout(type_layout)
        layout.addWidget(type_group)
        
        # Buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        if dialog.exec_() == QDialog.Accepted:
            # This would normally insert a sparkline, but for now, just report what was done
            self.statusBar().showMessage("Sparkline feature not fully implemented")

    def create_pivot_table(self):
        """Create a pivot table from the selected data"""
        selected_items = self.sheet_view.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "No Selection", "Please select the data for your pivot table.")
            return
        
        # Extract the data from the selection
        rows = set()
        cols = set()
        
        for item in selected_items:
            rows.add(item.row())
            cols.add(item.column())
        
        # Get ordered lists
        rows_list = sorted(list(rows))
        cols_list = sorted(list(cols))
        
        # Create pivot dialog
        dialog = QDialog(self)
        dialog.setWindowTitle("Create Pivot Table")
        dialog.setMinimumSize(700, 500)
        
        layout = QVBoxLayout(dialog)
        
        # Areas for fields
        fields_group = QGroupBox("Pivot Table Fields")
        fields_layout = QGridLayout()
        
        # Available fields
        available_label = QLabel("Available Fields:")
        available_list = QListWidget()
        
        # Get headers from first row of selection
        for col_idx in cols_list:
            item = self.sheet_view.item(rows_list[0], col_idx)
            header = item.text() if item else f"Column {col_idx+1}"
            available_list.addItem(header)
            
        fields_layout.addWidget(available_label, 0, 0)
        fields_layout.addWidget(available_list, 1, 0, 3, 1)
        
        # Pivot areas
        row_label = QLabel("Row Labels:")
        row_list = QListWidget()
        fields_layout.addWidget(row_label, 0, 1)
        fields_layout.addWidget(row_list, 1, 1)
        
        column_label = QLabel("Column Labels:")
        column_list = QListWidget()
        fields_layout.addWidget(column_label, 0, 2)
        fields_layout.addWidget(column_list, 1, 2)
        
        values_label = QLabel("Values:")
        values_list = QListWidget()
        fields_layout.addWidget(values_label, 2, 1, 1, 2)
        fields_layout.addWidget(values_list, 3, 1, 1, 2)
        
        # Add buttons
        button_layout = QVBoxLayout()
        add_to_row = QPushButton("Add to Row Labels")
        add_to_column = QPushButton("Add to Column Labels")
        add_to_values = QPushButton("Add to Values")
        remove = QPushButton("Remove Field")
        
        button_layout.addWidget(add_to_row)
        button_layout.addWidget(add_to_column)
        button_layout.addWidget(add_to_values)
        button_layout.addWidget(remove)
        
        fields_layout.addLayout(button_layout, 1, 3, 3, 1)
        fields_group.setLayout(fields_layout)
        layout.addWidget(fields_group)
        
        # Dialog buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        # Connect the add/remove buttons
        def add_to_row_labels():
            for item in available_list.selectedItems():
                row_list.addItem(item.text())
                
        def add_to_column_labels():
            for item in available_list.selectedItems():
                column_list.addItem(item.text())
                
        def add_to_values_area():
            for item in available_list.selectedItems():
                values_list.addItem(f"Sum of {item.text()}")
                
        def remove_field():
            for list_widget in [row_list, column_list, values_list]:
                for item in list_widget.selectedItems():
                    list_widget.takeItem(list_widget.row(item))
                    
        add_to_row.clicked.connect(add_to_row_labels)
        add_to_column.clicked.connect(add_to_column_labels)
        add_to_values.clicked.connect(add_to_values_area)
        remove.clicked.connect(remove_field)
        
        if dialog.exec_() == QDialog.Accepted:
            self.statusBar().showMessage("Pivot table functionality not fully implemented")
            # In a full implementation, we would:
            # 1. Create a new sheet for the pivot table
            # 2. Generate the pivot table structure based on the selections
            # 3. Calculate the aggregations (sums, etc.)
            # 4. Format the pivot table appropriately

    def show_data_validation(self):
        """Show data validation dialog"""
        selected_ranges = self.sheet_view.selectedRanges()
        if not selected_ranges:
            QMessageBox.warning(self, "No Selection", "Please select cells for data validation.")
            return
        
        # Create data validation dialog
        dialog = QDialog(self)
        dialog.setWindowTitle("Data Validation")
        dialog.setMinimumWidth(500)
        
        layout = QVBoxLayout(dialog)
        
        # Settings tab
        settings_group = QGroupBox("Validation criteria")
        settings_layout = QVBoxLayout()
        
        # Allow dropdown
        allow_layout = QHBoxLayout()
        allow_label = QLabel("Allow:")
        allow_combo = QComboBox()
        allow_combo.addItems(["Any value", "Whole number", "Decimal", "List", "Date", "Text length"])
        allow_layout.addWidget(allow_label)
        allow_layout.addWidget(allow_combo)
        settings_layout.addLayout(allow_layout)
        
        # Data dropdown (changes based on Allow selection)
        data_layout = QHBoxLayout()
        data_label = QLabel("Data:")
        data_combo = QComboBox()
        data_layout.addWidget(data_label)
        data_layout.addWidget(data_combo)
        settings_layout.addLayout(data_layout)
        
        # Min/max inputs
        minmax_layout = QHBoxLayout()
        min_label = QLabel("Minimum:")
        min_input = QLineEdit()
        max_label = QLabel("Maximum:")
        max_input = QLineEdit()
        minmax_layout.addWidget(min_label)
        minmax_layout.addWidget(min_input)
        minmax_layout.addWidget(max_label)
        minmax_layout.addWidget(max_input)
        settings_layout.addLayout(minmax_layout)
        
        # Source input for List validation
        source_layout = QHBoxLayout()
        source_label = QLabel("Source:")
        source_input = QLineEdit()
        source_label.hide()
        source_input.hide()
        source_layout.addWidget(source_label)
        source_layout.addWidget(source_input)
        settings_layout.addLayout(source_layout)
        
        # Ignore blank checkbox
        ignore_blank = QCheckBox("Ignore blank")
        ignore_blank.setChecked(True)
        settings_layout.addWidget(ignore_blank)
        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)
        
        # Error alert group
        error_group = QGroupBox("Error alert")
        error_group.setCheckable(True)
        error_group.setChecked(True)
        error_layout = QVBoxLayout()
        
        # Show error alert checkbox
        style_layout = QHBoxLayout()
        style_label = QLabel("Style:")
        style_combo = QComboBox()
        style_combo.addItems(["Stop", "Warning", "Information"])
        style_layout.addWidget(style_label)
        style_layout.addWidget(style_combo)
        error_layout.addLayout(style_layout)
        
        # Title and message
        title_layout = QHBoxLayout()
        title_label = QLabel("Title:")
        title_input = QLineEdit("Data Validation Error")
        title_layout.addWidget(title_label)
        title_layout.addWidget(title_input)
        error_layout.addLayout(title_layout)
        
        message_layout = QHBoxLayout()
        message_label = QLabel("Message:")
        message_input = QLineEdit("The value you entered is not valid.")
        message_layout.addWidget(message_label)
        message_layout.addWidget(message_input)
        error_layout.addLayout(message_layout)
        
        error_group.setLayout(error_layout)
        layout.addWidget(error_group)
        
        # Buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        # Handle allow type changes
        def update_data_options(index):
            allow_type = allow_combo.currentText()
            data_combo.clear()
            
            # Show/hide appropriate controls
            if allow_type == "Any value":
                data_label.hide()
                data_combo.hide()
                min_label.hide()
                min_input.hide()
                max_label.hide()
                max_input.hide()
                source_label.hide()
                source_input.hide()
            elif allow_type == "List":
                data_label.hide()
                data_combo.hide()
                min_label.hide()
                min_input.hide()
                max_label.hide()
                max_input.hide()
                source_label.show()
                source_input.show()
            else:
                data_label.show()
                data_combo.show()
                
                if allow_type in ["Whole number", "Decimal", "Date", "Text length"]:
                    data_combo.addItems(["between", "not between", "equal to", "not equal to", 
                                        "greater than", "less than"])
                    
                    # Show min/max for 'between' and 'not between'
                    if data_combo.currentText() in ["between", "not between"]:
                        min_label.show()
                        min_input.show()
                        max_label.show()
                        max_input.show()
                    else:
                        min_label.show()
                        min_input.show()
                        max_label.hide()
                        max_input.hide()
                        
                    source_label.hide()
                    source_input.hide()
                    
        allow_combo.currentIndexChanged.connect(update_data_options)
        update_data_options(0)  # Initialize with first option
        
        if dialog.exec_() == QDialog.Accepted:
            self.statusBar().showMessage("Data validation applied")

    def show_goal_seek(self):
        """Show goal seek dialog"""
        # Create simple goal seek dialog
        dialog = QDialog(self)
        dialog.setWindowTitle("Goal Seek")
        dialog.setMinimumWidth(300)
        
        layout = QGridLayout(dialog)
        
        # Set cell
        layout.addWidget(QLabel("Set cell:"), 0, 0)
        set_cell_input = QLineEdit()
        layout.addWidget(set_cell_input, 0, 1)
        
        # To value
        layout.addWidget(QLabel("To value:"), 1, 0)
        to_value_input = QLineEdit()
        layout.addWidget(to_value_input, 1, 1)
        
        # By changing cell
        layout.addWidget(QLabel("By changing cell:"), 2, 0)
        by_changing_input = QLineEdit()
        layout.addWidget(by_changing_input, 2, 1)
        
        # Buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons, 3, 0, 1, 2)
        
        if dialog.exec_() == QDialog.Accepted:
            # This would implement the actual goal seek algorithm
            self.statusBar().showMessage("Goal seek feature not fully implemented")

    def remove_duplicates(self):
        """Remove duplicate rows from selection"""
        selected_ranges = self.sheet_view.selectedRanges()
        if not selected_ranges:
            QMessageBox.warning(self, "No Selection", "Please select a range to remove duplicates from.")
            return
        
        # Use first range
        range_ = selected_ranges[0]
        top = range_.topRow()
        left = range_.leftColumn()
        bottom = range_.bottomRow()
        right = range_.rightColumn()
        
        # Extract data
        data = []
        for row in range(top, bottom + 1):
            row_data = []
            for col in range(left, right + 1):
                item = self.sheet_view.item(row, col)
                cell_value = item.text() if item else ""
                row_data.append(cell_value)
            data.append(tuple(row_data))  # Use tuple for hashability
        
        # Find unique rows
        unique_rows = []
        seen = set()
        duplicate_count = 0
        
        for i, row_tuple in enumerate(data):
            if row_tuple not in seen:
                seen.add(row_tuple)
                unique_rows.append((i, row_tuple))
            else:
                duplicate_count += 1
        
        if duplicate_count == 0:
            QMessageBox.information(self, "No Duplicates", "No duplicate values were found.")
            return
        
        # Confirm removal
        response = QMessageBox.question(
            self, "Remove Duplicates",
            f"Found {duplicate_count} duplicate row(s). Do you want to remove them?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if response == QMessageBox.Yes:
            # Block signals during update
            self.sheet_view.blockSignals(True)
            
            # Clear the range
            for row in range(top, bottom + 1):
                for col in range(left, right + 1):
                    item = self.sheet_view.item(row, col)
                    if item:
                        item.setText("")
            
            # Fill with unique data
            for i, (original_index, row_data) in enumerate(unique_rows):
                for j, value in enumerate(row_data):
                    item = self.sheet_view.item(top + i, left + j)
                    if item:
                        item.setText(value)
            
            self.sheet_view.blockSignals(False)
            
            self.statusBar().showMessage(f"Removed {duplicate_count} duplicate row(s)")

    # Help operations
    def show_quick_help(self):
        """Show the Quick Start Guide dialog"""
        from .dialogs.help_dialog import QuickStartDialog
        dialog = QuickStartDialog(self)
        dialog.exec_()

    def show_formulas_help(self):
        """Show the Formulas and Functions help dialog"""
        from .dialogs.help_dialog import FormulasHelpDialog
        dialog = FormulasHelpDialog(self)
        dialog.exec_()

    def show_shortcuts(self):
        """Show the keyboard shortcuts dialog"""
        from .dialogs.help_dialog import ShortcutsDialog
        dialog = ShortcutsDialog(self)
        dialog.exec_()

    def check_for_updates(self):
        """Check for software updates"""
        QMessageBox.information(self, "Check for Updates", 
                              "You are running the latest version of PySpreadsheet.")

    def show_about(self):
        """Show the About dialog"""
        QMessageBox.about(self, "About PySpreadsheet", 
                        """<h1>PySpreadsheet</h1>
                        <p>A modern spreadsheet application built with PyQt5.</p>
                        <p>Created by Aryan Daga</p>
                        <p>Version 1.0</p>
                        <p>This application provides spreadsheet functionality including:</p>
                        <ul>
                            <li>Formula calculation</li>
                            <li>Cell formatting</li>
                            <li>Data analysis tools</li>
                            <li>Charts and visualizations</li>
                            <li>Import/export from various formats</li>
                            <li>Advanced features like pivot tables</li>
                        </ul>
                        <p>Copyright © 2025 Aryan Daga. All rights reserved.</p>""")

    def show_preferences(self):
        """Show the preferences dialog"""
        dialog = PreferencesDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            # Apply theme immediately
            theme = dialog.preferences.get("theme", "light")
            self.apply_theme(theme)
            
            # Apply other settings
            self.apply_preferences(dialog.preferences)

    def apply_theme(self, theme_name):
        """Apply the selected theme"""
        from .style_manager import toggle_theme
        toggle_theme(self, theme_name)

    def apply_preferences(self, preferences):
        """Apply the selected preferences to the application"""
        # Update grid lines
        show_gridlines = preferences.get("show_gridlines", True)
        self.sheet_view.setShowGrid(show_gridlines)
        
        # Update grid color
        grid_color = QColor(preferences.get("grid_color", "#d0d0d0"))
        
        # Update cell and text colors for new cells
        self.default_cell_color = QColor(preferences.get("default_cell_color", "#FFFFFF"))
        self.default_text_color = QColor(preferences.get("default_text_color", "#000000"))
        
        # Update status
        self.statusBar().showMessage("Preferences applied", 3000)

    # Version Control Methods
    def save_document_version(self):
        """Save the current document state as a version"""
        if not hasattr(self, 'version_control'):
            return
            
        current_path = self.file_manager.get_current_file_path()
        if not current_path:
            QMessageBox.warning(self, "No File", "You need to save the file first.")
            return
            
        # Get comment from user
        comment, ok = QInputDialog.getText(
            self, "Save Version", 
            "Enter a comment for this version (optional):"
        )
        
        if not ok:
            return
            
        # Get current data
        data = {}
        data[self.current_sheet_name] = self.sheet_view.get_all_data()
            
        # Save version
        version_info = self.version_control.save_version(current_path, data, comment)
        if version_info:
            self.statusBar().showMessage(f"Version saved: {version_info['timestamp']}")
        else:
            QMessageBox.critical(self, "Error", "Failed to save version.")

    def view_version_history(self):
        """Show dialog with version history"""
        if not hasattr(self, 'version_control'):
            return
            
        current_path = self.file_manager.get_current_file_path()
        if not current_path:
            QMessageBox.warning(self, "No File", "You need to save the file first.")
            return
            
        # Get document ID and versions
        document_id = self.version_control.get_document_id_from_path(current_path)
        versions = self.version_control.get_versions(document_id)
        
        if not versions:
            QMessageBox.information(self, "No Versions", "No version history found for this document.")
            return
            
        # Create dialog to show versions
        dialog = QDialog(self)
        dialog.setWindowTitle("Version History")
        dialog.setMinimumWidth(600)
        dialog.setMinimumHeight(400)
        
        layout = QVBoxLayout(dialog)
        
        # Create a list widget to display versions
        version_list = QListWidget()
        layout.addWidget(version_list)
        
        # Add versions to the list
        for version in versions:
            timestamp = datetime.datetime.fromisoformat(version['timestamp']).strftime("%Y-%m-%d %H:%M:%S")
            comment = version['comment'] or "(No comment)"
            list_item = f"{timestamp} - {comment}"
            version_list.addItem(list_item)
            
        # Add buttons
        button_layout = QHBoxLayout()
        
        restore_button = QPushButton("Restore Selected Version")
        restore_button.clicked.connect(lambda: self.restore_version(dialog, document_id, versions, version_list.currentRow()))
        button_layout.addWidget(restore_button)
        
        delete_button = QPushButton("Delete Selected Version")
        delete_button.clicked.connect(lambda: self.delete_version(dialog, document_id, versions, version_list.currentRow()))
        button_layout.addWidget(delete_button)
        
        close_button = QPushButton("Close")
        close_button.clicked.connect(dialog.reject)
        button_layout.addWidget(close_button)
        
        layout.addLayout(button_layout)
        
        # Show dialog
        dialog.exec_()

    def restore_version(self, dialog, document_id, versions, index):
        """Restore to a selected version"""
        if index < 0 or index >= len(versions):
            return
            
        # Ask for confirmation
        version_timestamp = datetime.datetime.fromisoformat(versions[index]['timestamp']).strftime("%Y-%m-%d %H:%M:%S")
        response = QMessageBox.question(
            dialog, 
            "Confirm Restore",
            f"Are you sure you want to restore to version from {version_timestamp}?\n\nCurrent data will be saved as a new version.",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if response != QMessageBox.Yes:
            return
            
        # Restore the version
        version_id = versions[index]['version_id']
        data = self.version_control.restore_version(document_id, version_id)
        
        if data:
            # Load the data
            self.load_version_data(data)
            dialog.accept()  # Close the dialog
            self.statusBar().showMessage(f"Restored to version from {version_timestamp}")
        else:
            QMessageBox.critical(dialog, "Error", "Failed to restore version.")

    def delete_version(self, dialog, document_id, versions, index):
        """Delete a selected version"""
        if index < 0 or index >= len(versions):
            return
            
        # Ask for confirmation
        version_timestamp = datetime.datetime.fromisoformat(versions[index]['timestamp']).strftime("%Y-%m-%d %H:%M:%S")
        response = QMessageBox.question(
            dialog, 
            "Confirm Delete",
            f"Are you sure you want to delete the version from {version_timestamp}?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if response != QMessageBox.Yes:
            return
            
        # Delete the version
        version_id = versions[index]['version_id']
        success = self.version_control.delete_version(document_id, version_id)
        
        if success:
            # Refresh the list
            versions.pop(index)
            version_list = dialog.findChild(QListWidget)
            version_list.takeItem(index)
            self.statusBar().showMessage(f"Deleted version from {version_timestamp}")
        else:
            QMessageBox.critical(dialog, "Error", "Failed to delete version.")

    def load_version_data(self, data):
        """Load version data into the sheet tabs"""
        # Clear all sheets first
        self.sheet_view.clear()
            
        # Load each sheet
        for sheet_name, sheet_data in data.items():
            self.sheet_view = SheetView(self)
            self.sheet_view.load_data(sheet_data)
            
        # Set calculator to use first sheet
        self.calculator.set_sheet_view(self.sheet_view)

    # Advanced Data Analysis Methods
    def show_regression_analysis(self):
        """Show regression analysis dialog"""
        selected_ranges = self.sheet_view.selectedRanges()
        if not selected_ranges or len(selected_ranges) != 1:
            QMessageBox.warning(self, "Selection Required", "Please select a range with two columns (X and Y values)")
            return
            
        range_ = selected_ranges[0]
        # Extract data
        if range_.columnCount() != 2:
            QMessageBox.warning(self, "Invalid Selection", "Please select exactly two columns of data (X and Y values)")
            return
            
        x_values = []
        y_values = []
        for row in range(range_.topRow(), range_.bottomRow() + 1):
            x_item = self.sheet_view.item(row, range_.leftColumn())
            y_item = self.sheet_view.item(row, range_.leftColumn() + 1)
            
            if x_item and y_item:
                try:
                    x_val = float(x_item.text())
                    y_val = float(y_item.text())
                    x_values.append(x_val)
                    y_values.append(y_val)
                except ValueError:
                    # Skip non-numeric values
                    pass
                    
        if len(x_values) < 2:
            QMessageBox.warning(self, "Insufficient Data", "Need at least 2 numeric data points for regression")
            return
            
        # Perform regression analysis
        import numpy as np
        from scipy import stats
        
        # Perform regression analysis
        slope, intercept, r_value, p_value, std_err = stats.linregress(x_values, y_values)
        
        # Create result sheet
        sheet_name = "Regression Analysis"
        result_sheet = SheetView(self)
        self.sheet_view = result_sheet
        
        # Format the results
        results = [
            ["Regression Statistics"],
            ["Slope", f"{slope:.6f}"],
            ["Intercept", f"{intercept:.6f}"],
            ["R-squared", f"{r_value**2:.6f}"],
            ["R", f"{r_value:.6f}"],
            ["P-value", f"{p_value:.6f}"],
            ["Standard Error", f"{std_err:.6f}"],
            [],
            ["Equation", f"y = {slope:.4f}x + {intercept:.4f}"]
        ]
        
        # Load results into the sheet
        result_sheet.load_data(results)
        
        # Create a chart of the data with regression line
        self.statusBar().showMessage("Regression analysis complete")

    def show_descriptive_statistics(self):
        """Show descriptive statistics for selected data"""
        selected_ranges = self.sheet_view.selectedRanges()
        if not selected_ranges:
            QMessageBox.warning(self, "Selection Required", "Please select a range of data")
            return
            
        range_ = selected_ranges[0]
        # Extract data
        
        # Extract all numeric values from the selection
        data = []
        for row in range(range_.topRow(), range_.bottomRow() + 1):
            for col in range(range_.leftColumn(), range_.rightColumn() + 1):
                item = self.sheet_view.item(row, col)
                if item:
                    try:
                        value = float(item.text())
                        data.append(value)
                    except ValueError:
                        # Skip non-numeric values
                        pass
                        
        if not data:
            QMessageBox.warning(self, "No Numeric Data", "No numeric data found in selection")
            return
                        
        # Calculate statistics
        import numpy as np
        from scipy import stats
        
        # Create statistics sheet
        sheet_name = "Statistics"
        result_sheet = SheetView(self)
        self.sheet_view = result_sheet
        
        # Format the results
        results = [
            ["Descriptive Statistics", "", ""],
            ["Count", len(data), ""],
            ["Mean", np.mean(data), ""],
            ["Median", np.median(data), ""],
            ["Mode", stats.mode(data, keepdims=False)[0], ""],
            ["Standard Deviation", np.std(data, ddof=1), ""],
            ["Variance", np.var(data, ddof=1), ""],
            ["Minimum", np.min(data), ""],
            ["Maximum", np.max(data), ""],
            ["Range", np.max(data) - np.min(data), ""],
            ["Sum", np.sum(data), ""],
            ["Q1 (25th Percentile)", np.percentile(data, 25), ""],
            ["Q2 (50th Percentile)", np.percentile(data, 50), ""],
            ["Q3 (75th Percentile)", np.percentile(data, 75), ""],
            ["Skewness", stats.skew(data), ""],
            ["Kurtosis", stats.kurtosis(data), ""]
        ]
        
        # Load results into the sheet
        result_sheet.load_data(results)
        self.statusBar().showMessage("Descriptive statistics generated")
        
    def create_forecast_sheet(self):
        """Create a forecast sheet using time series data"""
        selected_ranges = self.sheet_view.selectedRanges()
        if not selected_ranges or len(selected_ranges) != 1:
            QMessageBox.warning(self, "Selection Required", "Please select a range with at least one column of data")
            return
            
        range_ = selected_ranges[0]
        # Extract data (assume it's time series data)
        data = []
        for row in range(range_.topRow(), range_.bottomRow() + 1):
            item = self.sheet_view.item(row, range_.leftColumn())
            if item:
                try:
                    value = float(item.text())
                    data.append(value)
                except ValueError:
                    # Skip non-numeric values
                    pass
        
        if len(data) < 5:
            QMessageBox.warning(self, "Insufficient Data", "Need at least 5 data points for forecast")
            return
            
        # Simple forecasting using exponential smoothing
        
        # How many periods to forecast
        periods_dialog = QInputDialog(self)
        periods_dialog.setWindowTitle("Forecast Periods")
        periods_dialog.setLabelText("Enter number of periods to forecast:")
        periods_dialog.setIntValue(5)
        periods_dialog.setIntRange(1, 100)
        
        if not periods_dialog.exec_():
            return
            
        forecast_periods = periods_dialog.intValue()
        
        # Simple exponential smoothing with alpha = 0.3
        alpha = 0.3
        forecast = [data[0]]
        for i in range(1, len(data)):
            forecast.append(alpha * data[i] + (1 - alpha) * forecast[i-1])
        
        # Forecast future values
        last_forecast = forecast[-1]
        future_forecast = []
        for _ in range(forecast_periods):
            last_forecast = alpha * last_forecast + (1 - alpha) * last_forecast
            future_forecast.append(last_forecast)
        
        # Create forecast sheet
        sheet_name = "Forecast"
        forecast_sheet = SheetView(self)
        self.sheet_view = forecast_sheet
        
        # Format the results
        results = [
            ["Period", "Actual", "Forecast", ""],
        ]
        
        # Add historical data and forecast
        for i, (actual, pred) in enumerate(zip(data, forecast)):
            results.append([i+1, actual, pred, ""])
        
        # Add future forecast
        for i, pred in enumerate(future_forecast):
            results.append([len(data) + i + 1, "", pred, ""])
            
        # Load results into the sheet
        forecast_sheet.load_data(results)
        self.statusBar().showMessage("Forecast sheet created")

    def create_correlation_matrix(self):
        """Create a correlation matrix from selected data"""
        selected_ranges = self.sheet_view.selectedRanges()
        if not selected_ranges:
            QMessageBox.warning(self, "Selection Required", "Please select numeric data")
            return
            
        # Extract data from selection
        range_ = selected_ranges[0]
        data = []
        headers = []
        
        # Get headers from first row
        for col in range(range_.leftColumn(), range_.rightColumn() + 1):
            item = self.sheet_view.item(range_.topRow(), col)
            headers.append(item.text() if item else f"Col {col+1}")
        
        # Get data from remaining rows
        for row in range(range_.topRow() + 1, range_.bottomRow() + 1):
            row_data = []
            for col in range(range_.leftColumn(), range_.rightColumn() + 1):
                item = self.sheet_view.item(row, col)
                try:
                    value = float(item.text() if item else 0)
                    row_data.append(value)
                except ValueError:
                    row_data.append(0)  # Non-numeric values as 0
            data.append(row_data)
        
        # Calculate correlation matrix
        try:
            import numpy as np
            import pandas as pd
            
            if not data:
                QMessageBox.warning(self, "No Data", "No valid data for correlation analysis")
                return
                
            df = pd.DataFrame(data, columns=headers)
            corr_matrix = df.corr()
            
            # Create a new sheet with the correlation matrix
            result_sheet = SheetView(self)
            
            # Format the results - create a list of lists with the data
            results = [["Correlation Matrix"]]
            results.append([""] + list(corr_matrix.columns))
            
            for i, row in enumerate(corr_matrix.values):
                results.append([corr_matrix.index[i]] + [f"{val:.4f}" for val in row])
            
            # Load data into the sheet
            result_sheet.load_data(results)
            
            # Replace the current sheet with the result sheet
            current_index = 0
            self.sheet_view = result_sheet
            
            self.statusBar().showMessage("Correlation matrix created")
        except ImportError:
            QMessageBox.warning(self, "Missing Libraries", 
                              "Please install numpy and pandas to use correlation analysis")
        except Exception as e:
            QMessageBox.warning(self, "Analysis Error", f"Error creating correlation matrix: {str(e)}")
            
    # Advanced Chart Methods
    def insert_advanced_chart(self, chart_type):
        """Create advanced chart types"""
        selected_ranges = self.sheet_view.selectedRanges()
        if not selected_ranges:
            QMessageBox.warning(self, "No Selection", "Please select data for the chart")
            return
            
        range_ = selected_ranges[0]
        # Extract data
        
        # Extract headers (first row)
        headers = []
        for col in range(range_.leftColumn(), range_.rightColumn() + 1):
            item = self.sheet_view.item(range_.topRow(), col)
            header = item.text() if item else f"Series {col - range_.leftColumn() + 1}"
            headers.append(header)
        
        # Extract data (remaining rows)
        data = []
        for row in range(range_.topRow() + 1, range_.bottomRow() + 1):
            row_data = []
            for col in range(range_.leftColumn(), range_.rightColumn() + 1):
                item = self.sheet_view.item(row, col)
                try:
                    value = float(item.text() if item else 0)
                except ValueError:
                    value = 0
                row_data.append(value)
            data.append(row_data)
        
        if not data:
            QMessageBox.warning(self, "No Data", "No valid data for chart")
            return
            
        # Convert to numpy array for easier manipulation
        import numpy as np
        data_array = np.array(data)
        
        # Create chart dialog
        chart_dialog = QDialog(self)
        chart_dialog.setWindowTitle(f"{chart_type.title()} Chart")
        chart_dialog.resize(800, 600)
        
        layout = QVBoxLayout(chart_dialog)
        
        # Create matplotlib figure
        import matplotlib.pyplot as plt
        from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
        
        fig = plt.figure(figsize=(8, 6))
        canvas = FigureCanvas(fig)
        layout.addWidget(canvas)
        
        # Different chart types
        if chart_type == "scatter":
            # Use first column for X, second for Y
            if data_array.shape[1] >= 2:
                plt.scatter(data_array[:, 0], data_array[:, 1])
                plt.xlabel(headers[0])
                plt.ylabel(headers[1])
            else:
                plt.scatter(range(len(data_array)), data_array[:, 0])
                plt.xlabel("Index")
                plt.ylabel(headers[0])
            
            plt.grid(True, linestyle='--', alpha=0.7)
            
        elif chart_type == "bubble":
            # Use first column for X, second for Y, third for bubble size
            if data_array.shape[1] >= 3:
                plt.scatter(data_array[:, 0], data_array[:, 1], s=data_array[:, 2]*10)
                plt.xlabel(headers[0])
                plt.ylabel(headers[1])
            else:
                QMessageBox.warning(self, "Insufficient Columns", "Bubble chart requires at least 3 columns")
                return
                
        elif chart_type == "area":
            # Area chart uses each column as a series
            x = range(len(data_array))
            for i in range(data_array.shape[1]):
                plt.fill_between(x, data_array[:, i], alpha=0.3)
                plt.plot(x, data_array[:, i], label=headers[i])  # Fixed: label is a parameter, not a function
            plt.legend()
            
        elif chart_type == "stock":
            # Stock chart (OHLC) - needs 4 columns: Open, High, Low, Close
            if data_array.shape[1] >= 4:
                # Format data for candlestick
                from mplfinance.original_flavor import candlestick_ohlc
                
                ohlc_data = []
                for i, row in enumerate(data_array):
                    ohlc_data.append([i, row[0], row[1], row[2], row[3]])
                    
                candlestick_ohlc(plt.gca(), ohlc_data, width=0.6)
                plt.xlabel("Period")
                plt.ylabel("Price")
            else:
                QMessageBox.warning(self, "Insufficient Columns", 
                                 "Stock chart requires at least 4 columns (Open, High, Low, Close)")
                return
                
        plt.title(f"{chart_type.title()} Chart")
        plt.grid(True, linestyle='--', alpha=0.7)
        fig.tight_layout()
        
        # Add buttons
        button_layout = QHBoxLayout()
        save_button = QPushButton("Save Chart")
        close_button = QPushButton("Close")
        
        # Connect buttons
        save_button.clicked.connect(lambda: self.save_chart(fig))
        close_button.clicked.connect(chart_dialog.reject)
        
        button_layout.addWidget(save_button)
        button_layout.addWidget(close_button)
        layout.addLayout(button_layout)
        
        # Show chart dialog
        chart_dialog.exec_()

    def save_chart(self, figure):
        """Save the chart figure to a file"""
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(
            self, "Save Chart", "", 
            "PNG Files (*.png);;JPEG Files (*.jpg);;PDF Files (*.pdf);;SVG Files (*.svg)",
            options=options
        )
        
        if file_name:
            figure.savefig(file_name, dpi=300, bbox_inches='tight')
            self.statusBar().showMessage(f"Chart saved to {file_name}")
            
    def manage_named_ranges(self):
        """Open the named range manager dialog"""
        from .dialogs.named_range_manager import NamedRangeManager
        
        # Get current named ranges
        named_ranges = getattr(self, 'named_ranges', {})
        
        dialog = NamedRangeManager(self, named_ranges)
        if dialog.exec_() == QDialog.Accepted:
            # Save named ranges
            self.named_ranges = dialog.named_ranges
            self.statusBar().showMessage("Named ranges updated")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())