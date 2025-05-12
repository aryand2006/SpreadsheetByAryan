from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QComboBox, QPushButton, QLabel, QHBoxLayout, 
    QGroupBox, QCheckBox, QColorDialog, QTableWidget, QTableWidgetItem,
    QDialogButtonBox, QFileDialog, QSplitter, QWidget, QGridLayout
)
from PyQt5.QtChart import (
    QChart, QChartView, QBarSeries, QBarSet, QValueAxis, QBarCategoryAxis, 
    QLineSeries, QPieSeries, QScatterSeries, QSplineSeries, QAreaSeries, 
    QPercentBarSeries, QStackedBarSeries, QPieSlice
)

from PyQt5.QtCore import Qt, QPointF, QMargins
from PyQt5.QtGui import QPainter, QColor, QBrush, QPen, QFont, QPixmap

class ChartDialog(QDialog):
    def __init__(self, data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Chart Creator")
        self.resize(900, 600)
        self.data = data
        
        # Main layout
        self.layout = QVBoxLayout(self)
        
        # Create splitter for resizable panels
        self.splitter = QSplitter(Qt.Horizontal)
        
        # Left panel - Options
        self.left_panel = QWidget()
        self.left_layout = QVBoxLayout(self.left_panel)
        self.create_options_panel()
        self.splitter.addWidget(self.left_panel)
        
        # Right panel - Chart preview
        self.right_panel = QWidget()
        self.right_layout = QVBoxLayout(self.right_panel)
        self.create_chart_panel()
        self.splitter.addWidget(self.right_panel)
        
        # Set initial splitter sizes
        self.splitter.setSizes([300, 600])
        
        # Add splitter to main layout
        self.layout.addWidget(self.splitter)
        
        # Bottom buttons
        self.create_buttons()
        
        # Process data and create initial chart
        self.process_data()
        self.update_chart()

    def create_options_panel(self):
        """Create the options panel with chart settings"""
        # Chart type selection
        chart_type_group = QGroupBox("Chart Type")
        chart_type_layout = QVBoxLayout()
        
        self.chart_type_combo = QComboBox()
        self.chart_type_combo.addItems([
            "Bar Chart", "Line Chart", "Pie Chart", 
            "Area Chart", "Scatter Chart", "Stacked Bar Chart"
        ])
        self.chart_type_combo.currentTextChanged.connect(self.update_chart)
        chart_type_layout.addWidget(self.chart_type_combo)
        
        chart_type_group.setLayout(chart_type_layout)
        self.left_layout.addWidget(chart_type_group)
        
        # Data options
        data_options_group = QGroupBox("Data Options")
        data_layout = QGridLayout()
        
        # First row has headers option
        self.has_header_row = QCheckBox("First row contains headers")
        self.has_header_row.setChecked(True)
        self.has_header_row.stateChanged.connect(self.update_chart)
        data_layout.addWidget(self.has_header_row, 0, 0, 1, 2)
        
        # First column has categories option
        self.has_header_col = QCheckBox("First column contains categories")
        self.has_header_col.setChecked(True)
        self.has_header_col.stateChanged.connect(self.update_chart)
        data_layout.addWidget(self.has_header_col, 1, 0, 1, 2)
        
        data_options_group.setLayout(data_layout)
        self.left_layout.addWidget(data_options_group)
        
        # Appearance options
        appearance_group = QGroupBox("Appearance")
        appearance_layout = QGridLayout()
        
        # Title
        title_label = QLabel("Chart Title:")
        self.title_edit = QComboBox()
        self.title_edit.setEditable(True)
        self.title_edit.addItem("Chart Title")
        self.title_edit.currentTextChanged.connect(self.update_chart)
        appearance_layout.addWidget(title_label, 0, 0)
        appearance_layout.addWidget(self.title_edit, 0, 1)
        
        # Legend visibility
        self.show_legend = QCheckBox("Show Legend")
        self.show_legend.setChecked(True)
        self.show_legend.stateChanged.connect(self.update_chart)
        appearance_layout.addWidget(self.show_legend, 1, 0, 1, 2)
        
        # Animation
        self.use_animation = QCheckBox("Use Animation")
        self.use_animation.setChecked(True)
        self.use_animation.stateChanged.connect(self.update_chart)
        appearance_layout.addWidget(self.use_animation, 2, 0, 1, 2)
        
        # Theme
        theme_label = QLabel("Theme:")
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Light", "Dark", "Blue Cerulean", "Brown Sand", "Blue Ncs", "High Contrast"])
        self.theme_combo.currentTextChanged.connect(self.update_chart)
        appearance_layout.addWidget(theme_label, 3, 0)
        appearance_layout.addWidget(self.theme_combo, 3, 1)
        
        appearance_group.setLayout(appearance_layout)
        self.left_layout.addWidget(appearance_group)
        
        # Data preview
        data_preview_group = QGroupBox("Data Preview")
        data_preview_layout = QVBoxLayout()
        
        self.data_table = QTableWidget()
        self.data_table.setMinimumHeight(150)
        data_preview_layout.addWidget(self.data_table)
        
        data_preview_group.setLayout(data_preview_layout)
        self.left_layout.addWidget(data_preview_group)

    def create_chart_panel(self):
        """Create the chart preview panel"""
        # Create chart view
        self.chart = QChart()
        self.chart_view = QChartView(self.chart)
        self.chart_view.setRenderHint(QPainter.Antialiasing)
        
        self.right_layout.addWidget(self.chart_view)

    def create_buttons(self):
        """Create bottom buttons"""
        button_layout = QHBoxLayout()
        
        self.save_button = QPushButton("Save Chart as Image")
        self.save_button.clicked.connect(self.save_chart_image)
        button_layout.addWidget(self.save_button)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        button_layout.addWidget(button_box)
        
        self.layout.addLayout(button_layout)

    def process_data(self):
        """Process the input data for charting"""
        if not self.data:
            return
            
        # Populate data preview table
        self.data_table.setRowCount(len(self.data))
        self.data_table.setColumnCount(len(self.data[0]) if self.data else 0)
        
        # Column headers
        if len(self.data) > 0:
            if self.has_header_row.isChecked():
                self.data_table.setHorizontalHeaderLabels(self.data[0])
            else:
                self.data_table.setHorizontalHeaderLabels([f"Column {i+1}" for i in range(len(self.data[0]))])
                
        # Row headers
        if len(self.data) > 0 and len(self.data[0]) > 0:
            if self.has_header_col.isChecked():
                row_headers = [row[0] for row in self.data[1:]] if self.has_header_row.isChecked() else [row[0] for row in self.data]
                self.data_table.setVerticalHeaderLabels(["Headers"] + row_headers if self.has_header_row.isChecked() else row_headers)
            else:
                self.data_table.setVerticalHeaderLabels([f"Row {i+1}" for i in range(len(self.data))])
        
        # Fill table with data
        for row, row_data in enumerate(self.data):
            for col, value in enumerate(row_data):
                item = QTableWidgetItem(str(value))
                self.data_table.setItem(row, col, item)
        
        # Set some suggested chart titles based on headers
        self.title_edit.clear()
        self.title_edit.addItem("Chart Title")
        
        if self.has_header_row.isChecked() and len(self.data) > 0:
            for header in self.data[0]:
                self.title_edit.addItem(header)

    def update_chart(self):
        """Update the chart based on current options"""
        chart_type = self.chart_type_combo.currentText()
        
        # Clear the current chart
        self.chart.removeAllSeries()
        self.chart.removeAxis(self.chart.axisX())
        self.chart.removeAxis(self.chart.axisY())
        
        # Set chart title
        self.chart.setTitle(self.title_edit.currentText())
        
        # Set chart animation
        if self.use_animation.isChecked():
            self.chart.setAnimationOptions(QChart.SeriesAnimations)
        else:
            self.chart.setAnimationOptions(QChart.NoAnimation)
        
        # Show/hide legend
        self.chart.legend().setVisible(self.show_legend.isChecked())
        
        # Apply theme
        theme_map = {
            "Light": QChart.ChartThemeLight,
            "Dark": QChart.ChartThemeDark,
            "Blue Cerulean": QChart.ChartThemeBlueCerulean,
            "Brown Sand": QChart.ChartThemeBrownSand,
            "Blue Ncs": QChart.ChartThemeBlueNcs,
            "High Contrast": QChart.ChartThemeHighContrast
        }
        selected_theme = self.theme_combo.currentText()
        if selected_theme in theme_map:
            self.chart.setTheme(theme_map[selected_theme])
        
        # Extract data
        if not self.data:
            return
            
        # Determine start indices based on whether we have headers
        data_start_row = 1 if self.has_header_row.isChecked() else 0
        data_start_col = 1 if self.has_header_col.isChecked() else 0
        
        # Get categories (X-axis labels)
        categories = []
        if self.has_header_col.isChecked():
            for row in range(data_start_row, len(self.data)):
                categories.append(str(self.data[row][0]))
        else:
            categories = [f"Item {i+1}" for i in range(len(self.data) - data_start_row)]
        
        # Get series names (from header row if available)
        series_names = []
        if self.has_header_row.isChecked() and len(self.data) > 0:
            for col in range(data_start_col, len(self.data[0])):
                series_names.append(str(self.data[0][col]))
        else:
            series_names = [f"Series {i+1}" for i in range(len(self.data[0]) - data_start_col)]
        
        # Create appropriate chart type
        if chart_type == "Bar Chart":
            self.create_bar_chart(categories, series_names, data_start_row, data_start_col)
        elif chart_type == "Line Chart":
            self.create_line_chart(categories, series_names, data_start_row, data_start_col)
        elif chart_type == "Pie Chart":
            self.create_pie_chart(categories, series_names, data_start_row, data_start_col)
        elif chart_type == "Area Chart":
            self.create_area_chart(categories, series_names, data_start_row, data_start_col)
        elif chart_type == "Scatter Chart":
            self.create_scatter_chart(categories, series_names, data_start_row, data_start_col)
        elif chart_type == "Stacked Bar Chart":
            self.create_stacked_bar_chart(categories, series_names, data_start_row, data_start_col)
        
        # Set legend alignment
        self.chart.legend().setAlignment(Qt.AlignBottom)

    def create_bar_chart(self, categories, series_names, start_row, start_col):
        """Create a bar chart"""
        bar_series = QBarSeries()
        
        # Create a bar set for each data series
        for col in range(len(series_names)):
            bar_set = QBarSet(series_names[col])
            
            # Add data for each category
            for row in range(len(categories)):
                try:
                    value = float(self.data[row + start_row][col + start_col])
                    bar_set.append(value)
                except (ValueError, IndexError):
                    bar_set.append(0)
            
            bar_series.append(bar_set)
        
        self.chart.addSeries(bar_series)
        
        # Create axes
        axis_x = QBarCategoryAxis()
        axis_x.append(categories)
        self.chart.addAxis(axis_x, Qt.AlignBottom)
        bar_series.attachAxis(axis_x)
        
        axis_y = QValueAxis()
        self.chart.addAxis(axis_y, Qt.AlignLeft)
        bar_series.attachAxis(axis_y)
        
        # Adjust Y axis range automatically
        self.adjust_y_axis_range(axis_y)

    def create_line_chart(self, categories, series_names, start_row, start_col):
        """Create a line chart"""
        # Create a line series for each data series
        for col in range(len(series_names)):
            series = QLineSeries()
            series.setName(series_names[col])
            
            # Add data points
            for row in range(len(categories)):
                try:
                    value = float(self.data[row + start_row][col + start_col])
                    series.append(row, value)
                except (ValueError, IndexError):
                    series.append(row, 0)
            
            self.chart.addSeries(series)
        
        # Create axes
        axis_x = QValueAxis()
        axis_x.setRange(0, max(0, len(categories) - 1))
        axis_x.setTickCount(min(10, len(categories)))
        axis_x.setLabelFormat("%.0f")
        self.chart.addAxis(axis_x, Qt.AlignBottom)
        
        axis_y = QValueAxis()
        self.chart.addAxis(axis_y, Qt.AlignLeft)
        
        # Attach all series to the axes
        for series in self.chart.series():
            series.attachAxis(axis_x)
            series.attachAxis(axis_y)
        
        # Adjust Y axis range automatically
        self.adjust_y_axis_range(axis_y)
        
        # Create category labels on X axis
        if categories:
            labels = []
            axis_x.setLabelsVisible(False)
            for i, category in enumerate(categories):
                label = QLabel(category)
                label.setAlignment(Qt.AlignCenter)
                chart_view_pos = self.chart_view.mapFromScene(self.chart.mapToScene(i, 0))
                label.move(chart_view_pos.x() - label.width() / 2, chart_view_pos.y() + 10)
                label.setParent(self.chart_view)
                labels.append(label)

    def create_pie_chart(self, categories, series_names, start_row, start_col):
        """Create a pie chart"""
        # Pie chart only supports a single series, so use the first data column
        series = QPieSeries()
        
        # Add slices for each category
        for row in range(len(categories)):
            try:
                value = float(self.data[row + start_row][start_col])
                slice = series.append(categories[row], value)
                slice.setLabelVisible(True)
            except (ValueError, IndexError):
                pass
        
        self.chart.addSeries(series)
        
        # Legend is particularly useful for pie charts
        series.setLabelsVisible(True)
        
        # Set some properties for better pie visualization
        if selected_theme := self.theme_combo.currentText():
            if selected_theme == "Dark":
                for slice in series.slices():
                    slice.setLabelColor(Qt.white)
        
        # Make chart more accessible
        series.setLabelsPosition(QPieSlice.LabelOutside)
        
        # Make sure pie is properly centered and sized
        self.chart.setBackgroundVisible(False)
        if hasattr(self.chart, 'layout') and callable(self.chart.layout):
            self.chart.layout().setContentsMargins(0, 0, 0, 0)
        self.chart.setMargins(QMargins(0, 0, 0, 0))

    def create_area_chart(self, categories, series_names, start_row, start_col):
        """Create an area chart"""
        # Create a QLineSeries for each data series
        upperSeries = []
        
        for col in range(len(series_names)):
            upper = QLineSeries()
            upper.setName(series_names[col])
            
            # Add data points
            for row in range(len(categories)):
                try:
                    value = float(self.data[row + start_row][col + start_col])
                    upper.append(row, value)
                except (ValueError, IndexError):
                    upper.append(row, 0)
            
            # Create area series
            lower = QLineSeries()
            for i in range(len(categories)):
                lower.append(i, 0)  # Base line at y=0
            
            areaSeries = QAreaSeries(upper, lower)
            areaSeries.setName(series_names[col])
            self.chart.addSeries(areaSeries)
            
            upperSeries.append(upper)
        
        # Create axes
        axis_x = QValueAxis()
        axis_x.setRange(0, max(0, len(categories) - 1))
        axis_x.setTickCount(min(10, len(categories)))
        axis_x.setLabelFormat("%.0f")
        self.chart.addAxis(axis_x, Qt.AlignBottom)
        
        axis_y = QValueAxis()
        self.chart.addAxis(axis_y, Qt.AlignLeft)
        
        # Attach all series to the axes
        for series in self.chart.series():
            series.attachAxis(axis_x)
            series.attachAxis(axis_y)
        
        # Adjust Y axis range automatically
        self.adjust_y_axis_range(axis_y)

    def create_scatter_chart(self, categories, series_names, start_row, start_col):
        """Create a scatter chart"""
        # For scatter chart, we'll use the first two columns as X and Y coordinates
        if len(self.data[0]) < start_col + 2:
            return  # Need at least two data columns
            
        series = QScatterSeries()
        series.setName("Data Points")
        series.setMarkerSize(10)
        
        for row in range(len(categories)):
            try:
                x_value = float(self.data[row + start_row][start_col])
                y_value = float(self.data[row + start_row][start_col + 1])
                series.append(x_value, y_value)
            except (ValueError, IndexError):
                pass
        
        self.chart.addSeries(series)
        
        # Create axes
        axis_x = QValueAxis()
        self.chart.addAxis(axis_x, Qt.AlignBottom)
        series.attachAxis(axis_x)
        
        axis_y = QValueAxis()
        self.chart.addAxis(axis_y, Qt.AlignLeft)
        series.attachAxis(axis_y)
        
        # Adjust axis ranges automatically
        self.adjust_axis_ranges(axis_x, axis_y, series)

    def create_stacked_bar_chart(self, categories, series_names, start_row, start_col):
        """Create a stacked bar chart"""
        bar_series = QStackedBarSeries()
        
        # Create a bar set for each data series
        for col in range(len(series_names)):
            bar_set = QBarSet(series_names[col])
            
            # Add data for each category
            for row in range(len(categories)):
                try:
                    value = float(self.data[row + start_row][col + start_col])
                    bar_set.append(value)
                except (ValueError, IndexError):
                    bar_set.append(0)
            
            bar_series.append(bar_set)
        
        self.chart.addSeries(bar_series)
        
        # Create axes
        axis_x = QBarCategoryAxis()
        axis_x.append(categories)
        self.chart.addAxis(axis_x, Qt.AlignBottom)
        bar_series.attachAxis(axis_x)
        
        axis_y = QValueAxis()
        self.chart.addAxis(axis_y, Qt.AlignLeft)
        bar_series.attachAxis(axis_y)
        
        # Adjust Y axis range automatically
        self.adjust_y_axis_range(axis_y)

    def adjust_y_axis_range(self, axis):
        """Automatically adjust the Y axis range based on data values"""
        min_val = float('inf')
        max_val = float('-inf')
        
        for series in self.chart.series():
            if hasattr(series, 'barSets'):  # For bar charts
                for bar_set in series.barSets():
                    for i in range(bar_set.count()):
                        val = bar_set.at(i)
                        min_val = min(min_val, val)
                        max_val = max(max_val, val)
            elif hasattr(series, 'slices'):  # For pie charts
                for slice in series.slices():
                    val = slice.value()
                    min_val = min(min_val, val)
                    max_val = max(max_val, val)
                    
        # Handle cases where we couldn't determine min/max
        if min_val == float('inf'):
            min_val = 0
        if max_val == float('-inf'):
            max_val = 10
            
        # Add some padding
        range_padding = (max_val - min_val) * 0.1
        if range_padding == 0:
            range_padding = 1  # In case all values are the same
            
        axis.setRange(max(0, min_val - range_padding), max_val + range_padding)

    def adjust_axis_ranges(self, x_axis, y_axis, series):
        """Automatically adjust both X and Y axis ranges for scatter charts"""
        min_x = float('inf')
        max_x = float('-inf')
        min_y = float('inf')
        max_y = float('-inf')
        
        # Find min and max values
        for point in series.points():
            min_x = min(min_x, point.x())
            max_x = max(max_x, point.x())
            min_y = min(min_y, point.y())
            max_y = max(max_y, point.y())
            
        # Handle cases where we couldn't determine min/max
        if min_x == float('inf'):
            min_x = 0
        if max_x == float('-inf'):
            max_x = 10
        if min_y == float('inf'):
            min_y = 0
        if max_y == float('-inf'):
            max_y = 10
            
        # Add padding
        x_padding = (max_x - min_x) * 0.1 or 1
        y_padding = (max_y - min_y) * 0.1 or 1
        
        x_axis.setRange(min_x - x_padding, max_x + x_padding)
        y_axis.setRange(min_y - y_padding, max_y + y_padding)

    def save_chart_image(self):
        """Save the chart as an image file"""
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(
            self, "Save Chart", "", 
            "PNG Files (*.png);;JPEG Files (*.jpg *.jpeg);;All Files (*)",
            options=options
        )
        
        if file_name:
            # Create a pixmap from the chart view
            pixmap = QPixmap(self.chart_view.size())
            pixmap.fill(Qt.white)
            
            # Render the chart onto the pixmap
            painter = QPainter(pixmap)
            self.chart_view.render(painter)
            painter.end()
            
            # Save the pixmap
            pixmap.save(file_name)
