from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QListWidget, QPushButton,
    QGroupBox, QDialogButtonBox, QListWidgetItem
)
from PyQt5.QtCore import Qt

class PivotTableDialog(QDialog):
    def __init__(self, data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Create Pivot Table")
        self.setMinimumSize(700, 500)
        self.data = data
        self.column_headers = data[0] if data else []
        self.values = []
        self.rows = []
        self.columns = []
        self.filters = []
        
        self.main_layout = QVBoxLayout(self)
        
        self.create_field_section()
        self.create_pivot_area_section()
        
        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        self.main_layout.addWidget(buttons)
    
    def create_field_section(self):
        field_group = QGroupBox("Choose fields for your PivotTable")
        field_layout = QVBoxLayout()
        
        field_label = QLabel("Available Fields:")
        field_layout.addWidget(field_label)
        
        self.field_list = QListWidget()
        for header in self.column_headers:
            item = QListWidgetItem(header)
            item.setFlags(item.flags() | Qt.ItemIsDragEnabled)
            self.field_list.addItem(item)
        
        field_layout.addWidget(self.field_list)
        field_group.setLayout(field_layout)
        self.main_layout.addWidget(field_group)
    
    def create_pivot_area_section(self):
        areas_group = QGroupBox("PivotTable Areas")
        areas_layout = QHBoxLayout()
        
        # Filters area
        filters_layout = QVBoxLayout()
        filters_label = QLabel("FILTERS:")
        self.filters_list = QListWidget()
        self.filters_list.setAcceptDrops(True)
        filters_layout.addWidget(filters_label)
        filters_layout.addWidget(self.filters_list)
        areas_layout.addLayout(filters_layout)
        
        # Columns area
        columns_layout = QVBoxLayout()
        columns_label = QLabel("COLUMNS:")
        self.columns_list = QListWidget()
        self.columns_list.setAcceptDrops(True)
        columns_layout.addWidget(columns_label)
        columns_layout.addWidget(self.columns_list)
        areas_layout.addLayout(columns_layout)
        
        # Rows and Values areas
        row_values_layout = QVBoxLayout()
        
        # Rows area
        rows_label = QLabel("ROWS:")
        self.rows_list = QListWidget()
        self.rows_list.setAcceptDrops(True)
        row_values_layout.addWidget(rows_label)
        row_values_layout.addWidget(self.rows_list)
        
        # Values area
        values_label = QLabel("VALUES:")
        self.values_list = QListWidget()
        self.values_list.setAcceptDrops(True)
        row_values_layout.addWidget(values_label)
        row_values_layout.addWidget(self.values_list)
        
        areas_layout.addLayout(row_values_layout)
        
        # Add buttons for moving fields
        buttons_layout = QVBoxLayout()
        
        add_to_filters = QPushButton("Add to Filters")
        add_to_filters.clicked.connect(self.add_to_filters)
        buttons_layout.addWidget(add_to_filters)
        
        add_to_columns = QPushButton("Add to Columns")
        add_to_columns.clicked.connect(self.add_to_columns)
        buttons_layout.addWidget(add_to_columns)
        
        add_to_rows = QPushButton("Add to Rows")
        add_to_rows.clicked.connect(self.add_to_rows)
        buttons_layout.addWidget(add_to_rows)
        
        add_to_values = QPushButton("Add to Values")
        add_to_values.clicked.connect(self.add_to_values)
        buttons_layout.addWidget(add_to_values)
        
        remove_field = QPushButton("Remove Field")
        remove_field.clicked.connect(self.remove_field)
        buttons_layout.addWidget(remove_field)
        
        areas_layout.addLayout(buttons_layout)
        
        areas_group.setLayout(areas_layout)
        self.main_layout.addWidget(areas_group)
    
    def add_to_filters(self):
        self._add_selected_to_list(self.filters_list)
    
    def add_to_columns(self):
        self._add_selected_to_list(self.columns_list)
    
    def add_to_rows(self):
        self._add_selected_to_list(self.rows_list)
    
    def add_to_values(self):
        self._add_selected_to_list(self.values_list, "(Sum of) ")
    
    def _add_selected_to_list(self, target_list, prefix=""):
        selected_items = self.field_list.selectedItems()
        if not selected_items:
            return
            
        for item in selected_items:
            new_item = QListWidgetItem(prefix + item.text())
            target_list.addItem(new_item)
    
    def remove_field(self):
        for list_widget in [self.filters_list, self.columns_list, self.rows_list, self.values_list]:
            selected_items = list_widget.selectedItems()
            for item in selected_items:
                list_widget.takeItem(list_widget.row(item))
    
    def get_settings(self):
        """Get the PivotTable configuration"""
        filters = [self.filters_list.item(i).text() for i in range(self.filters_list.count())]
        columns = [self.columns_list.item(i).text() for i in range(self.columns_list.count())]
        rows = [self.rows_list.item(i).text() for i in range(self.rows_list.count())]
        values = [self.values_list.item(i).text() for i in range(self.values_list.count())]
        
        return {
            'filters': filters,
            'columns': columns,
            'rows': rows,
            'values': values
        }

def generate_pivot_table(data, settings):
    """Generate a pivot table based on the settings"""
    if not data or len(data) <= 1:
        return [["No data"]]
    
    # This is a simplified implementation
    headers = data[0]
    rows_data = data[1:]
    
    # Get indices for the specified fields
    row_indices = [headers.index(field) for field in settings['rows'] if field in headers]
    column_indices = [headers.index(field) for field in settings['columns'] if field in headers]
    value_indices = [headers.index(field.replace("(Sum of) ", "")) 
                    for field in settings['values'] if field.replace("(Sum of) ", "") in headers]
    filter_indices = [headers.index(field) for field in settings['filters'] if field in headers]
    
    # Filter data if filters are specified
    filtered_data = rows_data
    # Filter implementation would go here
    
    # Create unique values for rows and columns
    unique_row_values = {}
    for idx in row_indices:
        unique_row_values[idx] = sorted(set(row[idx] for row in filtered_data))
    
    unique_column_values = {}
    for idx in column_indices:
        unique_column_values[idx] = sorted(set(row[idx] for row in filtered_data))
    
    # Generate pivot table header row
    pivot_headers = []
    
    # Add row headers
    for idx in row_indices:
        pivot_headers.append(headers[idx])
    
    # Add column headers (can be complex for multiple columns)
    if column_indices:
        for col_val in unique_column_values[column_indices[0]]:
            for val_idx in value_indices:
                pivot_headers.append(f"{col_val} - {headers[val_idx]}")
    else:
        # Just add value headers if no column fields
        for val_idx in value_indices:
            pivot_headers.append(headers[val_idx])
    
    pivot_table = [pivot_headers]
    
    # Generate data rows
    # This is a simplified approach - for a real pivot table, you'd need
    # a more sophisticated algorithm to handle multiple row/column fields
    if row_indices:
        for row_val in unique_row_values[row_indices[0]]:
            # Filter data for this row value
            row_data = [row for row in filtered_data if row[row_indices[0]] == row_val]
            
            pivot_row = [row_val]
            
            # Calculate values for each column/value combination
            if column_indices:
                for col_val in unique_column_values[column_indices[0]]:
                    # Filter further by column value
                    col_data = [row for row in row_data if row[column_indices[0]] == col_val]
                    
                    # Calculate sum for each value field
                    for val_idx in value_indices:
                        try:
                            # Sum numeric values
                            total = sum(float(row[val_idx]) for row in col_data 
                                      if row[val_idx] and row[val_idx].replace('.','').isdigit())
                            pivot_row.append(str(total))
                        except:
                            pivot_row.append("0")
            else:
                # No column fields, just calculate values directly
                for val_idx in value_indices:
                    try:
                        total = sum(float(row[val_idx]) for row in row_data 
                                  if row[val_idx] and row[val_idx].replace('.','').isdigit())
                        pivot_row.append(str(total))
                    except:
                        pivot_row.append("0")
            
            pivot_table.append(pivot_row)
    
    return pivot_table
