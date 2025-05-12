import matplotlib.pyplot as plt
import numpy as np
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
import mplfinance as mpf
from datetime import datetime
import pandas as pd

class ChartEngine:
    def __init__(self):
        """Initialize charting engine"""
        self.supported_chart_types = [
            'line', 'bar', 'scatter', 'pie', 'area', 'bubble', 'stock'
        ]
        self.supported_export_formats = [
            'png', 'jpg', 'pdf', 'svg', 'eps'
        ]
    
    def create_chart(self, chart_type, data, headers=None, chart_title=None):
        """Create a chart based on the specified type and data.
        
        Args:
            chart_type: String specifying the chart type
            data: 2D list or numpy array of data
            headers: List of column headers
            chart_title: Optional title for the chart
            
        Returns:
            Matplotlib figure object
        """
        if chart_type not in self.supported_chart_types:
            raise ValueError(f"Unsupported chart type: {chart_type}")
        
        # Convert data to numpy array for easier manipulation
        if isinstance(data, list):
            data_array = np.array(data)
        else:
            data_array = data
            
        # Create figure
        fig, ax = plt.subplots(figsize=(10, 6))
        
        # Apply headers or generate default column names
        if not headers:
            headers = [f"Column {i+1}" for i in range(data_array.shape[1])]
        
        # Apply chart title
        if chart_title:
            ax.set_title(chart_title)
        else:
            ax.set_title(f"{chart_type.capitalize()} Chart")
        
        # Create appropriate chart based on type
        if chart_type == 'line':
            self._create_line_chart(ax, data_array, headers)
        elif chart_type == 'bar':
            self._create_bar_chart(ax, data_array, headers)
        elif chart_type == 'scatter':
            self._create_scatter_chart(ax, data_array, headers)
        elif chart_type == 'pie':
            self._create_pie_chart(ax, data_array, headers)
        elif chart_type == 'area':
            self._create_area_chart(ax, data_array, headers)
        elif chart_type == 'bubble':
            self._create_bubble_chart(ax, data_array, headers)
        elif chart_type == 'stock':
            return self._create_stock_chart(data_array, headers, chart_title)
            
        plt.tight_layout()
        return fig
        
    def _create_line_chart(self, ax, data_array, headers):
        """Create a line chart"""
        # Use first column as x-axis if available, otherwise use indices
        if data_array.shape[1] > 1:
            for i in range(1, data_array.shape[1]):
                ax.plot(data_array[:, 0], data_array[:, i], 
                        label=headers[i] if i < len(headers) else f"Series {i}")
            ax.set_xlabel(headers[0] if headers else "X")
        else:
            ax.plot(data_array[:, 0], label=headers[0] if headers else "Series 1")
            ax.set_xlabel("Index")
            
        ax.set_ylabel("Value")
        ax.grid(True, linestyle='--', alpha=0.7)
        ax.legend()
        
    def _create_bar_chart(self, ax, data_array, headers):
        """Create a bar chart"""
        # Use first column as categories if available
        if data_array.shape[1] > 1:
            x = np.arange(len(data_array))
            width = 0.8 / (data_array.shape[1] - 1)
            
            for i in range(1, data_array.shape[1]):
                ax.bar(x + width * (i - 1.5), data_array[:, i], width,
                       label=headers[i] if i < len(headers) else f"Series {i}")
                
            ax.set_xticks(x)
            ax.set_xticklabels(data_array[:, 0])
            ax.set_xlabel(headers[0] if headers else "Categories")
        else:
            ax.bar(np.arange(len(data_array)), data_array[:, 0])
            ax.set_xlabel("Index")
            
        ax.set_ylabel("Value")
        ax.grid(True, linestyle='--', alpha=0.7, axis='y')
        ax.legend()
        
    def _create_scatter_chart(self, ax, data_array, headers):
        """Create a scatter chart"""
        if data_array.shape[1] >= 2:
            ax.scatter(data_array[:, 0], data_array[:, 1])
            ax.set_xlabel(headers[0] if headers else "X")
            ax.set_ylabel(headers[1] if len(headers) > 1 else "Y")
        else:
            ax.scatter(np.arange(len(data_array)), data_array[:, 0])
            ax.set_xlabel("Index")
            ax.set_ylabel(headers[0] if headers else "Value")
            
        ax.grid(True, linestyle='--', alpha=0.7)
        
    def _create_pie_chart(self, ax, data_array, headers):
        """Create a pie chart"""
        # Use only the first two columns (labels and values) or just values
        if data_array.shape[1] >= 2:
            ax.pie(data_array[:, 1], labels=data_array[:, 0], autopct='%1.1f%%', 
                   startangle=90)
        else:
            ax.pie(data_array[:, 0], autopct='%1.1f%%', startangle=90)
            
        ax.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle
        
    def _create_area_chart(self, ax, data_array, headers):
        """Create an area chart"""
        # Use first column as x-axis if available
        x = np.arange(len(data_array)) if data_array.shape[1] == 1 else data_array[:, 0]
        
        if data_array.shape[1] > 1:
            for i in range(1, data_array.shape[1]):
                ax.fill_between(x, data_array[:, i], alpha=0.3)
                ax.plot(x, data_array[:, i], 
                        label=headers[i] if i < len(headers) else f"Series {i}")
            ax.set_xlabel(headers[0] if headers else "X")
        else:
            ax.fill_between(x, data_array[:, 0], alpha=0.3)
            ax.plot(x, data_array[:, 0], label=headers[0] if headers else "Series 1")
            ax.set_xlabel("Index")
            
        ax.set_ylabel("Value")
        ax.grid(True, linestyle='--', alpha=0.7)
        ax.legend()
        
    def _create_bubble_chart(self, ax, data_array, headers):
        """Create a bubble chart"""
        if data_array.shape[1] >= 3:
            # Use first column for X, second for Y, third for bubble size
            scatter = ax.scatter(data_array[:, 0], data_array[:, 1], 
                                s=data_array[:, 2] * 20,  # Scale bubble sizes
                                alpha=0.5)
                                
            ax.set_xlabel(headers[0] if headers else "X")
            ax.set_ylabel(headers[1] if len(headers) > 1 else "Y")
            
            # Add a colorbar legend for bubble sizes
            plt.colorbar(scatter, ax=ax, label=headers[2] if len(headers) > 2 else "Size")
            
        else:
            raise ValueError("Bubble chart requires at least 3 columns of data (X, Y, Size)")
            
        ax.grid(True, linestyle='--', alpha=0.7)
        
    def _create_stock_chart(self, data_array, headers, chart_title):
        """Create a stock (OHLC) chart"""
        if data_array.shape[1] < 4:
            raise ValueError("Stock chart requires at least 4 columns (Date, Open, High, Low, Close)")
        
        # Check if first column has dates, if not create dummy dates
        try:
            dates = pd.to_datetime(data_array[:, 0])
        except:
            dates = pd.date_range(start='2020-01-01', periods=len(data_array))
            
        # Create a pandas DataFrame with OHLC data
        df = pd.DataFrame({
            'Date': dates,
            'Open': data_array[:, 1],
            'High': data_array[:, 2],
            'Low': data_array[:, 3],
            'Close': data_array[:, 4],
            'Volume': data_array[:, 5] if data_array.shape[1] > 5 else np.zeros(len(data_array))
        })
        df.set_index('Date', inplace=True)
        
        # Create the OHLC chart using mplfinance
        fig, ax = mpf.plot(df, type='candle', style='yahoo',
                         title=chart_title or "Stock Price Chart",
                         ylabel='Price',
                         volume=data_array.shape[1] > 5,  # Show volume panel if available
                         returnfig=True)
                         
        return fig[0]  # Return the figure object
    
    def save_chart(self, fig, filename, format=None):
        """Save a chart to a file.
        
        Args:
            fig: Matplotlib figure object
            filename: Output file path
            format: Optional file format (will be inferred from filename if not provided)
            
        Returns:
            Boolean indicating success
        """
        if not format:
            # Infer format from filename
            ext = filename.split('.')[-1].lower()
            if ext not in self.supported_export_formats:
                raise ValueError(f"Unsupported file format: {ext}")
            format = ext
        
        try:
            fig.savefig(filename, format=format, dpi=300, bbox_inches='tight')
            return True
        except Exception as e:
            print(f"Error saving chart: {e}")
            return False
    
    def get_chart_widget(self, fig):
        """Create a QWidget containing the chart for embedding in Qt applications.
        
        Args:
            fig: Matplotlib figure object
            
        Returns:
            FigureCanvasQTAgg widget
        """
        return FigureCanvasQTAgg(fig)
