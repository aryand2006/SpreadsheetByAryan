import numpy as np
import pandas as pd
import scipy.stats as stats
from sklearn.linear_model import LinearRegression
import matplotlib.pyplot as plt
from datetime import datetime

class RegressionAnalysis:
    def __init__(self):
        """Initialize regression analysis tools"""
        pass
        
    def linear_regression(self, x_data, y_data):
        """Perform linear regression and return statistics.
        
        Args:
            x_data: List or array of independent variable values
            y_data: List or array of dependent variable values
            
        Returns:
            Dictionary with regression statistics and coefficients
        """
        if len(x_data) != len(y_data) or len(x_data) < 2:
            return {'error': 'Insufficient or mismatched data'}
            
        try:
            # Reshape X for sklearn
            X = np.array(x_data).reshape(-1, 1)
            y = np.array(y_data)
            
            # Perform linear regression
            model = LinearRegression()
            model.fit(X, y)
            
            # Calculate predictions
            y_pred = model.predict(X)
            
            # Calculate statistics
            slope = model.coef_[0]
            intercept = model.intercept_
            
            # R-squared
            r_squared = model.score(X, y)
            
            # Other statistics using scipy.stats.linregress
            slope_alt, intercept_alt, r_value, p_value, std_err = stats.linregress(x_data, y_data)
            
            return {
                'slope': slope,
                'intercept': intercept,
                'r_squared': r_squared,
                'r_value': r_value,
                'p_value': p_value,
                'std_err': std_err,
                'predictions': y_pred.tolist(),
                'equation': f"y = {slope:.6f}x + {intercept:.6f}"
            }
        except Exception as e:
            return {'error': str(e)}
    
    def generate_regression_chart(self, x_data, y_data, results):
        """Generate a chart showing data points and regression line.
        
        Args:
            x_data: List or array of independent variable values
            y_data: List or array of dependent variable values
            results: Dictionary with regression results
            
        Returns:
            Matplotlib figure object
        """
        try:
            fig, ax = plt.subplots(figsize=(10, 6))
            
            # Plot the scatter points
            ax.scatter(x_data, y_data, color='blue', label='Data points')
            
            # Plot the regression line
            x_line = np.linspace(min(x_data), max(x_data), 100)
            y_line = results['slope'] * x_line + results['intercept']
            ax.plot(x_line, y_line, color='red', label=f"Regression line: {results['equation']}")
            
            # Add correlation coefficient to the chart
            ax.text(0.05, 0.95, f"R² = {results['r_squared']:.4f}", transform=ax.transAxes,
                    verticalalignment='top')
            
            ax.set_xlabel('X variable')
            ax.set_ylabel('Y variable')
            ax.set_title('Linear Regression Analysis')
            ax.grid(True, linestyle='--', alpha=0.7)
            ax.legend()
            
            plt.tight_layout()
            return fig
            
        except Exception as e:
            print(f"Error generating regression chart: {e}")
            return None

class DescriptiveStatistics:
    def __init__(self):
        """Initialize descriptive statistics tools"""
        pass
        
    def calculate_statistics(self, data):
        """Calculate descriptive statistics for a dataset.
        
        Args:
            data: List or array of numerical values
            
        Returns:
            Dictionary with statistics
        """
        if not data or len(data) < 1:
            return {'error': 'No data provided'}
            
        try:
            data_array = np.array(data, dtype=float)
            
            # Basic statistics
            stats_dict = {
                'count': len(data_array),
                'mean': np.mean(data_array),
                'median': np.median(data_array),
                'mode': stats.mode(data_array, keepdims=False)[0],
                'std_dev': np.std(data_array, ddof=1),
                'variance': np.var(data_array, ddof=1),
                'min': np.min(data_array),
                'max': np.max(data_array),
                'range': np.max(data_array) - np.min(data_array),
                'sum': np.sum(data_array),
                'quartile_1': np.percentile(data_array, 25),
                'quartile_2': np.percentile(data_array, 50),  # Same as median
                'quartile_3': np.percentile(data_array, 75),
                'iqr': np.percentile(data_array, 75) - np.percentile(data_array, 25)
            }
            
            # Add higher moments
            if len(data_array) > 1:
                stats_dict['skewness'] = stats.skew(data_array)
                stats_dict['kurtosis'] = stats.kurtosis(data_array)
            
            return stats_dict
            
        except Exception as e:
            return {'error': str(e)}
    
    def generate_statistics_chart(self, data, stats):
        """Generate charts for descriptive statistics.
        
        Args:
            data: List or array of numerical values
            stats: Dictionary with statistics
            
        Returns:
            Tuple of Matplotlib figure objects (histogram, boxplot)
        """
        try:
            data_array = np.array(data, dtype=float)
            
            # Create histogram with normal curve overlay
            hist_fig, hist_ax = plt.subplots(figsize=(10, 6))
            
            # Histogram
            hist_ax.hist(data_array, bins='auto', density=True, alpha=0.7, color='skyblue', 
                         edgecolor='black', label='Data Distribution')
            
            # Add a kernel density estimate
            x_min, x_max = hist_ax.get_xlim()
            x = np.linspace(x_min, x_max, 100)
            kde = stats.gaussian_kde(data_array)
            hist_ax.plot(x, kde(x), 'r-', lw=2, label='KDE')
            
            # Create a normal distribution curve
            if 'mean' in stats and 'std_dev' in stats:
                hist_ax.plot(x, stats.norm.pdf(x, stats['mean'], stats['std_dev']), 
                             'k--', lw=1.5, label='Normal Distribution')
            
            # Add annotations for key statistics
            hist_ax.axvline(stats['mean'], color='red', linestyle='-', lw=1.5, label='Mean')
            hist_ax.axvline(stats['median'], color='green', linestyle='--', lw=1.5, label='Median')
            
            hist_ax.set_xlabel('Value')
            hist_ax.set_ylabel('Density')
            hist_ax.set_title('Histogram with Distribution Overlay')
            hist_ax.grid(True, linestyle='--', alpha=0.7)
            hist_ax.legend()
            
            # Create boxplot
            box_fig, box_ax = plt.subplots(figsize=(10, 4))
            box_ax.boxplot(data_array, vert=False, patch_artist=True,
                          boxprops=dict(facecolor='lightblue'))
            
            # Add key statistics to boxplot
            y_pos = 1
            box_ax.axvline(stats['mean'], color='red', linestyle='-', lw=1.5, label='Mean')
            box_ax.text(stats['mean'], y_pos + 0.1, f"Mean: {stats['mean']:.2f}", 
                        ha='center', va='bottom', color='red')
            
            box_ax.set_xlabel('Value')
            box_ax.set_title('Boxplot')
            box_ax.grid(True, linestyle='--', alpha=0.7)
            
            plt.tight_layout()
            return (hist_fig, box_fig)
            
        except Exception as e:
            print(f"Error generating statistics charts: {e}")
            return None

class TimeSeriesForecasting:
    def __init__(self):
        """Initialize time series forecasting tools"""
        pass
        
    def exponential_smoothing(self, data, alpha=0.3, forecast_periods=5):
        """Perform simple exponential smoothing and forecast future values.
        
        Args:
            data: List or array of time series values
            alpha: Smoothing factor (0-1)
            forecast_periods: Number of periods to forecast
            
        Returns:
            Dictionary with forecast results
        """
        if not data or len(data) < 3:
            return {'error': 'Insufficient data for forecasting'}
            
        try:
            data_array = np.array(data, dtype=float)
            
            # Initialize forecast array with first value
            forecast = [data_array[0]]
            
            # Calculate forecast for historical periods
            for i in range(1, len(data_array)):
                forecast.append(alpha * data_array[i] + (1 - alpha) * forecast[i-1])
            
            # Forecast future periods
            future_forecast = []
            last_value = forecast[-1]
            
            for _ in range(forecast_periods):
                last_value = alpha * last_value + (1 - alpha) * last_value  # In practice, this equals last_value
                future_forecast.append(last_value)
            
            return {
                'original_data': data_array.tolist(),
                'historical_forecast': forecast,
                'future_forecast': future_forecast,
                'alpha': alpha,
                'periods': forecast_periods
            }
            
        except Exception as e:
            return {'error': str(e)}
    
    def generate_forecast_chart(self, results):
        """Generate a chart showing original data and forecast.
        
        Args:
            results: Dictionary with forecast results
            
        Returns:
            Matplotlib figure object
        """
        try:
            fig, ax = plt.subplots(figsize=(12, 6))
            
            # Extract data
            original_data = results['original_data']
            historical_forecast = results['historical_forecast']
            future_forecast = results['future_forecast']
            
            # Create date range for x-axis
            all_periods = len(original_data) + len(future_forecast)
            
            # Plot original data
            ax.plot(range(1, len(original_data) + 1), original_data, 'b-', 
                    marker='o', label='Original Data')
            
            # Plot historical forecast
            ax.plot(range(1, len(historical_forecast) + 1), historical_forecast, 'r--', 
                    label='Historical Forecast')
            
            # Plot future forecast
            ax.plot(range(len(original_data) + 1, all_periods + 1), future_forecast, 'r-', 
                    marker='x', label='Future Forecast')
            
            # Add shaded area for future forecast
            ax.axvspan(len(original_data) + 0.5, all_periods + 0.5, 
                       alpha=0.1, color='red', label='Forecast Region')
            
            # Add a title and labels
            ax.set_title(f'Time Series Forecast (α={results["alpha"]})')
            ax.set_xlabel('Period')
            ax.set_ylabel('Value')
            ax.grid(True, linestyle='--', alpha=0.7)
            ax.legend()
            
            plt.tight_layout()
            return fig
            
        except Exception as e:
            print(f"Error generating forecast chart: {e}")
            return None
