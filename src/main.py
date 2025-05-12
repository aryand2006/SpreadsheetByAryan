import sys
import os
import logging
from datetime import datetime

# Add the project root to the Python path to enable absolute imports
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.gui.main_window import MainWindow
from src.engine.calculator import Calculator
from src.engine.data_analysis import RegressionAnalysis, DescriptiveStatistics, TimeSeriesForecasting
from src.engine.chart_engine import ChartEngine
from src.engine.version_control import VersionControl
from PyQt5.QtWidgets import QApplication, QSplashScreen
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt, QTimer

# Configure logging
def configure_logging():
    log_dir = os.path.join(os.path.dirname(__file__), '..', 'logs')
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, f'pyspreadsheet_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger('PySpreadsheet')

def main():
    # Set up logging
    logger = configure_logging()
    logger.info("Application starting")
    
    # Initialize the application
    app = QApplication(sys.argv)
    
    # Show splash screen
    splash_path = os.path.join(os.path.dirname(__file__), '..', 'resources', 'splash.png')
    try:
        splash_pix = QPixmap(splash_path)
        splash = QSplashScreen(splash_pix)
        splash.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        splash.show()
        app.processEvents()
    except:
        logger.warning("Splash screen could not be loaded")
        splash = None
    
    # Initialize components with status updates on splash screen
    def update_splash(message):
        if splash:
            splash.showMessage(message, Qt.AlignBottom | Qt.AlignCenter, Qt.white)
            app.processEvents()
    
    # Initialize formula evaluation system
    update_splash("Initializing formula engine...")
    calculator = Calculator()
    logger.info("Formula engine initialized")
    
    # Initialize data analysis tools
    update_splash("Loading data analysis tools...")
    regression_analyzer = RegressionAnalysis()
    stats_analyzer = DescriptiveStatistics()
    forecaster = TimeSeriesForecasting()
    logger.info("Data analysis tools loaded")
    
    # Initialize charting engine
    update_splash("Setting up chart engine...")
    chart_engine = ChartEngine()
    logger.info("Chart engine ready")
    
    # Initialize version control system
    update_splash("Configuring version control...")
    version_control = VersionControl()
    logger.info("Version control system initialized")
    
    # Create main window
    update_splash("Starting application...")
    window = MainWindow()
    
    # Inject dependencies
    window.calculator = calculator
    window.regression_analyzer = regression_analyzer
    window.stats_analyzer = stats_analyzer
    window.forecaster = forecaster
    window.chart_engine = chart_engine
    window.version_control = version_control
    
    # Close splash and show main window
    if splash:
        QTimer.singleShot(1000, splash.close)
    
    window.show()
    logger.info("Application started successfully")
    
    # Run the application
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()