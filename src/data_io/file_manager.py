import os

class FileManager:
    """Handle file operations for the spreadsheet"""
    
    def __init__(self):
        self.current_file_path = None
        self.is_modified = False
        
    def open_file(self, file_path):
        """Open a file and update current path"""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
            
        self.current_file_path = file_path
        self.is_modified = False
        return file_path
        
    def save_file(self, file_path, content):
        """Save content to a file and update current path"""
        self.current_file_path = file_path
        self.is_modified = False
        return file_path
        
    def close_file(self):
        """Close the current file"""
        self.current_file_path = None
        self.is_modified = False
        
    def get_current_file_path(self):
        """Get the current file path"""
        return self.current_file_path
        
    def set_modified(self, is_modified=True):
        """Mark the file as modified"""
        self.is_modified = is_modified
        
    def is_file_modified(self):
        """Check if the file has unsaved changes"""
        return self.is_modified