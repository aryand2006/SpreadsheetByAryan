import os
import json
import shutil
import hashlib
import datetime
import zipfile
from pathlib import Path
import tempfile
import uuid

class VersionControl:
    def __init__(self):
        """Initialize version control system"""
        self.app_data_dir = self._get_app_data_dir()
        self.versions_dir = os.path.join(self.app_data_dir, 'versions')
        self._ensure_directories_exist()
        
    def _get_app_data_dir(self):
        """Get the application data directory based on OS"""
        home = str(Path.home())
        
        if os.name == 'nt':  # Windows
            return os.path.join(home, 'AppData', 'Local', 'PySpreadsheet')
        elif os.name == 'posix':  # macOS or Linux
            if os.path.exists(os.path.join(home, 'Library')):  # macOS
                return os.path.join(home, 'Library', 'Application Support', 'PySpreadsheet')
            else:  # Linux
                return os.path.join(home, '.config', 'pyspreadsheet')
        else:
            return os.path.join(home, '.pyspreadsheet')
    
    def _ensure_directories_exist(self):
        """Ensure the required directories exist"""
        os.makedirs(self.versions_dir, exist_ok=True)
        
    def _get_document_versions_dir(self, document_id):
        """Get the directory for a specific document's versions"""
        doc_dir = os.path.join(self.versions_dir, document_id)
        os.makedirs(doc_dir, exist_ok=True)
        return doc_dir
    
    def _generate_document_id(self, filepath):
        """Generate a unique document ID based on the filepath"""
        return hashlib.md5(filepath.encode('utf-8')).hexdigest()
    
    def _get_version_data_path(self, document_id):
        """Get the path to the versions metadata file"""
        return os.path.join(self._get_document_versions_dir(document_id), 'versions.json')
    
    def save_version(self, filepath, data, comment=''):
        """Save a new version of the document
        
        Args:
            filepath: Path to the original document
            data: Document data to save (workbook dictionary)
            comment: Optional comment for this version
            
        Returns:
            Dictionary with version info
        """
        try:
            # Generate document ID
            document_id = self._generate_document_id(filepath)
            doc_versions_dir = self._get_document_versions_dir(document_id)
            
            # Create version metadata
            timestamp = datetime.datetime.now().isoformat()
            version_id = str(uuid.uuid4())
            version_file = os.path.join(doc_versions_dir, f"{version_id}.zip")
            
            # Save document data
            with tempfile.NamedTemporaryFile(mode='w', delete=False) as temp_file:
                json.dump(data, temp_file)
                temp_path = temp_file.name
                
            # Create a zip archive of the document data
            with zipfile.ZipFile(version_file, 'w', compression=zipfile.ZIP_DEFLATED) as zipf:
                zipf.write(temp_path, arcname='document.json')
                
            # Clean up temp file
            os.unlink(temp_path)
            
            # Update versions metadata
            version_info = {
                'version_id': version_id,
                'timestamp': timestamp,
                'comment': comment,
                'filepath': filepath,
                'size': os.path.getsize(version_file)
            }
            
            versions_data = self.get_versions(document_id)
            versions_data.append(version_info)
            
            with open(self._get_version_data_path(document_id), 'w') as f:
                json.dump(versions_data, f, indent=2)
                
            return version_info
            
        except Exception as e:
            print(f"Error saving version: {e}")
            return None
    
    def get_versions(self, document_id):
        """Get all versions for a document
        
        Args:
            document_id: Document ID
            
        Returns:
            List of version metadata dictionaries
        """
        versions_file = self._get_version_data_path(document_id)
        if os.path.exists(versions_file):
            try:
                with open(versions_file, 'r') as f:
                    return json.load(f)
            except:
                return []
        else:
            return []
    
    def get_document_id_from_path(self, filepath):
        """Get a document ID from a filepath
        
        Args:
            filepath: Path to the document
            
        Returns:
            Document ID string
        """
        return self._generate_document_id(filepath)
    
    def load_version(self, document_id, version_id):
        """Load a specific version of a document
        
        Args:
            document_id: Document ID
            version_id: Version ID
            
        Returns:
            Document data dictionary
        """
        try:
            doc_versions_dir = self._get_document_versions_dir(document_id)
            version_file = os.path.join(doc_versions_dir, f"{version_id}.zip")
            
            if not os.path.exists(version_file):
                return None
                
            # Extract the document data from the zip archive
            with tempfile.TemporaryDirectory() as temp_dir:
                with zipfile.ZipFile(version_file, 'r') as zipf:
                    zipf.extractall(temp_dir)
                    
                doc_file = os.path.join(temp_dir, 'document.json')
                if os.path.exists(doc_file):
                    with open(doc_file, 'r') as f:
                        return json.load(f)
                        
            return None
            
        except Exception as e:
            print(f"Error loading version: {e}")
            return None
    
    def delete_version(self, document_id, version_id):
        """Delete a specific version of a document
        
        Args:
            document_id: Document ID
            version_id: Version ID
            
        Returns:
            Boolean indicating success
        """
        try:
            doc_versions_dir = self._get_document_versions_dir(document_id)
            version_file = os.path.join(doc_versions_dir, f"{version_id}.zip")
            
            if os.path.exists(version_file):
                os.remove(version_file)
                
            # Update versions metadata
            versions_data = self.get_versions(document_id)
            versions_data = [v for v in versions_data if v['version_id'] != version_id]
            
            with open(self._get_version_data_path(document_id), 'w') as f:
                json.dump(versions_data, f, indent=2)
                
            return True
            
        except Exception as e:
            print(f"Error deleting version: {e}")
            return False
    
    def restore_version(self, document_id, version_id, target_filepath=None):
        """Restore a document to a specific version
        
        Args:
            document_id: Document ID
            version_id: Version ID
            target_filepath: Path to save the restored version (if None, use original path)
            
        Returns:
            Document data dictionary or None if failed
        """
        data = self.load_version(document_id, version_id)
        if not data:
            return None
            
        # Find original filepath if not provided
        if not target_filepath:
            versions = self.get_versions(document_id)
            for version in versions:
                if version['version_id'] == version_id:
                    target_filepath = version['filepath']
                    break
                    
        if not target_filepath:
            return None
            
        # Save as new version before restoring (preserves history)
        current_data = None
        try:
            # Try to load current file data
            if os.path.exists(target_filepath):
                with open(target_filepath, 'r') as f:
                    current_data = json.load(f)
        except:
            pass
            
        if current_data:
            self.save_version(target_filepath, current_data, 
                            comment='Auto-saved before restoring to previous version')
        
        return data
