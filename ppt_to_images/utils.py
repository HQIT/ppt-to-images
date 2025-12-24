"""
Utility functions for ppt-to-images.
"""

import os
import base64
import tempfile
import shutil
from pathlib import Path
from typing import Union, Optional, BinaryIO
from enum import Enum


class FileType(Enum):
    """Supported file types."""
    PPT = "ppt"
    PPTX = "pptx"
    PDF = "pdf"
    UNKNOWN = "unknown"


def detect_file_type(file_path: Union[str, Path]) -> FileType:
    """
    Detect file type based on extension.
    
    Args:
        file_path: Path to the file
        
    Returns:
        FileType enum value
    """
    path = Path(file_path)
    ext = path.suffix.lower()
    
    if ext == ".ppt":
        return FileType.PPT
    elif ext == ".pptx":
        return FileType.PPTX
    elif ext == ".pdf":
        return FileType.PDF
    else:
        return FileType.UNKNOWN


def image_to_base64(image_path: Union[str, Path]) -> str:
    """
    Convert image file to base64 string.
    
    Args:
        image_path: Path to the image file
        
    Returns:
        Base64 encoded string
    """
    with open(image_path, 'rb') as f:
        return base64.b64encode(f.read()).decode('utf-8')


def base64_to_image(base64_str: str, output_path: Union[str, Path]) -> None:
    """
    Convert base64 string to image file.
    
    Args:
        base64_str: Base64 encoded string
        output_path: Path to save the image
    """
    with open(output_path, 'wb') as f:
        f.write(base64.b64decode(base64_str))


class TempFileManager:
    """
    Context manager for temporary files and directories.
    """
    
    def __init__(self, cleanup: bool = True, temp_dir: Optional[str] = None):
        """
        Initialize temporary file manager.
        
        Args:
            cleanup: Whether to clean up temp files on exit
            temp_dir: Custom temporary directory
        """
        self.cleanup = cleanup
        self.temp_dir = temp_dir
        self._created_dir = None
        self._temp_files = []
    
    def __enter__(self):
        if self.temp_dir:
            os.makedirs(self.temp_dir, exist_ok=True)
            self._created_dir = self.temp_dir
        else:
            self._created_dir = tempfile.mkdtemp(prefix="ppt_to_images_")
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.cleanup and self._created_dir:
            try:
                shutil.rmtree(self._created_dir)
            except Exception:
                pass
    
    def get_temp_path(self, filename: str) -> str:
        """
        Get a temporary file path.
        
        Args:
            filename: Name of the temporary file
            
        Returns:
            Full path to the temporary file
        """
        return os.path.join(self._created_dir, filename)
    
    def save_uploaded_file(self, content: bytes, filename: str) -> str:
        """
        Save uploaded file content to temp directory.
        
        Args:
            content: File content as bytes
            filename: Original filename
            
        Returns:
            Path to saved temporary file
        """
        temp_path = self.get_temp_path(filename)
        with open(temp_path, 'wb') as f:
            f.write(content)
        self._temp_files.append(temp_path)
        return temp_path


def ensure_directory(path: Union[str, Path]) -> Path:
    """
    Ensure a directory exists, create if not.
    
    Args:
        path: Directory path
        
    Returns:
        Path object
    """
    path = Path(path)
    path.mkdir(parents=True, exist_ok=True)
    return path

