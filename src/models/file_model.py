#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
File model for representing file information
"""

from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Optional


class FileModel:
    """Model for representing file information"""

    def __init__(self, file_path: Path):
        """
        Initialize a file model from a Path object

        Args:
            file_path: Path object pointing to the file
        """
        self.path = file_path
        self.name = file_path.name
        self.extension = file_path.suffix.lower()
        self.filename_without_ext = file_path.stem

        # Get file stats
        file_stat = file_path.stat()
        self.size = file_stat.st_size
        self.mod_time = file_stat.st_mtime
        self.mod_time_formatted = datetime.fromtimestamp(
            file_stat.st_mtime).strftime('%Y-%m-%d %H:%M')

    def to_dict(self) -> Dict[str, Any]:
        """
        Convert file model to a dictionary

        Returns:
            Dictionary representation of the file
        """
        return {
            "path": self.path,
            "name": self.name,
            "size": self.size,
            "mod_time": self.mod_time,
            "mod_time_formatted": self.mod_time_formatted,
            "extension": self.extension
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'FileModel':
        """
        Create a FileModel instance from a dictionary

        Args:
            data: Dictionary containing file information

        Returns:
            FileModel instance
        """
        file_model = cls(data["path"])
        file_model.size = data["size"]
        file_model.mod_time = data["mod_time"]
        file_model.mod_time_formatted = data["mod_time_formatted"]
        return file_model

    def format_size(self) -> str:
        """
        Format file size in human-readable format

        Returns:
            Formatted size string
        """
        size_bytes = self.size
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.2f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.2f} TB"
