#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
File utility functions for file operations
"""

from src.models.file_model import FileModel
import os
import shutil
from pathlib import Path
from typing import List, Dict, Any
import sys

# Add the parent directory to sys.path to allow relative imports
sys.path.insert(0, os.path.abspath(
    os.path.join(os.path.dirname(__file__), '../..')))


def scan_directory(directory_path: str) -> List[FileModel]:
    """
    Scan a directory and return a list of FileModel objects

    Args:
        directory_path: Path to the directory to scan

    Returns:
        List of FileModel objects
    """
    path = Path(directory_path)
    if not path.exists() or not path.is_dir():
        return []

    files = []
    for file_path in path.glob("*"):
        if file_path.is_file():
            files.append(FileModel(file_path))

    # Sort files by name
    files.sort(key=lambda x: x.name.lower())

    return files


def filter_files_by_extension(files: List[FileModel], extensions: List[str] = None) -> List[FileModel]:
    """
    Filter files by extension

    Args:
        files: List of FileModel objects
        extensions: List of extensions to include (e.g., ['.pdf', '.docx'])
                   If None, all files will be included

    Returns:
        Filtered list of FileModel objects
    """
    if extensions is None:
        return files

    return [file for file in files
            if any(file.extension.lower() == ext.lower() for ext in extensions)]


def search_files_by_name(files: List[FileModel], search_term: str) -> List[FileModel]:
    """
    Search files by name

    Args:
        files: List of FileModel objects
        search_term: Search term to look for in filenames

    Returns:
        Filtered list of FileModel objects
    """
    if not search_term:
        return files

    search_term = search_term.lower()
    return [file for file in files if search_term in file.name.lower()]


def rename_file(file_model: FileModel, new_name: str, keep_extension: bool = True) -> bool:
    """
    Rename a file

    Args:
        file_model: FileModel object
        new_name: New filename
        keep_extension: Whether to keep the original extension

    Returns:
        True if rename was successful, False otherwise
    """
    try:
        source_path = file_model.path
        parent_dir = source_path.parent

        # Add extension if needed
        if keep_extension and not new_name.lower().endswith(file_model.extension.lower()):
            new_name = f"{new_name}{file_model.extension}"

        dest_path = parent_dir / new_name

        # Check if destination file already exists
        if dest_path.exists():
            # In a real application, you might want to handle this case differently
            return False

        # Rename file
        os.rename(source_path, dest_path)

        # Update file model
        file_model.path = dest_path
        file_model.name = new_name

        return True
    except Exception as e:
        print(f"Error renaming file: {str(e)}")
        return False
