#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Excel utility functions for Excel operations
"""

from src.models.excel_model import ExcelModel
import pandas as pd
from typing import List, Dict, Any, Optional, Tuple
import sys
import os

# Add the parent directory to sys.path to allow relative imports
sys.path.insert(0, os.path.abspath(
    os.path.join(os.path.dirname(__file__), '../..')))


def get_extension_map() -> Dict[str, Optional[List[str]]]:
    """
    Get a mapping of filter names to file extensions

    Returns:
        Dictionary mapping filter names to extensions
    """
    return {
        "All Files": None,
        "PDF Files": [".pdf"],
        "Excel Files": [".xlsx", ".xls", ".csv"],
        "Word Files": [".docx", ".doc"],
        "Image Files": [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff"],
        "Text Files": [".txt", ".text", ".md", ".rtf"]
    }


def generate_filename_from_pattern(
    excel_row: Dict[str, Any],
    pattern: str,
    separator: str
) -> Optional[str]:
    """
    Generate a filename from a pattern using Excel data

    Args:
        excel_row: Dictionary containing Excel row data
        pattern: Pattern string with column names
        separator: Separator to use between pattern parts

    Returns:
        Generated filename or None if pattern is invalid
    """
    import re

    if not pattern:
        return None

    # Parse the pattern
    pattern_parts = pattern.split(separator)

    # Build the filename from pattern parts
    filename_parts = []
    for part in pattern_parts:
        if part in excel_row:
            # Clean the value - remove special characters
            value = str(excel_row[part])
            # Replace invalid filename chars
            value = re.sub(r'[\\/*?:"<>|]', '_', value)
            filename_parts.append(value)
        else:
            print(f"Warning: Column '{part}' not found in Excel data")
            return None

    # Join the parts with the separator
    custom_name = separator.join(filename_parts)

    # Handle empty result
    if not custom_name:
        return None

    return custom_name
