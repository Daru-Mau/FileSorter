#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
String utility functions for string operations
"""

import re
from typing import Optional


def is_valid_filename(filename: str) -> bool:
    """
    Check if a string is a valid filename

    Args:
        filename: String to check

    Returns:
        True if string is a valid filename, False otherwise
    """
    # Check for empty string
    if not filename or filename.isspace():
        return False

    # Check for invalid characters
    if re.search(r'[\\/*?:"<>|]', filename):
        return False

    return True


def sanitize_filename(filename: str) -> str:
    """
    Sanitize a string to be a valid filename

    Args:
        filename: String to sanitize

    Returns:
        Sanitized filename
    """
    # Replace invalid characters with underscore
    return re.sub(r'[\\/*?:"<>|]', '_', filename)
