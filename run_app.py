#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Excel-Based File Renamer - Entry Point

This is the main entry point for the application that correctly sets up the Python import path.
"""

import os
import sys

# Add the src directory to the Python path
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

# Import and run the application
from src.main import main

if __name__ == "__main__":
    main()
