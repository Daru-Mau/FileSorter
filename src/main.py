#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Excel-Based File Renamer

A GUI application for renaming files based on matching data in Excel spreadsheets.
"""

import tkinter as tk
import sys
import os

# Add the parent directory to sys.path to allow relative imports
sys.path.insert(0, os.path.abspath(
    os.path.join(os.path.dirname(__file__), '..')))

from src.ui.main_window import MainWindow


def main():
    """Main entry point for the application"""
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()


if __name__ == "__main__":
    main()
