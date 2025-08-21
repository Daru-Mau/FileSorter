#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Pattern builder component for creating filename patterns
"""

from src.utils.excel_utils import generate_filename_from_pattern
import tkinter as tk
from tkinter import ttk, StringVar, BOTH, X, Y, LEFT, RIGHT, END, W
from typing import List, Dict, Any, Callable, Optional
import sys
import os

# Add the parent directory to sys.path to allow relative imports
sys.path.insert(0, os.path.abspath(
    os.path.join(os.path.dirname(__file__), '../..')))


class PatternBuilder:
    """UI component for building filename patterns"""

    # type: ignore
    # type: ignore
    # type: ignore
    # type: ignore
    def __init__(self, parent, on_pattern_applied: Callable[[str], None] = None):
        """
        Initialize the pattern builder

        Args:
            parent: Parent widget
            on_pattern_applied: Callback function for when a pattern is applied
        """
        self.parent = parent
        self.on_pattern_applied = on_pattern_applied

        self.pattern_var = StringVar()
        self.separator_var = StringVar(value="-")
        self.auto_apply_pattern = tk.BooleanVar(value=True)

        self.available_columns = []

        # Initialize UI components
        self.frame = ttk.LabelFrame(
            parent, text="Pattern Builder", padding="10")
        self._setup_ui()

    def _setup_ui(self):
        """Set up the UI components"""
        pattern_grid = ttk.Frame(self.frame)
        pattern_grid.pack(fill=X, expand=True)

        # Columns selection
        ttk.Label(pattern_grid, text="Columns to include:").grid(
            row=0, column=0, sticky=W, pady=2)

        # Available columns dropdown
        self.available_columns_combo = ttk.Combobox(
            pattern_grid, width=15, state="readonly")
        self.available_columns_combo.grid(row=0, column=1, padx=5, pady=2)

        # Add column button
        ttk.Button(pattern_grid, text="Add Column", command=self.add_column_to_pattern).grid(
            row=0, column=2, padx=5, pady=2)

        # Pattern display
        ttk.Label(pattern_grid, text="Pattern:").grid(
            row=1, column=0, sticky=W, pady=2)
        pattern_entry = ttk.Entry(
            pattern_grid, textvariable=self.pattern_var, width=40)
        pattern_entry.grid(row=1, column=1, columnspan=2,
                           padx=5, pady=2, sticky="ew")

        # Reset pattern button
        ttk.Button(pattern_grid, text="Reset", command=self.reset_pattern).grid(
            row=1, column=3, padx=5, pady=2)

        # Separator selection
        ttk.Label(pattern_grid, text="Separator:").grid(
            row=2, column=0, sticky=W, pady=2)
        separator_options = ["-", "_", ".", " ", ",", ";"]
        separator_combo = ttk.Combobox(pattern_grid, textvariable=self.separator_var,
                                       values=separator_options, width=5)
        separator_combo.grid(row=2, column=1, padx=5, pady=2, sticky="w")

        # Apply pattern button
        ttk.Button(pattern_grid, text="Apply Pattern", command=self.apply_pattern).grid(
            row=2, column=2, padx=5, pady=2)

        # Auto-apply pattern checkbox
        ttk.Checkbutton(pattern_grid, text="Auto-apply pattern on file selection",
                        variable=self.auto_apply_pattern).grid(
            row=2, column=3, padx=5, pady=2, sticky="w")

    def add_column_to_pattern(self):
        """Add the selected column to the pattern"""
        if not self.available_columns_combo.get():
            return

        column_name = self.available_columns_combo.get()

        # Get current pattern
        current_pattern = self.pattern_var.get()

        # Add separator if needed
        if current_pattern:
            current_pattern += self.separator_var.get()

        # Add the column name
        self.pattern_var.set(current_pattern + column_name)

    def reset_pattern(self):
        """Reset the pattern to empty"""
        self.pattern_var.set("")

    def apply_pattern(self):
        """Apply the current pattern and notify via callback"""
        if self.on_pattern_applied:
            self.on_pattern_applied(self.pattern_var.get())

    def set_available_columns(self, columns: List[str]):
        """
        Set the available columns for the pattern builder

        Args:
            columns: List of column names
        """
        self.available_columns = columns
        self.available_columns_combo["values"] = columns

        if columns:
            self.available_columns_combo.current(0)

    def get_pattern(self) -> str:
        """
        Get the current pattern

        Returns:
            Pattern string
        """
        return self.pattern_var.get()

    def get_separator(self) -> str:
        """
        Get the current separator

        Returns:
            Separator string
        """
        return self.separator_var.get()

    def is_auto_apply(self) -> bool:
        """
        Check if auto-apply is enabled

        Returns:
            True if auto-apply is enabled, False otherwise
        """
        return self.auto_apply_pattern.get()

    def get_frame(self) -> ttk.LabelFrame:
        """
        Get the frame widget

        Returns:
            Frame widget
        """
        return self.frame

    def generate_filename(self, excel_row: Dict[str, Any]) -> Optional[str]:
        """
        Generate a filename based on the current pattern and Excel data

        Args:
            excel_row: Dictionary containing Excel row data

        Returns:
            Generated filename or None if pattern is invalid
        """
        return generate_filename_from_pattern(
            excel_row,
            self.pattern_var.get(),
            self.separator_var.get()
        )
