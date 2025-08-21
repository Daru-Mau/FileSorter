#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Excel panel component for displaying Excel data
"""

from src.models.excel_model import ExcelModel
import tkinter as tk
from tkinter import ttk, StringVar, BOTH, X, Y, LEFT, RIGHT, END, W
from typing import List, Dict, Any, Callable, Optional, Tuple
import sys
import os

import pandas as pd

# Add the parent directory to sys.path to allow relative imports
sys.path.insert(0, os.path.abspath(
    os.path.join(os.path.dirname(__file__), '../..')))


class ExcelPanel:
    """UI component for displaying Excel data"""

    # type: ignore
    # type: ignore
    # type: ignore
    # type: ignore
    def __init__(self, parent, on_excel_row_select: Callable[[int, Dict[str, Any]], None] = None):
        """
        Initialize the Excel panel

        Args:
            parent: Parent widget
            on_excel_row_select: Callback function for row selection
        """
        self.parent = parent
        self.on_excel_row_select = on_excel_row_select

        self.excel_model = ExcelModel()
        self.name_column = StringVar()
        self.id_column = StringVar()
        self.date_column = StringVar()

        # Initialize UI components
        self.frame = ttk.LabelFrame(parent, text="Excel Data", padding="10")
        self._setup_ui()

    def _setup_ui(self):
        """Set up the UI components"""
        # Excel column mapping frame
        self.column_map_frame = ttk.LabelFrame(
            self.frame, text="Column Mapping", padding="5")
        self.column_map_frame.pack(fill=X, pady=5)

        # Name column mapping
        ttk.Label(self.column_map_frame, text="Name Column:").grid(
            row=0, column=0, sticky=W, pady=5)
        self.name_column_combo = ttk.Combobox(
            self.column_map_frame, textvariable=self.name_column, width=20)
        self.name_column_combo.grid(row=0, column=1, padx=5, pady=5)

        # ID column mapping
        ttk.Label(self.column_map_frame, text="ID Column:").grid(
            row=0, column=2, sticky=W, pady=5)
        self.id_column_combo = ttk.Combobox(
            self.column_map_frame, textvariable=self.id_column, width=20)
        self.id_column_combo.grid(row=0, column=3, padx=5, pady=5)

        # Date column mapping
        ttk.Label(self.column_map_frame, text="Date Column:").grid(
            row=0, column=4, sticky=W, pady=5)
        self.date_column_combo = ttk.Combobox(
            self.column_map_frame, textvariable=self.date_column, width=20)
        self.date_column_combo.grid(row=0, column=5, padx=5, pady=5)

        # Excel data treeview
        self.excel_frame = ttk.Frame(self.frame)
        self.excel_frame.pack(fill=BOTH, expand=True, pady=5)

        self.excel_tree = ttk.Treeview(self.excel_frame, show="headings")

        # Scrollbars for excel tree
        excel_y_scrollbar = ttk.Scrollbar(
            self.excel_frame, orient=tk.VERTICAL, command=self.excel_tree.yview)
        excel_x_scrollbar = ttk.Scrollbar(
            self.excel_frame, orient=tk.HORIZONTAL, command=self.excel_tree.xview)
        self.excel_tree.configure(
            yscrollcommand=excel_y_scrollbar.set, xscrollcommand=excel_x_scrollbar.set)

        # Pack excel tree and scrollbars
        self.excel_tree.pack(side=LEFT, fill=BOTH, expand=True)
        excel_y_scrollbar.pack(side=RIGHT, fill=Y)
        excel_x_scrollbar.pack(side=tk.BOTTOM, fill=X)

        # Bind treeview selection event
        self.excel_tree.bind('<<TreeviewSelect>>',
                             self._on_excel_row_select_internal)

    def load_excel_data(self, file_path: str) -> bool:
        """
        Load data from Excel file

        Args:
            file_path: Path to the Excel file

        Returns:
            True if loading was successful, False otherwise
        """
        # Load data into the model
        success = self.excel_model.load_data(file_path)

        if success:
            self._update_ui_from_model()
            return True

        return False

    def _update_ui_from_model(self):
        """Update UI components from the Excel model"""
        if self.excel_model.data is None:
            return

        # Clear current treeview
        for column in self.excel_tree["columns"]:
            self.excel_tree.heading(column, text="")

        for item in self.excel_tree.get_children():
            self.excel_tree.delete(item)

        # Get columns from model
        columns = self.excel_model.columns

        # Configure treeview columns
        self.excel_tree["columns"] = columns

        # Update column mapping dropdown options
        self.name_column_combo["values"] = columns
        self.id_column_combo["values"] = columns
        self.date_column_combo["values"] = columns

        # Set column mappings from model
        self.name_column.set(self.excel_model.name_column)
        self.id_column.set(self.excel_model.id_column)
        self.date_column.set(self.excel_model.date_column)

        # Set up headings
        for col in columns:
            self.excel_tree.heading(col, text=col)
            # Adjust column width based on content
            max_width = max([len(str(self.excel_model.data[col].iloc[i])) for i in range(
                min(10, len(self.excel_model.data)))] + [len(col)])
            self.excel_tree.column(col, width=max_width * 10)

        # Insert data rows
        for i, row in self.excel_model.data.iterrows():
            values = [row[col] for col in columns]
            self.excel_tree.insert("", END, values=values)

    def _on_excel_row_select_internal(self, event):
        """Internal handler for Excel row selection"""
        if not self.excel_tree.selection():
            return

        # Get selected item
        selected_item = self.excel_tree.selection()[0]

        # Get item values
        values = self.excel_tree.item(selected_item)["values"]

        # Create a dictionary of column names and values
        if self.excel_model.data is not None:
            columns = self.excel_model.columns
            row_data = {columns[i]: values[i] for i in range(len(columns))}

            # Get the row index
            items = self.excel_tree.get_children()
            row_index = items.index(selected_item)

            # Call the external handler if provided
            if self.on_excel_row_select:
                self.on_excel_row_select(row_index, row_data)

    def find_match_for_filename(self, filename: str) -> Tuple[bool, Optional[str]]:
        """
        Find a matching row for a filename

        Args:
            filename: Filename to match

        Returns:
            Tuple containing:
            - Boolean indicating if a match was found
            - Item ID of the match in the treeview (or None)
        """
        if self.excel_model.data is None or not self.name_column.get():
            return False, None

        # Update the Excel model with current column mappings
        self.excel_model.name_column = self.name_column.get()
        self.excel_model.id_column = self.id_column.get()
        self.excel_model.date_column = self.date_column.get()

        # Remove extension from filename for matching
        filename_without_ext = filename
        if '.' in filename:
            filename_without_ext = filename.rsplit('.', 1)[0]

        # Clear previous selection
        for item in self.excel_tree.selection():
            self.excel_tree.selection_remove(item)

        # Try exact match first
        for item in self.excel_tree.get_children():
            values = self.excel_tree.item(item)["values"]
            col_index = self.excel_model.columns.index(self.name_column.get())
            excel_value = str(values[col_index])

            if excel_value == filename_without_ext or excel_value == filename:
                self.excel_tree.selection_set(item)
                self.excel_tree.see(item)
                return True, item

        # Try partial match if exact match failed
        for item in self.excel_tree.get_children():
            values = self.excel_tree.item(item)["values"]
            col_index = self.excel_model.columns.index(self.name_column.get())
            excel_value = str(values[col_index])

            if filename_without_ext in excel_value or excel_value in filename_without_ext:
                self.excel_tree.selection_set(item)
                self.excel_tree.see(item)
                return True, item

        return False, None

    def get_selected_row_data(self) -> Optional[Dict[str, Any]]:
        """
        Get the currently selected row data

        Returns:
            Dictionary with column names as keys and row values as values,
            or None if no selection
        """
        if not self.excel_tree.selection():
            return None

        selected_item = self.excel_tree.selection()[0]
        values = self.excel_tree.item(selected_item)["values"]

        if self.excel_model.data is not None:
            columns = self.excel_model.columns
            return {columns[i]: values[i] for i in range(len(columns))}

        return None

    def get_column_mappings(self) -> Dict[str, str]:
        """
        Get the current column mappings

        Returns:
            Dictionary with mapping types as keys and column names as values
        """
        return {
            "name": self.name_column.get(),
            "id": self.id_column.get(),
            "date": self.date_column.get()
        }

    def get_frame(self) -> ttk.LabelFrame:
        """
        Get the frame widget

        Returns:
            Frame widget
        """
        return self.frame
