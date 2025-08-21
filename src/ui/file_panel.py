#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
File panel component for displaying and managing files
"""

from src.utils.file_utils import filter_files_by_extension, search_files_by_name
from src.models.file_model import FileModel
import tkinter as tk
from tkinter import ttk, StringVar, BOTH, X, Y, LEFT, RIGHT, END, W
from typing import List, Dict, Any, Callable, Optional
import sys
import os

# Add the parent directory to sys.path to allow relative imports
sys.path.insert(0, os.path.abspath(
    os.path.join(os.path.dirname(__file__), '../..')))


class FilePanel:
    """UI component for displaying and managing files"""

    def __init__(self, parent, on_file_select: Callable[[FileModel], None]):
        """
        Initialize the file panel

        Args:
            parent: Parent widget
            on_file_select: Callback function for file selection
        """
        self.parent = parent
        self.on_file_select = on_file_select
        self.search_term = StringVar()
        self.file_extension_filter = StringVar(value="All Files")

        self.all_files = []  # All files in directory
        self.filtered_files = []  # Files after filtering
        self.selected_file = None

        # Initialize UI components
        self.frame = ttk.LabelFrame(parent, text="Files", padding="10")
        self._setup_ui()

    def _setup_ui(self):
        """Set up the UI components"""
        # Search box and filter for files
        search_frame = ttk.Frame(self.frame)
        search_frame.pack(fill=X, pady=5)

        # File filter dropdown
        ttk.Label(search_frame, text="Filter:").pack(side=LEFT, padx=5)
        extension_values = ["All Files", "PDF Files", "Excel Files",
                            "Word Files", "Image Files", "Text Files", "Custom..."]
        self.extension_filter_combo = ttk.Combobox(search_frame, textvariable=self.file_extension_filter,
                                                   values=extension_values, width=12, state="readonly")
        self.extension_filter_combo.pack(side=LEFT, padx=5)
        self.extension_filter_combo.bind(
            "<<ComboboxSelected>>", self.apply_filter)

        # Custom extension entry (initially hidden)
        self.custom_extension = StringVar()
        self.custom_ext_frame = ttk.Frame(search_frame)
        self.custom_ext_frame.pack(side=LEFT, padx=5)
        self.custom_ext_entry = ttk.Entry(
            self.custom_ext_frame, textvariable=self.custom_extension, width=8)
        self.custom_ext_entry.pack(side=LEFT)
        ttk.Button(self.custom_ext_frame, text="Apply",
                   command=self.apply_custom_filter).pack(side=LEFT, padx=2)
        self.custom_ext_frame.pack_forget()  # Hide initially

        # Search box
        ttk.Label(search_frame, text="Search:").pack(side=LEFT, padx=5)
        ttk.Entry(search_frame, textvariable=self.search_term,
                  width=30).pack(side=LEFT, padx=5)
        ttk.Button(search_frame, text="Search",
                   command=self.search_files).pack(side=LEFT, padx=5)
        ttk.Button(search_frame, text="Refresh",
                   command=self.refresh).pack(side=LEFT, padx=5)

        # Files listbox with scrollbar
        self.files_listbox = tk.Listbox(self.frame, width=40, height=20)
        self.files_listbox.pack(side=LEFT, fill=BOTH, expand=True)
        files_scrollbar = ttk.Scrollbar(
            self.frame, orient=tk.VERTICAL, command=self.files_listbox.yview)
        self.files_listbox.configure(yscrollcommand=files_scrollbar.set)
        files_scrollbar.pack(side=RIGHT, fill=Y)

        # Bind listbox selection event
        self.files_listbox.bind('<<ListboxSelect>>',
                                self._on_file_select_internal)

    def apply_filter(self, event=None):
        """Apply file extension filter based on selected value"""
        filter_value = self.file_extension_filter.get()

        # Show custom entry field if "Custom..." is selected
        if filter_value == "Custom...":
            self.custom_ext_frame.pack(side=LEFT, padx=5)
            self.custom_ext_entry.focus()
        else:
            self.custom_ext_frame.pack_forget()
            self._filter_files_by_extension(filter_value)

    def apply_custom_filter(self):
        """Apply custom file extension filter"""
        custom_ext = self.custom_extension.get().strip()
        if not custom_ext:
            return

        # Add dot if not provided
        if not custom_ext.startswith('.'):
            custom_ext = '.' + custom_ext

        self._filter_files_by_extension(custom_ext)

    def _filter_files_by_extension(self, filter_value):
        """Filter files by extension and update display"""
        # Map filter values to extensions
        extension_map = {
            "All Files": None,
            "PDF Files": [".pdf"],
            "Excel Files": [".xlsx", ".xls", ".csv"],
            "Word Files": [".docx", ".doc"],
            "Image Files": [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff"],
            "Text Files": [".txt", ".text", ".md", ".rtf"]
        }

        # Get extensions to filter by
        if filter_value in extension_map:
            extensions = extension_map[filter_value]
        else:
            # Custom extension
            extensions = [filter_value.lower()]

        # Clear listbox
        self.files_listbox.delete(0, END)

        # Apply filter
        if not self.all_files:
            return

        self.filtered_files = filter_files_by_extension(
            self.all_files, extensions)

        # Update listbox
        for file in self.filtered_files:
            self.files_listbox.insert(END, file.name)

    def search_files(self):
        """Search files by name"""
        search_text = self.search_term.get().lower()

        # If search is empty, show all files based on current filter
        if not search_text:
            self.apply_filter()
            return

        # Clear listbox
        self.files_listbox.delete(0, END)

        # Filter by search term
        matching_files = search_files_by_name(self.filtered_files, search_text)

        # Update filtered_files to only include search matches
        self.filtered_files = matching_files

        # Update listbox
        for file in self.filtered_files:
            self.files_listbox.insert(END, file.name)

    def _on_file_select_internal(self, event):
        """Internal handler for file selection from listbox"""
        if not self.files_listbox.curselection():
            return

        # Get selected file data
        selected_index = self.files_listbox.curselection()[0]
        if selected_index < len(self.filtered_files):
            self.selected_file = self.filtered_files[selected_index]

            # Call the external handler if provided
            if self.on_file_select:
                self.on_file_select(self.selected_file)

    def refresh(self):
        """Refresh the file list based on current filters"""
        self.apply_filter()

    def update_files(self, files: List[FileModel]):
        """
        Update the file list with new data

        Args:
            files: List of FileModel objects
        """
        self.all_files = files
        self.apply_filter()

    def get_selected_file(self) -> Optional[FileModel]:
        """
        Get the currently selected file

        Returns:
            Selected FileModel or None if no selection
        """
        return self.selected_file

    def get_frame(self) -> ttk.LabelFrame:
        """
        Get the frame widget

        Returns:
            Frame widget
        """
        return self.frame
