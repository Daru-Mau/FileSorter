#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Main window component for the Excel File Renamer application
"""

from src.utils.string_utils import is_valid_filename, sanitize_filename
from src.utils.file_utils import scan_directory, rename_file
from src.ui.pattern_builder import PatternBuilder
from src.ui.excel_panel import ExcelPanel
from src.ui.file_panel import FilePanel
from src.models.excel_model import ExcelModel
from src.models.file_model import FileModel
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, StringVar, BOTH, X, Y, LEFT, RIGHT, END, W, SUNKEN
from typing import List, Dict, Any, Optional
import sys
import os

# Add the parent directory to sys.path to allow relative imports
sys.path.insert(0, os.path.abspath(
    os.path.join(os.path.dirname(__file__), '../..')))


class MainWindow:
    """Main window for the Excel File Renamer application"""

    def __init__(self, root):
        """
        Initialize the main window

        Args:
            root: Tkinter root window
        """
        self.root = root
        self.root.title("Excel-Based File Renamer")
        self.root.geometry("1100x700")
        self.root.minsize(900, 600)

        self.source_folder = StringVar()
        self.excel_file_path = StringVar()
        self.manual_filename = StringVar()
        self.keep_extension = tk.BooleanVar(value=True)

        self.selected_file = None

        # Initialize UI components
        self.setup_ui()

    def setup_ui(self):
        """Set up the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=BOTH, expand=True)

        # Top section - file and excel selection
        top_frame = ttk.LabelFrame(
            main_frame, text="File Selection", padding="10")
        top_frame.pack(fill=X, pady=5)

        # Source folder
        ttk.Label(top_frame, text="Source Folder:").grid(
            row=0, column=0, sticky=W, pady=5)
        ttk.Entry(top_frame, textvariable=self.source_folder,
                  width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(top_frame, text="Browse...", command=self.browse_source).grid(
            row=0, column=2, padx=5, pady=5)

        # Excel file
        ttk.Label(top_frame, text="Excel File:").grid(
            row=1, column=0, sticky=W, pady=5)
        ttk.Entry(top_frame, textvariable=self.excel_file_path,
                  width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(top_frame, text="Browse...", command=self.browse_excel).grid(
            row=1, column=2, padx=5, pady=5)
        ttk.Button(top_frame, text="Load Excel", command=self.load_excel_data).grid(
            row=1, column=3, padx=5, pady=5)

        # Pattern builder
        self.pattern_builder = PatternBuilder(
            main_frame, self.on_pattern_applied)
        self.pattern_builder.get_frame().pack(fill=X, pady=5)

        # Middle section - paned window for files and excel data
        paned_window = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned_window.pack(fill=BOTH, expand=True, pady=5)

        # File panel
        self.file_panel = FilePanel(paned_window, self.on_file_select)
        paned_window.add(self.file_panel.get_frame(), weight=1)

        # Excel panel
        self.excel_panel = ExcelPanel(paned_window, self.on_excel_row_select)
        paned_window.add(self.excel_panel.get_frame(), weight=2)

        # Manual rename section
        manual_frame = ttk.LabelFrame(
            main_frame, text="Manual Rename", padding="10")
        manual_frame.pack(fill=X, pady=5)

        # Manual rename entry
        ttk.Label(manual_frame, text="Custom Filename:").grid(
            row=0, column=0, sticky=W, pady=5)
        self.manual_entry = ttk.Entry(
            manual_frame, textvariable=self.manual_filename, width=40)
        self.manual_entry.grid(row=0, column=1, padx=5, pady=5, sticky=W)

        # Keep extension checkbox
        ttk.Checkbutton(manual_frame, text="Keep original extension", variable=self.keep_extension).grid(
            row=0, column=2, padx=5, pady=5, sticky=W)

        # Manual rename button
        self.manual_rename_button = ttk.Button(manual_frame, text="Apply Manual Rename",
                                               command=self.manual_rename, state=tk.DISABLED)
        self.manual_rename_button.grid(row=0, column=3, padx=5, pady=5)

        # Bottom section - actions
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=X, pady=10)

        # Details label
        self.details_label = ttk.Label(
            action_frame, text="", wraplength=800, justify=tk.LEFT)
        self.details_label.pack(side=LEFT, fill=X, expand=True, padx=5)

        # Rename button
        self.rename_button = ttk.Button(
            action_frame, text="Rename Selected File", command=self.rename_file, state=tk.DISABLED)
        self.rename_button.pack(side=RIGHT, padx=5)

        # Status bar
        self.status_var = StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(
            main_frame, textvariable=self.status_var, relief=SUNKEN, anchor=W)
        status_bar.pack(side=tk.BOTTOM, fill=X)

    def browse_source(self):
        """Browse for source folder"""
        folder = filedialog.askdirectory()
        if folder:
            self.source_folder.set(folder)
            self.scan_files()

    def browse_excel(self):
        """Browse for Excel file"""
        file_path = filedialog.askopenfilename(filetypes=[
            ("Excel files", "*.xlsx;*.xls"),
            ("All files", "*.*")
        ])
        if file_path:
            self.excel_file_path.set(file_path)

    def load_excel_data(self):
        """Load data from Excel file"""
        excel_path = self.excel_file_path.get()
        if not excel_path:
            messagebox.showerror("Error", "Please select an Excel file")
            return

        try:
            # Load Excel data into the Excel panel
            success = self.excel_panel.load_excel_data(excel_path)

            if success:
                # Update pattern builder with available columns
                columns = self.excel_panel.excel_model.columns
                self.pattern_builder.set_available_columns(columns)

                self.status_var.set(
                    # type: ignore
                    # type: ignore
                    # type: ignore
                    # type: ignore
                    f"Loaded {len(self.excel_panel.excel_model.data)} rows from Excel")
            else:
                self.status_var.set("Error loading Excel data")

        except Exception as e:
            messagebox.showerror(
                "Error", f"Failed to load Excel file: {str(e)}")
            self.status_var.set("Error loading Excel data")

    def scan_files(self):
        """Scan files in the source folder"""
        source = self.source_folder.get()
        if not source:
            messagebox.showerror("Error", "Please select a source folder")
            return

        try:
            # Scan directory and update file panel
            files = scan_directory(source)
            self.file_panel.update_files(files)

            self.status_var.set(f"Scanned {len(files)} files")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.status_var.set("Error scanning files")

    def on_file_select(self, file_model: FileModel):
        """
        Handle file selection

        Args:
            file_model: Selected FileModel
        """
        self.selected_file = file_model

        # Enable manual rename button
        self.manual_rename_button.config(state=tk.NORMAL)

        # Populate manual rename field with current filename without extension
        self.manual_filename.set(file_model.filename_without_ext)

        # Try to find a match in Excel data
        if self.excel_panel.excel_model.data is not None:
            match_found, matched_item = self.excel_panel.find_match_for_filename(
                file_model.name)

            if match_found:
                self.status_var.set(f"Found match for {file_model.name}")
                self.rename_button.config(state=tk.NORMAL)

                # Get row data
                row_data = self.excel_panel.get_selected_row_data()
                if row_data:
                    self.update_details(file_model, row_data)

                    # Auto-apply pattern if enabled
                    if self.pattern_builder.is_auto_apply() and self.pattern_builder.get_pattern():
                        custom_name = self.pattern_builder.generate_filename(
                            row_data)
                        if custom_name:
                            self.manual_filename.set(custom_name)
                            self.status_var.set(
                                f"Auto-applied pattern: {custom_name}")

                            # Add preview to details
                            extension = file_model.extension
                            if self.keep_extension.get():
                                preview = f"{custom_name}{extension}"
                            else:
                                preview = custom_name

                            self.details_label.config(
                                text=self.details_label.cget("text") + f"\n\nPattern-based filename: {preview}")
            else:
                self.status_var.set(f"No match found for {file_model.name}")
                self.details_label.config(
                    text=f"No match found in Excel data for {file_model.name}")
                self.rename_button.config(state=tk.DISABLED)
        else:
            self.status_var.set(
                "Excel data not loaded or columns not configured")

    def on_excel_row_select(self, row_index: int, row_data: Dict[str, Any]):
        """
        Handle Excel row selection

        Args:
            row_index: Index of the selected row
            row_data: Dictionary of row data
        """
        # Update details if a file is selected
        if self.selected_file:
            self.update_details(self.selected_file, row_data)
            self.rename_button.config(state=tk.NORMAL)

    def on_pattern_applied(self, pattern: str):
        """
        Handle pattern application

        Args:
            pattern: Applied pattern
        """
        if not self.selected_file or not self.excel_panel.get_selected_row_data():
            messagebox.showerror("Error", "No file or Excel entry selected")
            return

        if not pattern:
            messagebox.showerror(
                "Error", "Pattern is empty. Please add columns to your pattern.")
            return

        # Get selected Excel data
        row_data = self.excel_panel.get_selected_row_data()
        if not row_data:
            return

        # Generate filename from pattern
        custom_name = self.pattern_builder.generate_filename(row_data)
        if not custom_name:
            messagebox.showerror(
                "Error", "Failed to generate filename from pattern")
            return

        # Update the manual filename field
        self.manual_filename.set(custom_name)

        # Preview the result
        extension = self.selected_file.extension
        if self.keep_extension.get():
            preview = f"{custom_name}{extension}"
        else:
            preview = custom_name

        messagebox.showinfo("Pattern Applied",
                            f"Pattern applied successfully!\n\nProposed filename: {preview}\n\n"
                            f"Click 'Apply Manual Rename' to rename the file.")

    def update_details(self, file_model: FileModel, excel_data: Dict[str, Any]):
        """
        Update details label with file and Excel information

        Args:
            file_model: FileModel object
            excel_data: Dictionary of Excel row data
        """
        # Format file details
        file_details = (
            f"File: {file_model.name}\n"
            f"Size: {file_model.format_size()}\n"
            f"Modified: {file_model.mod_time_formatted}\n\n"
        )

        # Format Excel match details
        excel_details = "Excel Match:\n"
        for col, val in excel_data.items():
            excel_details += f"{col}: {val}\n"

        # Show proposed new filename based on ID column
        mappings = self.excel_panel.get_column_mappings()
        id_val = excel_data.get(mappings["id"], "")
        if id_val:
            extension = file_model.extension
            proposed_name = f"{id_val}{extension}"
            excel_details += f"\nProposed new filename: {proposed_name}"

        # Update details label
        self.details_label.config(text=file_details + excel_details)

    def rename_file(self):
        """Rename the selected file based on Excel data"""
        if not self.selected_file:
            messagebox.showerror("Error", "No file selected")
            return

        # Get Excel row data
        row_data = self.excel_panel.get_selected_row_data()
        if not row_data:
            messagebox.showerror("Error", "No Excel entry selected")
            return

        # Get column mappings
        mappings = self.excel_panel.get_column_mappings()

        # Get ID value for new name
        id_val = row_data.get(mappings["id"], "")
        if not id_val:
            messagebox.showerror(
                "Error", f"No ID value found in the '{mappings['id']}' column")
            return

        # Format ID value to ensure it's a valid filename
        id_val = str(id_val).strip()
        id_val = sanitize_filename(id_val)

        # Create new filename
        extension = self.selected_file.extension
        new_filename = f"{id_val}{extension}"

        # Check if source and new names are the same
        if new_filename == self.selected_file.name:
            messagebox.showinfo(
                "Info", "The file already has the correct name")
            return

        try:
            # Rename file
            success = rename_file(self.selected_file,
                                  id_val, keep_extension=True)

            if success:
                # Refresh file list
                self.scan_files()

                # Update status
                self.status_var.set(f"Renamed file to {new_filename}")
                messagebox.showinfo(
                    "Success", f"Renamed file to {new_filename}")
            else:
                messagebox.showerror("Error", "Failed to rename file")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to rename file: {str(e)}")
            self.status_var.set("Error renaming file")

    def manual_rename(self):
        """Rename the selected file using a custom name"""
        if not self.selected_file:
            messagebox.showerror("Error", "No file selected")
            return

        # Get the custom filename
        custom_name = self.manual_filename.get().strip()
        if not custom_name:
            messagebox.showerror("Error", "Please enter a custom filename")
            return

        # Validate filename
        if not is_valid_filename(custom_name):
            messagebox.showerror(
                "Error", "Filename contains invalid characters (\\/*?:\"<>|)")
            return

        try:
            # Rename file
            success = rename_file(
                self.selected_file,
                custom_name,
                keep_extension=self.keep_extension.get()
            )

            if success:
                # Refresh file list
                self.scan_files()

                # Clear the manual filename entry
                self.manual_filename.set("")

                # Determine the new filename for status message
                extension = self.selected_file.extension
                if self.keep_extension.get() and not custom_name.lower().endswith(extension.lower()):
                    new_filename = f"{custom_name}{extension}"
                else:
                    new_filename = custom_name

                # Update status
                self.status_var.set(f"Manually renamed file to {new_filename}")
                messagebox.showinfo(
                    "Success", f"Renamed file to {new_filename}")
            else:
                messagebox.showerror("Error", "Failed to rename file")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to rename file: {str(e)}")
            self.status_var.set("Error renaming file")
