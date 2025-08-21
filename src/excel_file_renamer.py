#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Excel-Based File Renamer

A GUI application for renaming files based on matching data in Excel spreadsheets.
"""

import os
import shutil
import re
import pandas as pd
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, StringVar, BOTH, X, Y, LEFT, RIGHT, END, W, SUNKEN, VERTICAL


class ExcelFileRenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel-Based File Renamer")
        self.root.geometry("1100x700")
        self.root.minsize(900, 600)

        self.source_folder = StringVar()
        self.excel_file_path = StringVar()
        self.search_term = StringVar()
        self.file_extension_filter = StringVar(value="All Files")

        self.selected_file = None
        self.file_list_data = []
        self.all_files_data = []  # Store all files before filtering
        self.excel_data = None
        self.name_column = StringVar()
        self.id_column = StringVar()
        self.date_column = StringVar()

        self.setup_ui()

    def setup_ui(self):
        """Setup the user interface"""
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

        # Excel column mapping frame
        excel_map_frame = ttk.LabelFrame(
            main_frame, text="Excel Column Mapping", padding="10")
        excel_map_frame.pack(fill=X, pady=5)

        # Name column mapping
        ttk.Label(excel_map_frame, text="Name Column:").grid(
            row=0, column=0, sticky=W, pady=5)
        self.name_column_combo = ttk.Combobox(
            excel_map_frame, textvariable=self.name_column, width=20)
        self.name_column_combo.grid(row=0, column=1, padx=5, pady=5)

        # ID column mapping
        ttk.Label(excel_map_frame, text="ID Column:").grid(
            row=0, column=2, sticky=W, pady=5)
        self.id_column_combo = ttk.Combobox(
            excel_map_frame, textvariable=self.id_column, width=20)
        self.id_column_combo.grid(row=0, column=3, padx=5, pady=5)

        # Date column mapping
        ttk.Label(excel_map_frame, text="Date Column:").grid(
            row=0, column=4, sticky=W, pady=5)
        self.date_column_combo = ttk.Combobox(
            excel_map_frame, textvariable=self.date_column, width=20)
        self.date_column_combo.grid(row=0, column=5, padx=5, pady=5)

        # Pattern builder section - moved to the top
        pattern_frame = ttk.LabelFrame(
            main_frame, text="Pattern Builder", padding="10")
        pattern_frame.pack(fill=X, pady=5)

        pattern_grid = ttk.Frame(pattern_frame)
        pattern_grid.pack(fill=X, expand=True)

        # Columns selection
        ttk.Label(pattern_grid, text="Columns to include:").grid(
            row=0, column=0, sticky=W, pady=2)

        # Available columns dropdown
        self.available_columns = ttk.Combobox(
            pattern_grid, width=15, state="readonly")
        self.available_columns.grid(row=0, column=1, padx=5, pady=2)

        # Add column button
        ttk.Button(pattern_grid, text="Add Column", command=self.add_column_to_pattern).grid(
            row=0, column=2, padx=5, pady=2)

        # Pattern display
        ttk.Label(pattern_grid, text="Pattern:").grid(
            row=1, column=0, sticky=W, pady=2)
        self.pattern_var = StringVar()
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
        self.separator_var = StringVar(value="-")
        separator_options = ["-", "_", ".", " ", ",", ";"]
        separator_combo = ttk.Combobox(pattern_grid, textvariable=self.separator_var,
                                       values=separator_options, width=5)
        separator_combo.grid(row=2, column=1, padx=5, pady=2, sticky="w")

        # Apply pattern button
        ttk.Button(pattern_grid, text="Apply Pattern", command=self.apply_pattern).grid(
            row=2, column=2, padx=5, pady=2)

        # Auto-apply pattern checkbox
        self.auto_apply_pattern = tk.BooleanVar(value=True)
        ttk.Checkbutton(pattern_grid, text="Auto-apply pattern on file selection",
                        variable=self.auto_apply_pattern).grid(
            row=2, column=3, padx=5, pady=2, sticky="w")

        # Middle section - paned window for files and excel data
        paned_window = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned_window.pack(fill=BOTH, expand=True, pady=5)

        # Left panel - Files list
        files_frame = ttk.LabelFrame(paned_window, text="Files", padding="10")
        paned_window.add(files_frame, weight=1)

        # Search box for files
        search_frame = ttk.Frame(files_frame)
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

        ttk.Label(search_frame, text="Search:").pack(side=LEFT, padx=5)
        ttk.Entry(search_frame, textvariable=self.search_term,
                  width=30).pack(side=LEFT, padx=5)
        ttk.Button(search_frame, text="Search",
                   command=self.search_files).pack(side=LEFT, padx=5)
        ttk.Button(search_frame, text="Scan Files",
                   command=self.scan_files).pack(side=LEFT, padx=5)

        # Files listbox with scrollbar
        self.files_listbox = tk.Listbox(files_frame, width=40, height=20)
        self.files_listbox.pack(side=LEFT, fill=BOTH, expand=True)
        files_scrollbar = ttk.Scrollbar(
            files_frame, orient=VERTICAL, command=self.files_listbox.yview)
        self.files_listbox.configure(yscrollcommand=files_scrollbar.set)
        files_scrollbar.pack(side=RIGHT, fill=Y)

        # Bind listbox selection event
        self.files_listbox.bind('<<ListboxSelect>>', self.on_file_select)

        # Right panel - Excel data display
        excel_frame = ttk.LabelFrame(
            paned_window, text="Excel Data", padding="10")
        paned_window.add(excel_frame, weight=2)

        # Excel data treeview
        self.excel_tree = ttk.Treeview(excel_frame, show="headings")

        # Scrollbars for excel tree
        excel_y_scrollbar = ttk.Scrollbar(
            excel_frame, orient=VERTICAL, command=self.excel_tree.yview)
        excel_x_scrollbar = ttk.Scrollbar(
            excel_frame, orient=tk.HORIZONTAL, command=self.excel_tree.xview)
        self.excel_tree.configure(
            yscrollcommand=excel_y_scrollbar.set, xscrollcommand=excel_x_scrollbar.set)

        # Pack excel tree and scrollbars
        self.excel_tree.pack(side=LEFT, fill=BOTH, expand=True)
        excel_y_scrollbar.pack(side=RIGHT, fill=Y)
        excel_x_scrollbar.pack(side=tk.BOTTOM, fill=X)

        # Manual rename section
        manual_frame = ttk.LabelFrame(
            main_frame, text="Manual Rename", padding="10")
        manual_frame.pack(fill=X, pady=5)

        # Manual rename entry
        ttk.Label(manual_frame, text="Custom Filename:").grid(
            row=0, column=0, sticky=W, pady=5)
        self.manual_filename = StringVar()
        self.manual_entry = ttk.Entry(
            manual_frame, textvariable=self.manual_filename, width=40)
        self.manual_entry.grid(row=0, column=1, padx=5, pady=5, sticky=W)

        # Keep extension checkbox
        self.keep_extension = tk.BooleanVar(value=True)
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
            # Load Excel data
            self.excel_data = pd.read_excel(excel_path)

            # Clear current treeview
            for column in self.excel_tree["columns"]:
                self.excel_tree.heading(column, text="")

            for item in self.excel_tree.get_children():
                self.excel_tree.delete(item)

            # Configure columns
            columns = list(self.excel_data.columns)
            self.excel_tree["columns"] = columns

            # Update column mapping dropdown options
            self.name_column_combo["values"] = columns
            self.id_column_combo["values"] = columns
            self.date_column_combo["values"] = columns

            # Set default values if possible
            name_matches = [col for col in columns if 'name' in col.lower()]
            id_matches = [col for col in columns if 'id' in col.lower()]
            date_matches = [col for col in columns if 'date' in col.lower()]

            if name_matches:
                self.name_column.set(name_matches[0])
            elif columns:
                self.name_column.set(columns[0])

            if id_matches:
                self.id_column.set(id_matches[0])
            elif len(columns) > 1:
                self.id_column.set(columns[1])

            if date_matches:
                self.date_column.set(date_matches[0])

            # Set up headings
            for col in columns:
                self.excel_tree.heading(col, text=col)
                # Adjust column width based on content
                max_width = max([len(str(self.excel_data[col].iloc[i])) for i in range(
                    min(10, len(self.excel_data)))] + [len(col)])
                self.excel_tree.column(col, width=max_width * 10)

            # Insert data rows
            for i, row in self.excel_data.iterrows():
                values = [row[col] for col in columns]
                self.excel_tree.insert("", END, values=values)

            self.status_var.set(
                f"Loaded {len(self.excel_data)} rows from Excel")

            # Update available columns for pattern builder
            self.update_available_columns()

        except Exception as e:
            messagebox.showerror(
                "Error", f"Failed to load Excel file: {str(e)}")
            self.status_var.set("Error loading Excel data")

    def apply_filter(self, event=None):
        """Apply file extension filter based on selected value"""
        filter_value = self.file_extension_filter.get()

        # Show custom entry field if "Custom..." is selected
        if filter_value == "Custom...":
            self.custom_ext_frame.pack(side=LEFT, padx=5)
            self.custom_ext_entry.focus()
        else:
            self.custom_ext_frame.pack_forget()
            self.filter_files_by_extension(filter_value)

    def apply_custom_filter(self):
        """Apply custom file extension filter"""
        custom_ext = self.custom_extension.get().strip()
        if not custom_ext:
            messagebox.showinfo(
                "Info", "Please enter a custom extension (e.g., .pdf)")
            return

        # Add dot if not provided
        if not custom_ext.startswith('.'):
            custom_ext = '.' + custom_ext

        self.filter_files_by_extension(custom_ext)

    def filter_files_by_extension(self, filter_value):
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

        # Start with all files
        if not self.all_files_data:
            # If no scan has been done yet, nothing to filter
            return

        # Apply filter
        if extensions is None:  # All Files
            self.file_list_data = self.all_files_data.copy()
        else:
            self.file_list_data = [file for file in self.all_files_data
                                   if any(file["extension"].lower() == ext.lower() for ext in extensions)]

        # Update listbox
        for file_data in self.file_list_data:
            self.files_listbox.insert(END, file_data["name"])

        # Update status
        extension_text = ", ".join(
            extensions) if extensions else "all extensions"
        self.status_var.set(
            f"Showing {len(self.file_list_data)} files with {extension_text}")

    def scan_files(self):
        """Scan files in the source folder"""
        source = self.source_folder.get()
        if not source:
            messagebox.showerror("Error", "Please select a source folder")
            return

        # Clear current data
        self.all_files_data = []
        self.file_list_data = []
        self.files_listbox.delete(0, END)

        try:
            # Get all files in the folder
            files = list(Path(source).glob("*"))

            # Filter out directories
            files = [f for f in files if f.is_file()]

            # Sort files by name
            files.sort(key=lambda x: x.name.lower())

            # Store file data
            for file_path in files:
                file_stat = file_path.stat()
                mod_time = datetime.fromtimestamp(
                    file_stat.st_mtime).strftime('%Y-%m-%d %H:%M')

                # Store file data
                file_data = {
                    "path": file_path,
                    "name": file_path.name,
                    "size": file_stat.st_size,
                    "mod_time": file_stat.st_mtime,
                    "mod_time_formatted": mod_time,
                    "extension": file_path.suffix.lower()
                }
                self.all_files_data.append(file_data)

            # Apply current filter
            self.filter_files_by_extension(self.file_extension_filter.get())

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.status_var.set("Error scanning files")

    def search_files(self):
        """Search files by name"""
        search_text = self.search_term.get().lower()
        if not search_text:
            # If search is empty, show all files based on the current filter
            self.filter_files_by_extension(self.file_extension_filter.get())
            return

        # Clear listbox
        self.files_listbox.delete(0, END)

        # Get filtered files first (based on extension)
        filtered_files = self.file_list_data.copy()

        # Filter and display matching files
        matching_files = []
        for file_data in filtered_files:
            if search_text in file_data["name"].lower():
                matching_files.append(file_data)
                self.files_listbox.insert(END, file_data["name"])

        # Update the file_list_data to only include search matches
        self.file_list_data = matching_files

        # Update status message with both filter and search info
        filter_text = self.file_extension_filter.get()
        if filter_text == "Custom...":
            filter_text = self.custom_extension.get()
        self.status_var.set(
            f"Found {len(matching_files)} matches for '{search_text}' with filter '{filter_text}'")

    def on_file_select(self, event):
        """Handle file selection from listbox"""
        if not self.files_listbox.curselection():
            return

        # Get selected file data
        selected_index = self.files_listbox.curselection()[0]
        selected_filename = self.files_listbox.get(selected_index)

        # Find matching file data
        for file_data in self.file_list_data:
            if file_data["name"] == selected_filename:
                self.selected_file = file_data
                break

        # Enable manual rename button when a file is selected
        self.manual_rename_button.config(state=tk.NORMAL)

        # Populate manual rename field with current filename without extension
        filename_without_ext = os.path.splitext(selected_filename)[0]
        self.manual_filename.set(filename_without_ext)

        # Find matching entry in Excel
        if self.excel_data is not None and self.name_column.get():
            # Get filename without extension for matching
            filename_without_ext = os.path.splitext(selected_filename)[0]

            # Try different matching strategies
            match_found = False
            matched_item = None

            # Clear previous selection
            for item in self.excel_tree.selection():
                self.excel_tree.selection_remove(item)

            # Strategy 1: Exact match
            for i, item in enumerate(self.excel_tree.get_children()):
                values = self.excel_tree.item(item)["values"]
                col_index = list(self.excel_data.columns).index(
                    self.name_column.get())
                excel_value = str(values[col_index])

                if excel_value == filename_without_ext or excel_value == selected_filename:
                    self.excel_tree.selection_set(item)
                    self.excel_tree.see(item)
                    match_found = True
                    matched_item = item
                    self.update_details(item)
                    break

            # Strategy 2: Contains match (if no exact match found)
            if not match_found:
                for i, item in enumerate(self.excel_tree.get_children()):
                    values = self.excel_tree.item(item)["values"]
                    col_index = list(self.excel_data.columns).index(
                        self.name_column.get())
                    excel_value = str(values[col_index])

                    if filename_without_ext in excel_value or excel_value in filename_without_ext:
                        self.excel_tree.selection_set(item)
                        self.excel_tree.see(item)
                        match_found = True
                        matched_item = item
                        self.update_details(item)
                        break

            if match_found:
                self.status_var.set(f"Found match for {selected_filename}")
                self.rename_button.config(state=tk.NORMAL)

                # Auto-apply pattern if enabled and we have a pattern
                if self.auto_apply_pattern.get() and self.pattern_var.get() and matched_item:
                    custom_name = self.generate_filename_from_pattern(
                        matched_item)
                    if custom_name and self.selected_file:  # Check if selected_file is not None
                        self.manual_filename.set(custom_name)
                        self.status_var.set(
                            f"Auto-applied pattern: {custom_name}")

                        # Also update the details to show the pattern-based filename
                        if self.selected_file:  # Double-check selected_file before accessing it
                            extension = self.selected_file["extension"]
                            if self.keep_extension.get():
                                preview = f"{custom_name}{extension}"
                            else:
                                preview = custom_name

                            # Add preview to details
                            self.details_label.config(
                                text=self.details_label.cget("text") + f"\n\nPattern-based filename: {preview}")
            else:
                self.status_var.set(f"No match found for {selected_filename}")
                self.details_label.config(
                    text=f"No match found in Excel data for {selected_filename}")
                self.rename_button.config(state=tk.DISABLED)
        else:
            self.status_var.set(
                "Excel data not loaded or columns not configured")

    def update_details(self, item):
        """Update details label with information about the selected file and Excel match"""
        if not self.selected_file:
            return

        # Get Excel row data
        values = self.excel_tree.item(item)["values"]

        # Make sure excel_data is not None before accessing columns
        if self.excel_data is None:
            columns = []
        else:
            columns = list(self.excel_data.columns)

        excel_data_dict = {columns[i]: values[i] for i in range(len(columns))}

        # Format file details
        file_details = (
            f"File: {self.selected_file['name']}\n"
            f"Size: {self.format_size(self.selected_file['size'])}\n"
            f"Modified: {self.selected_file['mod_time_formatted']}\n\n"
        )

        # Format Excel match details
        excel_details = "Excel Match:\n"
        for col, val in excel_data_dict.items():
            excel_details += f"{col}: {val}\n"

        # Show proposed new filename
        id_val = excel_data_dict.get(self.id_column.get(), "")
        if id_val:
            extension = self.selected_file["extension"]
            proposed_name = f"{id_val}{extension}"
            excel_details += f"\nProposed new filename: {proposed_name}"

        # Update details label
        self.details_label.config(text=file_details + excel_details)

    def format_size(self, size_bytes):
        """Format file size in human-readable format"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.2f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.2f} TB"

    def rename_file(self):
        """Rename the selected file based on Excel data"""
        if not self.selected_file or not self.excel_tree.selection():
            messagebox.showerror("Error", "No file or Excel entry selected")
            return

        # Get Excel row data
        selected_item = self.excel_tree.selection()[0]
        values = self.excel_tree.item(selected_item)["values"]
        columns = list(self.excel_data.columns)  # type: ignore
        excel_data_dict = {columns[i]: values[i] for i in range(len(columns))}

        # Get ID value for new name
        id_val = excel_data_dict.get(self.id_column.get(), "")
        if not id_val:
            messagebox.showerror(
                "Error", f"No ID value found in the '{self.id_column.get()}' column")
            return

        # Format ID value to ensure it's a valid filename
        id_val = str(id_val).strip()
        # Replace invalid filename chars
        id_val = re.sub(r'[\\/*?:"<>|]', '_', id_val)

        # Create new filename
        extension = self.selected_file["extension"]
        new_filename = f"{id_val}{extension}"

        # Check if source and new names are the same
        if new_filename == self.selected_file["name"]:
            messagebox.showinfo(
                "Info", "The file already has the correct name")
            return

        # Create full paths
        source_path = self.selected_file["path"]
        dest_folder = self.source_folder.get()  # Same folder for now
        dest_path = Path(dest_folder) / new_filename

        # Check if destination file already exists
        if dest_path.exists():
            confirm = messagebox.askyesno(
                "Confirm Overwrite",
                f"File {new_filename} already exists. Overwrite?")
            if not confirm:
                return

        try:
            # Rename file
            os.rename(source_path, dest_path)

            # Update file data
            self.selected_file["name"] = new_filename
            self.selected_file["path"] = dest_path

            # Update listbox
            selected_index = self.files_listbox.curselection()[0]
            self.files_listbox.delete(selected_index)
            self.files_listbox.insert(selected_index, new_filename)
            self.files_listbox.selection_set(selected_index)

            # Update status
            self.status_var.set(f"Renamed file to {new_filename}")
            messagebox.showinfo("Success", f"Renamed file to {new_filename}")

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

        # Validate filename (no invalid characters)
        if re.search(r'[\\/*?:"<>|]', custom_name):
            messagebox.showerror(
                "Error", "Filename contains invalid characters (\\/*?:\"<>|)")
            return

        # Add extension if needed
        extension = self.selected_file["extension"]
        if self.keep_extension.get() and not custom_name.lower().endswith(extension.lower()):
            new_filename = f"{custom_name}{extension}"
        else:
            new_filename = custom_name

        # Create full paths
        source_path = self.selected_file["path"]
        dest_folder = self.source_folder.get()
        dest_path = Path(dest_folder) / new_filename

        # Check if destination file already exists
        if dest_path.exists():
            confirm = messagebox.askyesno(
                "Confirm Overwrite",
                f"File {new_filename} already exists. Overwrite?")
            if not confirm:
                return

        try:
            # Rename file
            os.rename(source_path, dest_path)

            # Update file data
            self.selected_file["name"] = new_filename
            self.selected_file["path"] = dest_path

            # Update listbox
            selected_index = self.files_listbox.curselection()[0]
            self.files_listbox.delete(selected_index)
            self.files_listbox.insert(selected_index, new_filename)
            self.files_listbox.selection_set(selected_index)

            # Clear the manual filename entry
            self.manual_filename.set("")

            # Update status
            self.status_var.set(f"Manually renamed file to {new_filename}")
            messagebox.showinfo("Success", f"Renamed file to {new_filename}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to rename file: {str(e)}")
            self.status_var.set("Error renaming file")

    def add_column_to_pattern(self):
        """Add the selected column to the pattern"""
        if not self.available_columns.get():
            messagebox.showinfo("Info", "Please select a column to add")
            return

        column_name = self.available_columns.get()

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

    def select_common_pattern(self, event):
        """Select a common pattern from the dropdown"""
        pattern = self.common_patterns_combo.get()  # type: ignore
        if pattern:
            self.pattern_var.set(pattern)

    def generate_filename_from_pattern(self, excel_item):
        """Generate a filename based on the current pattern and Excel data"""
        if not self.pattern_var.get():
            return None

        # Get Excel row data
        values = self.excel_tree.item(excel_item)["values"]

        # Make sure excel_data is not None before accessing columns
        if self.excel_data is None:
            return None

        columns = list(self.excel_data.columns)
        excel_data_dict = {columns[i]: values[i] for i in range(len(columns))}

        # Parse the pattern
        pattern_parts = self.pattern_var.get().split(self.separator_var.get())

        # Build the filename from pattern parts
        filename_parts = []
        for part in pattern_parts:
            if part in excel_data_dict:
                # Clean the value - remove special characters
                value = str(excel_data_dict[part])
                # Replace invalid filename chars
                value = re.sub(r'[\\/*?:"<>|]', '_', value)
                filename_parts.append(value)
            else:
                self.status_var.set(
                    f"Warning: Column '{part}' not found in Excel data")
                return None

        # Join the parts with the separator
        custom_name = self.separator_var.get().join(filename_parts)

        # Handle empty result
        if not custom_name:
            return None

        return custom_name

    def apply_pattern(self):
        """Apply the current pattern to create a filename"""
        if not self.selected_file or not self.excel_tree.selection():
            messagebox.showerror("Error", "No file or Excel entry selected")
            return

        if not self.pattern_var.get():
            messagebox.showerror(
                "Error", "Pattern is empty. Please add columns to your pattern.")
            return

        # Get selected Excel item
        selected_item = self.excel_tree.selection()[0]

        # Generate filename from pattern
        custom_name = self.generate_filename_from_pattern(selected_item)
        if not custom_name:
            messagebox.showerror(
                "Error", "Failed to generate filename from pattern")
            return

        # Update the manual filename field
        self.manual_filename.set(custom_name)

        # Preview the result
        extension = self.selected_file["extension"]
        if self.keep_extension.get():
            preview = f"{custom_name}{extension}"
        else:
            preview = custom_name

        messagebox.showinfo("Pattern Applied",
                            f"Pattern applied successfully!\n\nProposed filename: {preview}\n\n"
                            f"Click 'Apply Manual Rename' to rename the file.")

    def update_available_columns(self):
        """Update available columns dropdown based on loaded Excel data"""
        if self.excel_data is not None:
            columns = list(self.excel_data.columns)
            self.available_columns["values"] = columns
            if columns:
                self.available_columns.current(0)


def main():
    root = tk.Tk()
    app = ExcelFileRenamerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
